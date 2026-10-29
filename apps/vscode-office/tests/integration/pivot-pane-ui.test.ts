import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import JSZip from 'jszip'
import { suite, test } from 'mocha'
import { chromium, type Browser, type Frame } from '@playwright/test'
import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'

type Api = { getDocument(uri: vscode.Uri): WorkbookDocument | undefined }

async function viewOf(browser: Browser, name: string) {
  let frame: Frame | undefined
  const deadline = Date.now() + 20_000
  while (!frame && Date.now() < deadline) {
    for (const context of browser.contexts())
      for (const page of context.pages()) {
        if (!(await page.title()).includes(name)) continue
        for (const candidate of page.frames())
          if (
            await candidate
              .locator('#ribbon-tab-data')
              .count()
              .catch(() => 0)
          )
            frame = candidate
      }
    if (!frame) await new Promise((resolve) => setTimeout(resolve, 100))
  }
  assert.ok(frame)
  return frame
}

async function goTo(frame: Frame, reference: string) {
  await vscode.commands.executeCommand('genoffice.workbook.goto')
  const go = frame.getByRole('dialog', { name: 'Go To', exact: true })
  await go.getByRole('textbox', { name: 'Reference', exact: true }).fill(reference)
  await go.getByRole('button', { name: 'Go', exact: true }).click()
  await go.waitFor({ state: 'hidden' })
}

async function settle(document: WorkbookDocument, revision: number) {
  const deadline = Date.now() + 20_000
  while (document.engine.descriptor.revision < revision && Date.now() < deadline)
    await new Promise((resolve) => setTimeout(resolve, 50))
  assert.equal(document.engine.descriptor.revision, revision)
}

async function reopen(api: Api, uri: vscode.Uri, previous: string) {
  await vscode.commands.executeCommand('workbench.action.closeActiveEditor')
  const deadline = Date.now() + 10_000
  while (api.getDocument(uri) && Date.now() < deadline)
    await new Promise((resolve) => setTimeout(resolve, 50))
  await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
  const document = api.getDocument(uri)!
  await document.whenViewReady()
  assert.notEqual(document.sessionId, previous)
  return document
}

suite('Desktop PivotTable Fields pane UI', () => {
  test('relayouts session and file PivotTables through the shared Fields pane with history, deferral and reopen', async function () {
    const port = process.env.GENOFFICE_TEST_CDP_PORT
    if (!port) return this.skip()
    this.timeout(180_000)
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as Api
    const name = 'pivot-pane-ui.xlsx'
    const uri = vscode.Uri.joinPath(vscode.workspace.workspaceFolders![0]!.uri, name)
    const values: (string | number)[][] = [
      ['Region', 'Fruit', 'Amount'],
      ['North', 'Apples', 12.5],
      ['South', 'Pears', 7],
      ['North', 'Pears', 3],
      ['East', 'Apples', 10],
    ]
    const fixture = await applyCellEditsToXlsx(
      await manualWorkbook(),
      values.flatMap((row, r) =>
        row.map((value, c) => ({
          sheetName: 'Other',
          row: r + 2,
          column: c + 2,
          writeValue: true,
          cell: { value },
        })),
      ),
    )
    await vscode.workspace.fs.writeFile(uri, fixture.buffer)
    const browser = await chromium.connectOverCDP(`http://127.0.0.1:${port}`)
    try {
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      let document = api.getDocument(uri)!
      await document.whenViewReady()
      let frame = await viewOf(browser, name)
      await vscode.commands.executeCommand('notifications.clearAll')
      await vscode.commands.executeCommand('workbench.action.closeAuxiliaryBar')
      await vscode.commands.executeCommand('workbench.action.closeSidebar')
      await vscode.commands.executeCommand('workbench.action.closePanel')
      await frame.page().setViewportSize({ width: 1200, height: 900 })
      await goTo(frame, 'Other!C3:E7')
      await frame.locator('#ribbon-tab-insert').click()
      await frame
        .getByRole('button', { name: /^PivotTable\b/ })
        .first()
        .click()
      const dialog = frame.getByRole('dialog', { name: 'Create PivotTable', exact: true })
      await dialog.locator('input:not([readonly])').fill('G3')
      let revision = document.engine.descriptor.revision
      await dialog.getByRole('button', { name: 'Create', exact: true }).click()
      await dialog.waitFor({ state: 'hidden' })
      await settle(document, revision + 1)
      // The Fields pane opens on the new pivot, as on the desktop.
      await goTo(frame, 'Other!G4')
      let pane = frame.locator('.pivot-fields-pane')
      await pane.waitFor()
      const field = (label: string) => pane.getByRole('checkbox', { name: label, exact: true })
      assert.equal(await field('Region').isChecked(), true)
      assert.equal(await field('Amount').isChecked(), true)

      revision = document.engine.descriptor.revision
      await field('Fruit').check()
      await settle(document, revision + 1)
      const nested = document.engine.state.sheets[1]!
      assert.deepEqual(
        (nested.features?.pivotAdditions as { rowFieldIndices: number[] }[])[0]!.rowFieldIndices,
        [0, 1],
      )
      const wide = structuredClone(document.engine.state)

      await field('Region').uncheck()
      await settle(document, revision + 2)
      let cells = document.engine.state.sheets[1]!.cells
      assert.equal(cells.G3!.rawValue, 'Fruit')
      assert.equal(cells.H3!.rawValue, 'Sum of Amount')
      assert.equal(cells.G6!.rawValue, 'Grand Total')
      assert.equal(cells.H6!.rawValue, 32.5)
      assert.equal(cells.I3?.rawValue ?? null, null)
      assert.equal(cells.G9?.rawValue ?? null, null)
      if (await pane.getByRole('alert').count())
        assert.fail(await pane.getByRole('alert').innerText())

      // Deferred updates wait for Update.
      const narrow = structuredClone(document.engine.state)
      await pane.getByRole('checkbox', { name: 'Defer Layout Update' }).check()
      await field('Region').check()
      await new Promise((resolve) => setTimeout(resolve, 300))
      assert.equal(document.engine.descriptor.revision, revision + 2)
      await pane.getByRole('button', { name: 'Update', exact: true }).click()
      await settle(document, revision + 3)
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, narrow)
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, wide)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, narrow)
      await pane.getByRole('checkbox', { name: 'Defer Layout Update' }).uncheck()

      await vscode.commands.executeCommand('workbench.action.files.save')
      document = await reopen(api, uri, document.sessionId)
      frame = await viewOf(browser, name)
      assert.deepEqual(
        (document.engine.state.sheets[1]!.features?.originalMetadata as { pivotRanges: unknown })
          .pivotRanges,
        [{ startRow: 2, startColumn: 6, endRow: 5, endColumn: 7 }],
      )
      // The pane follows the cursor into the file pivot.
      await goTo(frame, 'Other!G4')
      pane = frame.locator('.pivot-fields-pane')
      await pane.waitFor()
      assert.equal(await field('Fruit').isChecked(), true)
      assert.equal(await field('Region').isChecked(), false)
      revision = document.engine.descriptor.revision
      await field('Fruit').uncheck()
      await pane.getByRole('alert').waitFor()
      assert.equal(document.engine.descriptor.revision, revision)
      await field('Region').check()
      await settle(document, revision + 1)
      cells = document.engine.state.sheets[1]!.cells
      assert.equal(cells.G3!.rawValue, 'Region')
      assert.equal(cells.G7!.rawValue, 'Grand Total')
      assert.equal(cells.H7!.rawValue, 32.5)
      await vscode.commands.executeCommand('workbench.action.files.save')
      const zip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const table = Object.keys(zip.files).find((part) =>
        /^xl\/pivotTables\/pivotTable\d+\.xml$/.test(part),
      )!
      assert.match(await zip.file(table)!.async('string'), /<location\b[^>]*ref="G3:H7"/)
      document = await reopen(api, uri, document.sessionId)
      assert.deepEqual(
        (document.engine.state.sheets[1]!.features?.originalMetadata as { pivotRanges: unknown })
          .pivotRanges,
        [{ startRow: 2, startColumn: 6, endRow: 6, endColumn: 7 }],
      )
    } finally {
      if (api.getDocument(uri))
        await vscode.commands.executeCommand('workbench.action.files.save', uri)
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
      await browser.close()
    }
  })

  test('inserts a PivotChart beside a reopened file PivotTable through the shared Insert menu', async function () {
    const port = process.env.GENOFFICE_TEST_CDP_PORT
    if (!port) return this.skip()
    this.timeout(180_000)
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as Api
    const name = 'pivot-chart-ui.xlsx'
    const uri = vscode.Uri.joinPath(vscode.workspace.workspaceFolders![0]!.uri, name)
    const values: (string | number)[][] = [
      ['Region', 'Fruit', 'Amount'],
      ['North', 'Apples', 12.5],
      ['South', 'Pears', 7],
      ['North', 'Pears', 3],
      ['East', 'Apples', 10],
    ]
    const fixture = await applyCellEditsToXlsx(
      await manualWorkbook(),
      values.flatMap((row, r) =>
        row.map((value, c) => ({
          sheetName: 'Other',
          row: r + 2,
          column: c + 2,
          writeValue: true,
          cell: { value },
        })),
      ),
    )
    await vscode.workspace.fs.writeFile(uri, fixture.buffer)
    const browser = await chromium.connectOverCDP(`http://127.0.0.1:${port}`)
    try {
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      let document = api.getDocument(uri)!
      await document.whenViewReady()
      let frame = await viewOf(browser, name)
      await vscode.commands.executeCommand('notifications.clearAll')
      await vscode.commands.executeCommand('workbench.action.closeAuxiliaryBar')
      await vscode.commands.executeCommand('workbench.action.closeSidebar')
      await vscode.commands.executeCommand('workbench.action.closePanel')
      await frame.page().setViewportSize({ width: 1200, height: 900 })
      await goTo(frame, 'Other!C3:E7')
      await frame.locator('#ribbon-tab-insert').click()
      await frame
        .getByRole('button', { name: /^PivotTable\b/ })
        .first()
        .click()
      const dialog = frame.getByRole('dialog', { name: 'Create PivotTable', exact: true })
      await dialog.locator('input:not([readonly])').fill('G3')
      const created = document.engine.descriptor.revision
      await dialog.getByRole('button', { name: 'Create', exact: true }).click()
      await settle(document, created + 1)
      await vscode.commands.executeCommand('workbench.action.files.save')
      document = await reopen(api, uri, document.sessionId)
      frame = await viewOf(browser, name)
      await goTo(frame, 'Other!G4')
      await frame.locator('#ribbon-tab-insert').click()
      const revision = document.engine.descriptor.revision
      await frame
        .locator('.workbook-desktop-header')
        .getByRole('button', { name: 'PivotChart', exact: true })
        .click()
      await frame.getByRole('listbox').getByRole('option', { name: 'Column', exact: true }).click()
      await settle(document, revision + 1)
      if (await frame.getByRole('alert').count())
        assert.fail(await frame.getByRole('alert').first().innerText())
      const visuals = document.engine.state.sheets[1]!.features?.visuals as {
        chart: { series: { valuesRef?: string }[] }
      }[]
      assert.equal(visuals.length, 1)
      assert.match(String(visuals[0]!.chart.series[0]!.valuesRef), /\$H\$4:\$H\$6$/)
      await vscode.commands.executeCommand('workbench.action.files.save')
      const zip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      assert.ok(Object.keys(zip.files).some((part) => /^xl\/charts\/chart\d+\.xml$/.test(part)))
    } finally {
      if (api.getDocument(uri))
        await vscode.commands.executeCommand('workbench.action.files.save', uri)
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
      await browser.close()
    }
  })

  test('filters a reopened file PivotTable through the shared Insert → Slicer panel', async function () {
    const port = process.env.GENOFFICE_TEST_CDP_PORT
    if (!port) return this.skip()
    this.timeout(180_000)
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as Api
    const name = 'pivot-slicer-ui.xlsx'
    const uri = vscode.Uri.joinPath(vscode.workspace.workspaceFolders![0]!.uri, name)
    const values: (string | number)[][] = [
      ['Region', 'Fruit', 'Amount'],
      ['North', 'Apples', 12.5],
      ['South', 'Pears', 7],
      ['North', 'Pears', 3],
      ['East', 'Apples', 10],
    ]
    const fixture = await applyCellEditsToXlsx(
      await manualWorkbook(),
      values.flatMap((row, r) =>
        row.map((value, c) => ({
          sheetName: 'Other',
          row: r + 2,
          column: c + 2,
          writeValue: true,
          cell: { value },
        })),
      ),
    )
    await vscode.workspace.fs.writeFile(uri, fixture.buffer)
    const browser = await chromium.connectOverCDP(`http://127.0.0.1:${port}`)
    try {
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      let document = api.getDocument(uri)!
      await document.whenViewReady()
      let frame = await viewOf(browser, name)
      await vscode.commands.executeCommand('notifications.clearAll')
      await vscode.commands.executeCommand('workbench.action.closeAuxiliaryBar')
      await vscode.commands.executeCommand('workbench.action.closeSidebar')
      await vscode.commands.executeCommand('workbench.action.closePanel')
      await frame.page().setViewportSize({ width: 1200, height: 900 })
      await goTo(frame, 'Other!C3:E7')
      await frame.locator('#ribbon-tab-insert').click()
      await frame
        .getByRole('button', { name: /^PivotTable\b/ })
        .first()
        .click()
      const dialog = frame.getByRole('dialog', { name: 'Create PivotTable', exact: true })
      await dialog.locator('input:not([readonly])').fill('G3')
      const created = document.engine.descriptor.revision
      await dialog.getByRole('button', { name: 'Create', exact: true }).click()
      await settle(document, created + 1)
      // Session pivots are written on save before slicers can bind, as on the desktop.
      await goTo(frame, 'Other!G4')
      await frame.locator('#ribbon-tab-insert').click()
      const header = frame.locator('.workbook-desktop-header')
      await header.getByRole('button', { name: 'Slicer', exact: true }).click()
      await frame.getByRole('alert').filter({ hasText: /save/i }).first().waitFor()
      await vscode.commands.executeCommand('workbench.action.files.save')
      document = await reopen(api, uri, document.sessionId)
      frame = await viewOf(browser, name)
      await goTo(frame, 'Other!G4')
      await frame.locator('#ribbon-tab-insert').click()
      // No date field: the timeline picker refuses, as on the desktop.
      await frame
        .locator('.workbook-desktop-header')
        .getByRole('button', { name: 'Timeline', exact: true })
        .click()
      await frame.getByRole('alert').first().waitFor()
      await frame
        .locator('.workbook-desktop-header')
        .getByRole('button', { name: 'Slicer', exact: true })
        .click()
      const picker = frame.getByRole('dialog', { name: 'Insert Slicer', exact: true })
      await picker.getByRole('button', { name: 'Region', exact: true }).click()
      await picker.waitFor({ state: 'hidden' })
      const panel = frame.getByRole('region', { name: 'Slicer Region', exact: true })
      await panel.waitFor()
      let revision = document.engine.descriptor.revision
      const unfiltered = structuredClone(document.engine.state)
      await panel.getByRole('button', { name: 'North', exact: true }).click()
      await settle(document, revision + 1)
      let cells = document.engine.state.sheets[1]!.cells
      assert.deepEqual(
        ['G4', 'G5', 'G6'].map((address) => cells[address]?.rawValue),
        ['South', 'East', 'Grand Total'],
      )
      assert.equal(cells.H6!.rawValue, 17)
      assert.equal(cells.G7?.rawValue ?? null, null)
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, unfiltered)
      await vscode.commands.executeCommand('redo')
      await vscode.commands.executeCommand('workbench.action.files.save')
      const zip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const table = Object.keys(zip.files).find((part) =>
        /^xl\/pivotTables\/pivotTable\d+\.xml$/.test(part),
      )!
      assert.match(await zip.file(table)!.async('string'), /<location\b[^>]*ref="G3:H6"/)
      revision = document.engine.descriptor.revision
      await panel.getByRole('button', { name: 'Remove slicer (restore all members)' }).click()
      await settle(document, revision + 1)
      await panel.waitFor({ state: 'hidden' })
      cells = document.engine.state.sheets[1]!.cells
      assert.equal(cells.G7!.rawValue, 'Grand Total')
      assert.equal(cells.H7!.rawValue, 32.5)
    } finally {
      if (api.getDocument(uri))
        await vscode.commands.executeCommand('workbench.action.files.save', uri)
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
      await browser.close()
    }
  })
})
