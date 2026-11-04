import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import JSZip from 'jszip'
import { suite, test } from 'mocha'
import { chromium, type Browser, type Frame } from '@playwright/test'
import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'
import { gotoReference } from './goto-helper'

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
  await gotoReference(frame, reference)
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

suite('Desktop Tables UI', () => {
  test('creates a table and edits it through the shared Table Design tab with history, save and reopen', async function () {
    const port = process.env.GENOFFICE_TEST_CDP_PORT
    if (!port) return this.skip()
    this.timeout(180_000)
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as Api
    const name = 'tables-ui.xlsx'
    const uri = vscode.Uri.joinPath(vscode.workspace.workspaceFolders![0]!.uri, name)
    const values: (string | number)[][] = [
      ['Region', 'Amount'],
      ['North', 12.5],
      ['South', 7],
      ['North', 3],
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
      await goTo(frame, 'Other!C3:D6')
      await frame.locator('#ribbon-tab-insert').click()
      let header = frame.locator('.workbook-desktop-header')
      const before = structuredClone(document.engine.state)
      let revision = document.engine.descriptor.revision
      await header
        .getByRole('button', { name: /^Table\b/ })
        .first()
        .click()
      await settle(document, ++revision)
      const tables = () =>
        document.engine.state.sheets[1]!.features?.tableAdditions as {
          name: string
          area: object
          bandedColumns?: boolean
          totalsRow?: boolean
        }[]
      assert.equal(tables()[0]!.name, 'Table1')
      await frame.locator('#ribbon-tab-tableDesign').click()
      header = frame.locator('.workbook-desktop-header')
      await header.getByRole('button', { name: /Banded Columns/ }).click()
      await settle(document, ++revision)
      assert.equal(tables()[0]!.bandedColumns, true)
      await header.getByRole('button', { name: /Total Row/ }).click()
      await settle(document, ++revision)
      assert.equal(tables()[0]!.totalsRow, true)
      const cells = document.engine.state.sheets[1]!.cells
      assert.equal(cells.C7!.rawValue, 'Total')
      assert.match(String(cells.D7!.formula), /SUBTOTAL\(109,D4:D6\)/)
      const rename = header.locator('.table-name-field input')
      await rename.fill('Sales')
      await rename.press('Enter')
      await settle(document, ++revision)
      assert.equal(tables()[0]!.name, 'Sales')
      if (await frame.getByRole('alert').count())
        assert.fail(await frame.getByRole('alert').first().innerText())
      const after = structuredClone(document.engine.state)
      for (let step = 0; step < 4; step++) await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, before)
      for (let step = 0; step < 4; step++) await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, after)
      await vscode.commands.executeCommand('workbench.action.files.save')
      const zip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const part = Object.keys(zip.files).find((path) => /^xl\/tables\/table\d+\.xml$/.test(path))!
      const xml = await zip.file(part)!.async('string')
      assert.match(xml, /displayName="Sales"/)
      assert.match(xml, /ref="C3:D7"/)
      assert.match(xml, /totalsRowCount="1"/)
      assert.match(xml, /showColumnStripes="1"/)
      document = await reopen(api, uri, document.sessionId)
      const file = (
        document.engine.state.sheets[1]!.features?.originalMetadata as {
          tables: { name: string; totalsRowCount?: number }[]
        }
      ).tables
      assert.equal(file[0]!.name, 'Sales')
      frame = await viewOf(browser, name)
      await goTo(frame, 'Other!C4')
      await frame.locator('#ribbon-tab-tableDesign').click()
      revision = document.engine.descriptor.revision
      await frame
        .locator('.workbook-desktop-header')
        .getByRole('button', { name: /Convert to Range/ })
        .click()
      await settle(document, revision + 1)
      const edits = document.engine.state.sheets[1]!.features?.tableEdits as Record<
        string,
        { remove?: boolean }
      >
      assert.equal(edits.Sales!.remove, true)
    } finally {
      if (api.getDocument(uri))
        await vscode.commands.executeCommand('workbench.action.files.save', uri)
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
      await browser.close()
    }
  })
})
