import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import JSZip from 'jszip'
import { suite, test } from 'mocha'
import { chromium, expect, type Frame } from '@playwright/test'
import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'
import { gotoReference } from './goto-helper'

suite('Desktop Filter Popup Sort UI', () => {
  test('sorts visible row slots and preserves source-hidden rows, styles, native history and saved criteria', async function () {
    const port = process.env.GENOFFICE_TEST_CDP_PORT
    if (!port) return this.skip()
    this.timeout(90_000)
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as {
      getDocument(uri: vscode.Uri): WorkbookDocument | undefined
    }
    const uri = vscode.Uri.joinPath(
      vscode.workspace.workspaceFolders![0]!.uri,
      'filter-popup-sort-ui.xlsx',
    )
    const values = [
      ['Group', 'Amount'],
      ['Keep', 30],
      ['Drop', 999],
      ['Keep', 10],
      ['Keep', 5],
      ['Keep', 20],
    ]
    const fixture = await applyCellEditsToXlsx(await manualWorkbook(), [
      ...values.flatMap((row, r) =>
        row.map((value, c) => ({
          sheetName: 'Data',
          row: r,
          column: c,
          writeValue: true,
          cell: { value },
          ...(r === 3 && c === 1 ? { style: { fillColor: '#FF0000' } } : {}),
          ...(r === 5 && c === 1 ? { style: { numberFormat: '0.00' } } : {}),
        })),
      ),
      {
        sheetName: 'Data',
        row: 79,
        column: 0,
        writeValue: true,
        cell: { value: 'Separate far table' },
      },
    ])
    const originalZip = await JSZip.loadAsync(fixture.buffer)
    const originalXml = await originalZip.file('xl/worksheets/sheet1.xml')!.async('string')
    originalZip.file(
      'xl/worksheets/sheet1.xml',
      originalXml.replace(/(<row\b[^>]*\br="5")/, '$1 hidden="1"'),
    )
    await vscode.workspace.fs.writeFile(
      uri,
      await originalZip.generateAsync({ type: 'nodebuffer' }),
    )
    const browser = await chromium.connectOverCDP(`http://127.0.0.1:${port}`)
    try {
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      const document = api.getDocument(uri)!
      await document.whenViewReady()
      let frame: Frame | undefined
      const deadline = Date.now() + 20_000
      while (!frame && Date.now() < deadline) {
        for (const context of browser.contexts())
          for (const page of context.pages()) {
            if (!(await page.title()).includes('filter-popup-sort-ui.xlsx')) continue
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
      await frame.page().setViewportSize({ width: 1200, height: 900 })
      await vscode.commands.executeCommand('notifications.clearAll')
      await vscode.commands.executeCommand('workbench.action.closeAuxiliaryBar')
      await vscode.commands.executeCommand('workbench.action.closeSidebar')
      await vscode.commands.executeCommand('workbench.action.closePanel')
      await gotoReference(frame, 'Data!A1:B6')
      await frame.locator('#ribbon-tab-data').click()
      const band = frame.locator('.workbook-shared-sort-filter')
      await band.getByRole('button', { name: /^Filter\b/ }).click()
      await document.bridge!.request('view.flush', {})
      await document.engine.submit({
        schemaVersion: 1,
        sessionId: document.sessionId,
        transactionId: crypto.randomUUID(),
        expectedRevision: document.engine.descriptor.revision,
        origin: 'agent',
        summary: 'Set known header width',
        operations: [
          {
            op: 'set_col_width',
            sheetId: document.engine.state.sheets[0]!.id,
            column: 'A',
            count: 2,
            widthPx: 140,
          },
        ],
      })
      await document.bridge!.request('view.flush', {})
      const panel = frame
        .locator('[data-u-comp="sheets-filter-panel"]')
        .filter({ has: frame.locator('[data-u-comp="sheets-filter-panel-footer"]') })
      const openHeader = async (column: number) => {
        await expect
          .poll(() =>
            frame!
              .locator('.workbook-grid canvas')
              .evaluateAll((elements) =>
                Math.max(...elements.map((canvas) => canvas.getBoundingClientRect().height)),
              ),
          )
          .toBeGreaterThan(150)
        const index = await frame!.locator('.workbook-grid canvas').evaluateAll(
          (elements) =>
            elements
              .map((canvas, index) => ({
                index,
                area: canvas.getBoundingClientRect().width * canvas.getBoundingClientRect().height,
              }))
              .sort((a, b) => b.area - a.area)[0]!.index,
        )
        const box = (await frame!.locator('.workbook-grid canvas').nth(index).boundingBox())!
        await frame!.page().mouse.click(box.x + 46 + 140 * (column + 1) - 9, box.y + 20 + 20 - 9)
        await expect(panel).toBeVisible()
      }
      await openHeader(0)
      await panel
        .getByText('Drop', { exact: true })
        .locator('xpath=ancestor::div[label[@data-u-comp="checkbox"]][1]')
        .locator('label[data-u-comp="checkbox"]')
        .click()
      await panel.getByRole('button', { name: 'Confirm', exact: true }).click()
      await document.bridge!.request('view.flush', {})
      await expect(panel).toBeHidden()
      const before = structuredClone(document.engine.state)
      const revision = document.engine.descriptor.revision
      assert.deepEqual((before.sheets[0]!.features!.filterEdits as any).hiddenRows, [2])
      assert.deepEqual(before.sheets[0]!.features!.visibilityEdits, [
        { axis: 'rows', start: 4, count: 1, hidden: true },
      ])
      await openHeader(1)
      await panel.getByRole('button', { name: 'Ascending', exact: true }).click()
      await document.bridge!.request('view.flush', {})
      await expect(frame.getByRole('alert')).toHaveCount(0)
      await expect(panel).toBeHidden()
      const after = structuredClone(document.engine.state)
      assert.equal(document.engine.descriptor.revision, revision + 1)
      const sortedSnapshot = document.snapshot()
      try {
        const sorted = await sortedSnapshot.read(after.sheets[0]!.id, 'A1:B6')
        const cells = new Map(sorted.cells.map((cell) => [cell.address, cell]))
        assert.deepEqual(
          ['B2', 'B3', 'B4', 'B5', 'B6'].map((address) => cells.get(address)!.rawValue),
          [10, 999, 20, 5, 30],
        )
      } finally {
        sortedSnapshot.release()
      }
      assert.equal(after.sheets[0]!.cells.B3, before.sheets[0]!.cells.B3)
      assert.equal(after.sheets[0]!.cells.B5, before.sheets[0]!.cells.B5)
      assert.equal(after.sheets[0]!.cells.B2!.style!.fillColor, '#FF0000')
      assert.equal(after.sheets[0]!.cells.B4!.numberFormat, '0.00')
      assert.equal(after.sheets[0]!.cells.A6!.style!.fillColor, '#CCEEFF')
      assert.deepEqual(
        after.sheets[0]!.features!.filterEdits,
        before.sheets[0]!.features!.filterEdits,
      )
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, before)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, after)
      await vscode.commands.executeCommand('workbench.action.files.save')
      const saved = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const xml = await saved.file('xl/worksheets/sheet1.xml')!.async('string')
      assert.match(xml, /<autoFilter ref="A1:B6"/)
      assert.match(xml, /<filter val="Keep"/)
      for (const row of [3, 5])
        assert.match(xml, new RegExp('<row\\b[^>]*r="' + row + '"[^>]*hidden="1"'))
      assert.equal(
        await saved.file('customXml/manual-sentinel.xml')!.async('string'),
        '<preserve id="manual-fixture">unchanged</preserve>',
      )
      const session = document.sessionId
      await vscode.commands.executeCommand('workbench.action.closeActiveEditor')
      const closeDeadline = Date.now() + 10000
      while (api.getDocument(uri) && Date.now() < closeDeadline)
        await new Promise((resolve) => setTimeout(resolve, 50))
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      const reopened = api.getDocument(uri)!
      await reopened.whenViewReady()
      assert.notEqual(reopened.sessionId, session)
      assert.equal(reopened.engine.descriptor.revision, 0)
      assert.equal(reopened.engine.descriptor.dirty, false)
      const snapshot = reopened.snapshot()
      try {
        const result = await snapshot.read(reopened.engine.state.sheets[0]!.id, 'A1:B6')
        const cells = new Map(result.cells.map((cell) => [cell.address, cell]))
        assert.deepEqual(
          ['B2', 'B3', 'B4', 'B5', 'B6'].map((address) => cells.get(address)!.rawValue),
          [10, 999, 20, 5, 30],
        )
        assert.equal(cells.get('B2')!.style!.fillColor, '#FF0000')
        assert.equal(cells.get('B4')!.numberFormat, '0.00')
        assert.equal(cells.get('A6')!.style!.fillColor, '#CCEEFF')
      } finally {
        snapshot.release()
      }
    } finally {
      if (api.getDocument(uri))
        await vscode.commands.executeCommand('workbench.action.files.save', uri)
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
      await browser.close()
    }
  })
})
