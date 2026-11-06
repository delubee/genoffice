import { workbookFileSchema } from '@genoffice/xlsx-gateway/shared/workbook-protocol'
import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import JSZip from 'jszip'
import { suite, test } from 'mocha'
import { chromium, expect, type Frame } from '@playwright/test'
import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'
import { gotoReference } from './goto-helper'

suite('Desktop Subtotal UI', () => {
  test('creates offset subtotals, outline and grand total in one history entry and preserves references after reopen', async function () {
    const port = process.env.GENOFFICE_TEST_CDP_PORT
    if (!port) return this.skip()
    this.timeout(90_000)
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as {
      getDocument(uri: vscode.Uri): WorkbookDocument | undefined
    }
    const uri = vscode.Uri.joinPath(vscode.workspace.workspaceFolders![0]!.uri, 'subtotal-ui.xlsx')
    const values = [
      ['Region', 'Amount'],
      ['North', 12.5],
      ['North', 7],
      ['South', 3],
    ]
    const fixture = await applyCellEditsToXlsx(await manualWorkbook(), [
      ...values.flatMap((row, r) =>
        row.map((value, c) => ({
          sheetName: 'Data',
          row: r + 2,
          column: c + 1,
          writeValue: true,
          cell: { value },
          ...(r === 1 && c === 1 ? { style: { fillColor: '#CCEEFF', numberFormat: '0.00' } } : {}),
        })),
      ),
      { sheetName: 'Data', row: 7, column: 0, writeValue: true, cell: { value: 42 } },
      {
        sheetName: 'Other',
        row: 0,
        column: 0,
        writeValue: true,
        cell: { value: 42, formula: '=Data!A8' },
      },
    ])
    await vscode.workspace.fs.writeFile(uri, fixture.buffer)
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
            if (!(await page.title()).includes('subtotal-ui.xlsx')) continue
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
      await vscode.commands.executeCommand('notifications.clearAll')
      await vscode.commands.executeCommand('workbench.action.closeAuxiliaryBar')
      await vscode.commands.executeCommand('workbench.action.closeSidebar')
      await vscode.commands.executeCommand('workbench.action.closePanel')
      await frame.page().setViewportSize({ width: 1200, height: 900 })
      await gotoReference(frame, 'Data!B3:C6')
      const before = structuredClone(document.engine.state)
      const revision = document.engine.descriptor.revision
      await frame.locator('#ribbon-tab-data').click()
      await frame.getByRole('button', { name: 'Subtotal', exact: true }).click()
      const dialog = frame.getByRole('dialog', { name: 'Subtotal', exact: true })
      await dialog.getByRole('button', { name: 'OK', exact: true }).click()
      await document.bridge!.request('view.flush', {})
      if (await frame.getByRole('alert').count())
        assert.fail(await frame.getByRole('alert').first().innerText())
      assert.equal(document.engine.descriptor.revision, revision + 1)
      const after = structuredClone(document.engine.state)
      assert.equal(after.structuralLog.length - before.structuralLog.length, 3)
      assert.equal(after.sheets[0]!.cells.B6!.rawValue, 'North Total')
      assert.equal(after.sheets[0]!.cells.C6!.formula, '=SUBTOTAL(9,$C$4:$C$5)')
      assert.equal(after.sheets[0]!.cells.C8!.formula, '=SUBTOTAL(9,$C$7:$C$7)')
      assert.equal(after.sheets[0]!.cells.C9!.formula, '=SUBTOTAL(9,$C$4:$C$8)')
      assert.equal(after.sheets[0]!.cells.A11!.rawValue, 42)
      assert.equal(after.sheets[1]!.cells.A1!.formula, '=Data!A11')
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, before)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, after)
      await vscode.commands.executeCommand('workbench.action.files.save')
      const zip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      assert.match(
        await zip.file('xl/worksheets/sheet1.xml')!.async('string'),
        /<f>SUBTOTAL\(9,\$C\$4:\$C\$5\)<\/f>/,
      )
      const savedXml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
      assert.match(savedXml, /<row\b[^>]*r="4"[^>]*outlineLevel="2"/)
      assert.match(savedXml, /<row\b[^>]*r="6"[^>]*outlineLevel="1"/)
      assert.doesNotMatch(savedXml, /summaryBelow="0"/)
      assert.equal(
        await zip.file('customXml/manual-sentinel.xml')!.async('string'),
        '<preserve id="manual-fixture">unchanged</preserve>',
      )
      const previousSession = document.sessionId
      await vscode.commands.executeCommand('workbench.action.closeActiveEditor')
      const closeDeadline = Date.now() + 10_000
      while (api.getDocument(uri) && Date.now() < closeDeadline)
        await new Promise((resolve) => setTimeout(resolve, 50))
      assert.equal(api.getDocument(uri), undefined)
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      const reopened = api.getDocument(uri)!
      await reopened.whenViewReady()
      assert.notEqual(reopened.sessionId, previousSession)
      const metadata = workbookFileSchema.parse(reopened.metadata)
      assert.equal(metadata.sheets[0]!.outlineLevelRow, 2)
      assert.notEqual(metadata.sheets[0]!.outlineSummaryBelow, false)
      let importedFrame: Frame | undefined
      const reopenDeadline = Date.now() + 20_000
      while (!importedFrame && Date.now() < reopenDeadline) {
        for (const context of browser.contexts())
          for (const page of context.pages()) {
            if (!(await page.title()).includes('subtotal-ui.xlsx')) continue
            for (const candidate of page.frames())
              if (
                await candidate
                  .locator('#ribbon-tab-data')
                  .count()
                  .catch(() => 0)
              )
                importedFrame = candidate
          }
        if (!importedFrame) await new Promise((resolve) => setTimeout(resolve, 100))
      }
      assert.ok(importedFrame)
      await expect(importedFrame.locator('.outline-gutter-rows')).toBeVisible()
      await importedFrame.locator('#ribbon-tab-data').click()
      await importedFrame.getByRole('button', { name: 'Group', exact: true }).click()
      await expect(
        importedFrame.getByRole('option', { name: '✓ Summary rows below detail', exact: true }),
      ).toBeVisible()
      await importedFrame.getByRole('button', { name: 'Group', exact: true }).press('Escape')
      const snapshot = reopened.snapshot()
      try {
        const result = await snapshot.read(reopened.engine.state.sheets[0]!.id, 'A1:D11')
        const cells = new Map(result.cells.map((cell) => [cell.address, cell]))
        assert.equal(cells.get('B6')!.rawValue, 'North Total')
        assert.equal(cells.get('C6')!.formula, '=SUBTOTAL(9,$C$4:$C$5)')
        assert.equal(cells.get('C8')!.formula, '=SUBTOTAL(9,$C$7:$C$7)')
        assert.equal(cells.get('C9')!.formula, '=SUBTOTAL(9,$C$4:$C$8)')
        assert.equal(cells.get('C4')!.style?.fillColor, '#CCEEFF')
        assert.equal(cells.get('C4')!.numberFormat, '0.00')
        assert.equal(cells.get('B9')!.rawValue, 'Grand Total')
        assert.equal(cells.get('A11')!.rawValue, 42)
        assert.equal(cells.get('A2')!.rawValue, 12.5)
        const other = await snapshot.read(reopened.engine.state.sheets[1]!.id, 'A1')
        assert.equal(other.cells[0]!.formula, '=Data!A11')
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
