import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import JSZip from 'jszip'
import { suite, test } from 'mocha'
import { chromium, type Frame } from '@playwright/test'
import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'

suite('Desktop PivotTable create UI', () => {
  test('creates a PivotTable through the shared dialog, saves native pivot parts and reopens it', async function () {
    const port = process.env.GENOFFICE_TEST_CDP_PORT
    if (!port) return this.skip()
    this.timeout(120_000)
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as {
      getDocument(uri: vscode.Uri): WorkbookDocument | undefined
    }
    const uri = vscode.Uri.joinPath(
      vscode.workspace.workspaceFolders![0]!.uri,
      'pivot-create-ui.xlsx',
    )
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
      const document = api.getDocument(uri)!
      await document.whenViewReady()
      let frame: Frame | undefined
      const deadline = Date.now() + 20_000
      while (!frame && Date.now() < deadline) {
        for (const context of browser.contexts())
          for (const page of context.pages()) {
            if (!(await page.title()).includes('pivot-create-ui.xlsx')) continue
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
      await vscode.commands.executeCommand('genoffice.workbook.goto')
      const go = frame.getByRole('dialog', { name: 'Go To', exact: true })
      await go.getByRole('textbox', { name: 'Reference', exact: true }).fill('Other!C3:E7')
      await go.getByRole('button', { name: 'Go', exact: true }).click()
      await go.waitFor({ state: 'hidden' })

      // Refresh on a sheet without pivots mirrors the desktop message.
      await frame.locator('#ribbon-tab-data').click()
      await frame
        .getByRole('button', { name: /^Refresh\b/ })
        .first()
        .click()
      await frame
        .getByRole('alert')
        .filter({ hasText: 'The current sheet has no PivotTables.' })
        .waitFor()

      const before = structuredClone(document.engine.state)
      const revision = document.engine.descriptor.revision
      await frame.locator('#ribbon-tab-insert').click()
      await frame
        .getByRole('button', { name: /^PivotTable\b/ })
        .first()
        .click()
      const dialog = frame.getByRole('dialog', { name: 'Create PivotTable', exact: true })
      await dialog.waitFor()
      assert.equal(await dialog.locator('input[readonly]').inputValue(), 'C3:E7')
      // An overlapping destination is refused inside the dialog and keeps the draft.
      await dialog.locator('input:not([readonly])').fill('D4')
      await dialog.getByRole('button', { name: 'Create', exact: true }).click()
      await dialog.getByRole('alert').waitFor()
      assert.equal(document.engine.descriptor.revision, revision)
      await dialog.locator('input:not([readonly])').fill('G3')
      await dialog.getByRole('button', { name: 'Create', exact: true }).click()
      await dialog.waitFor({ state: 'hidden' })
      await document.bridge!.request('view.flush', {})
      if (await frame.getByRole('alert').count())
        assert.fail(await frame.getByRole('alert').first().innerText())
      assert.equal(document.engine.descriptor.revision, revision + 1)
      const after = structuredClone(document.engine.state)
      const other = after.sheets[1]!
      assert.equal(other.cells.G3!.rawValue, 'Region')
      assert.equal(other.cells.H3!.rawValue, 'Sum of Amount')
      assert.equal(other.cells.G7!.rawValue, 'Grand Total')
      assert.equal(other.cells.H7!.rawValue, 32.5)
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, before)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, after)
      await vscode.commands.executeCommand('workbench.action.files.save')
      const zip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      assert.ok(
        Object.keys(zip.files).some((name) => /^xl\/pivotTables\/pivotTable\d+\.xml$/.test(name)),
      )
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
      const sheet = reopened.engine.state.sheets[1]!
      assert.deepEqual(
        (sheet.features?.originalMetadata as { pivotRanges?: unknown[] }).pivotRanges,
        [{ startRow: 2, startColumn: 6, endRow: 6, endColumn: 7 }],
      )
      const snapshot = reopened.snapshot()
      try {
        const result = await snapshot.read(sheet.id, 'G3:H7')
        const saved = new Map(result.cells.map((cell) => [cell.address, cell]))
        assert.equal(saved.get('H7')!.rawValue, 32.5)
        assert.equal(saved.get('G3')!.rawValue, 'Region')
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
