import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import { suite, test } from 'mocha'
import JSZip from 'jszip'
import { chromium, type Frame } from '@playwright/test'
import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'
import { gotoReference } from './goto-helper'

suite('Desktop Custom Sort large selections', () => {
  test('sorts a selection beyond 2000 cells as one revision', async function () {
    const port = process.env.GENOFFICE_TEST_CDP_PORT
    if (!port) return this.skip()
    this.timeout(180_000)
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as {
      getDocument(uri: vscode.Uri): WorkbookDocument | undefined
    }
    const uri = vscode.Uri.joinPath(
      vscode.workspace.workspaceFolders![0]!.uri,
      'custom-sort-large-ui.xlsx',
    )
    // A header plus 1200 rows in descending order: 1201 x 3 = 3603 cells.
    const rows: (string | number)[][] = [['Rank', 'Label', 'Group']]
    for (let index = 1200; index >= 1; index--)
      rows.push([index, `Label ${index}`, `G${index % 5}`])
    const fixture = await applyCellEditsToXlsx(
      await manualWorkbook(),
      rows.flatMap((row, r) =>
        row.map((value, c) => ({
          sheetName: 'Data',
          row: r,
          column: c,
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
            if (!(await page.title()).includes('custom-sort-large-ui.xlsx')) continue
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
      await gotoReference(frame, 'Data!A1:C1201')
      const revision = document.engine.descriptor.revision
      await frame.locator('#ribbon-tab-data').click()
      await frame
        .locator('.workbook-shared-sort-filter')
        .getByRole('button', { name: 'Sort', exact: true })
        .click()
      await frame.getByRole('option', { name: 'Custom Sort…', exact: true }).click()
      const dialog = frame.locator('.sort-dialog')
      await dialog.getByRole('button', { name: 'OK', exact: true }).click()
      await Promise.race([
        dialog.waitFor({ state: 'hidden', timeout: 60_000 }),
        dialog.getByRole('alert').waitFor({ timeout: 60_000 }),
      ])
      if (await dialog.getByRole('alert').count())
        assert.fail(await dialog.getByRole('alert').innerText())
      await document.bridge!.request('view.flush', {})
      assert.equal(document.engine.descriptor.revision, revision + 1)
      const sheet = document.engine.state.sheets[0]!
      assert.equal(sheet.cells.A2!.rawValue, 1)
      assert.equal(sheet.cells.B2!.rawValue, 'Label 1')
      assert.equal(sheet.cells.A1201!.rawValue, 1200)
      assert.equal(sheet.cells.B1201!.rawValue, 'Label 1200')
      await vscode.commands.executeCommand('undo')
      // Undo returns to the source file's value, which the session state does not keep as an edit.
      assert.equal(document.engine.state.sheets[0]!.cells.A2?.rawValue ?? 1200, 1200)
      await vscode.commands.executeCommand('redo')
      assert.equal(document.engine.state.sheets[0]!.cells.A2!.rawValue, 1)
      await vscode.commands.executeCommand('workbench.action.files.save')
      const xml = await (
        await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      )
        .file('xl/worksheets/sheet1.xml')!
        .async('string')
      assert.match(xml, /<row\b[^>]*r="2"[^>]*>\s*<c\b[^>]*r="A2"[^>]*>\s*<v>1<\/v>/)
      assert.match(xml, /<row\b[^>]*r="1201"[^>]*>\s*<c\b[^>]*r="A1201"[^>]*>\s*<v>1200<\/v>/)
    } finally {
      if (api.getDocument(uri))
        await vscode.commands.executeCommand('workbench.action.files.save', uri)
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
      await browser.close()
    }
  })
})
