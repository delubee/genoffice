import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import { suite, test } from 'mocha'
import { chromium, type Frame } from '@playwright/test'
import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'
import { gotoReference } from './goto-helper'

suite('Desktop Subtotal large selections', () => {
  test('subtotals a source beyond 2000 cells as one revision', async function () {
    const port = process.env.GENOFFICE_TEST_CDP_PORT
    if (!port) return this.skip()
    this.timeout(240_000)
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as {
      getDocument(uri: vscode.Uri): WorkbookDocument | undefined
    }
    const uri = vscode.Uri.joinPath(
      vscode.workspace.workspaceFolders![0]!.uri,
      'subtotal-large-ui.xlsx',
    )
    // A header plus 1500 rows in 30 sorted groups of 35: 1051 x 2 = 2102 cells.
    const rows: (string | number)[][] = [['Group', 'Amount']]
    for (let index = 0; index < 1050; index++)
      rows.push([`G${String(Math.floor(index / 35)).padStart(3, '0')}`, 1])
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
            if (!(await page.title()).includes('subtotal-large-ui.xlsx')) continue
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
      await gotoReference(frame, 'Data!A1:B1051')
      const revision = document.engine.descriptor.revision
      await frame.locator('#ribbon-tab-data').click()
      await frame.getByRole('button', { name: 'Subtotal', exact: true }).click()
      const dialog = frame.getByRole('dialog', { name: 'Subtotal', exact: true })
      await dialog.getByRole('button', { name: 'OK', exact: true }).click()
      const deadline2 = Date.now() + 120_000
      while (document.engine.descriptor.revision === revision && Date.now() < deadline2) {
        if (await frame.getByRole('alert').count())
          assert.fail(await frame.getByRole('alert').first().innerText())
        await new Promise((resolve) => setTimeout(resolve, 100))
      }
      await document.bridge!.request('view.flush', {})
      assert.equal(document.engine.descriptor.revision, revision + 1)
      const cells = document.engine.state.sheets[0]!.cells
      // 30 subtotal rows plus the grand total follow the 1500 data rows.
      assert.equal(cells.A37!.rawValue, 'G000 Total')
      assert.equal(cells.B37!.formula, '=SUBTOTAL(9,$B$2:$B$36)')
      assert.equal(cells.A1082!.rawValue, 'Grand Total')
      await vscode.commands.executeCommand('undo')
      assert.equal(document.engine.state.sheets[0]!.cells.A37?.rawValue ?? null, null)
    } finally {
      if (api.getDocument(uri))
        await vscode.commands.executeCommand('workbench.action.files.save', uri)
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
      await browser.close()
    }
  })
})
