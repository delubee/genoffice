import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import { suite, test } from 'mocha'
import { chromium, type Frame } from '@playwright/test'
import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'
import { gotoReference } from './goto-helper'

suite('Desktop Text to Columns large selections', () => {
  test('splits more than 2000 output cells through one staged revision', async function () {
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
      'text-to-columns-large-ui.xlsx',
    )
    // 1500 text cells "a<i>,b<i>,c<i>": 3 output columns beyond the source, 6000 cells in total.
    const rows: string[][] = Array.from({ length: 1500 }, (_, index) => [
      `a${index},b${index},c${index}`,
    ])
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
            if (!(await page.title()).includes('text-to-columns-large-ui.xlsx')) continue
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
      await gotoReference(frame, 'Data!A1:A1500')
      const revision = document.engine.descriptor.revision
      await frame.locator('#ribbon-tab-data').click()
      await frame.getByRole('button', { name: 'Text to Columns', exact: true }).click()
      await frame.getByRole('option', { name: 'Split by Comma', exact: true }).click()
      await document.bridge!.request('view.flush', {})
      const deadline2 = Date.now() + 60_000
      while (document.engine.descriptor.revision === revision && Date.now() < deadline2) {
        if (await frame.getByRole('alert').count())
          assert.fail(await frame.getByRole('alert').first().innerText())
        await new Promise((resolve) => setTimeout(resolve, 100))
      }
      assert.equal(document.engine.descriptor.revision, revision + 1)
      const sheet = document.engine.state.sheets[0]!
      assert.equal(sheet.cells.A1!.rawValue, 'a0')
      assert.equal(sheet.cells.B1!.rawValue, 'b0')
      assert.equal(sheet.cells.C1!.rawValue, 'c0')
      assert.equal(sheet.cells.C1500!.rawValue, 'c1499')
    } finally {
      if (api.getDocument(uri))
        await vscode.commands.executeCommand('workbench.action.files.save', uri)
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
      await browser.close()
    }
  })
})
