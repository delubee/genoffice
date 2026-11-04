import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import { suite, test } from 'mocha'
import { chromium, type Frame } from '@playwright/test'
import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'
import { gotoReference } from './goto-helper'

suite('Desktop Flash Fill large ranges', () => {
  test('fills more than 2000 rows through one staged revision', async function () {
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
      'flash-fill-large-ui.xlsx',
    )
    // 2500 rows of first and last names in A and B; the joined form is typed as examples in C1:C2.
    const fixture = await applyCellEditsToXlsx(await manualWorkbook(), [
      ...Array.from({ length: 2500 }, (_, r) => ({
        sheetName: 'Data',
        row: r,
        column: 0,
        writeValue: true,
        cell: { value: `First${r}` },
      })),
      ...Array.from({ length: 2500 }, (_, r) => ({
        sheetName: 'Data',
        row: r,
        column: 1,
        writeValue: true,
        cell: { value: `Last${r}` },
      })),
      { sheetName: 'Data', row: 0, column: 2, writeValue: true, cell: { value: 'First0 Last0' } },
      { sheetName: 'Data', row: 1, column: 2, writeValue: true, cell: { value: 'First1 Last1' } },
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
            if (!(await page.title()).includes('flash-fill-large-ui.xlsx')) continue
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
      await gotoReference(frame, 'Data!C1:C2500')
      const revision = document.engine.descriptor.revision
      await frame.locator('#ribbon-tab-data').click()
      await frame.getByRole('button', { name: 'Flash Fill', exact: true }).first().click()
      const deadline2 = Date.now() + 120_000
      while (document.engine.descriptor.revision === revision && Date.now() < deadline2) {
        if (await frame.getByRole('alert').count())
          assert.fail(await frame.getByRole('alert').first().innerText())
        await new Promise((resolve) => setTimeout(resolve, 100))
      }
      await document.bridge!.request('view.flush', {})
      assert.equal(document.engine.descriptor.revision, revision + 1)
      const cells = document.engine.state.sheets[0]!.cells
      assert.equal(cells.C3!.rawValue, 'First2 Last2')
      assert.equal(cells.C2500!.rawValue, 'First2499 Last2499')
    } finally {
      if (api.getDocument(uri))
        await vscode.commands.executeCommand('workbench.action.files.save', uri)
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
      await browser.close()
    }
  })
})
