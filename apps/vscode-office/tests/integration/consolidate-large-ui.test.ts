import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import { suite, test } from 'mocha'
import { chromium, type Frame } from '@playwright/test'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'
import { gotoReference } from './goto-helper'

suite('Desktop Consolidate large outputs', () => {
  test('writes more than 2000 consolidated formulas through one staged revision', async function () {
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
      'consolidate-large-ui.xlsx',
    )
    const fixture = { buffer: await manualWorkbook() }
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
            if (!(await page.title()).includes('consolidate-large-ui.xlsx')) continue
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
      await gotoReference(frame, 'Other!E1')
      const revision = document.engine.descriptor.revision
      await frame.locator('#ribbon-tab-data').click()
      await frame.getByRole('button', { name: 'Consolidate', exact: true }).click()
      const dialog = frame.getByRole('dialog', { name: 'Consolidate', exact: true })
      await dialog.locator('input:not([type=checkbox])').fill('Data!A1:C800')
      await dialog.getByRole('button', { name: 'OK', exact: true }).click()
      const deadline2 = Date.now() + 120_000
      while (document.engine.descriptor.revision === revision && Date.now() < deadline2) {
        if (await dialog.getByRole('alert').count())
          assert.fail(await dialog.getByRole('alert').innerText())
        await new Promise((resolve) => setTimeout(resolve, 100))
      }
      await document.bridge!.request('view.flush', {})
      assert.equal(document.engine.descriptor.revision, revision + 1)
      const cells = document.engine.state.sheets[1]!.cells
      assert.equal(cells.E1!.formula, '=SUM(Data!A1)')
      assert.equal(cells.G800!.formula, '=SUM(Data!C800)')
      await vscode.commands.executeCommand('undo')
      assert.equal(document.engine.state.sheets[1]!.cells.G800?.formula ?? null, null)
    } finally {
      if (api.getDocument(uri))
        await vscode.commands.executeCommand('workbench.action.files.save', uri)
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
      await browser.close()
    }
  })
})
