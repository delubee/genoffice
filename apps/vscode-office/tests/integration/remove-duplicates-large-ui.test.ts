import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import { suite, test } from 'mocha'
import { chromium, expect, type Frame } from '@playwright/test'
import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'
import { gotoReference } from './goto-helper'

suite('Desktop Remove Duplicates large selections', () => {
  test('reads a selection beyond 2000 cells in bands and compacts it through one staged revision', async function () {
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
      'remove-duplicates-large-ui.xlsx',
    )
    // A header plus 600 distinct rows, each appearing twice: 1201 rows x 3 columns = 3603 cells.
    const rows: (string | number)[][] = [['Name', 'Quantity', 'Group']]
    for (let index = 0; index < 600; index++)
      for (let copy = 0; copy < 2; copy++) rows.push([`Item ${index}`, index, `G${index % 7}`])
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
            if (!(await page.title()).includes('remove-duplicates-large-ui.xlsx')) continue
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
      await frame.getByRole('button', { name: 'Remove Duplicates', exact: true }).click()
      const dialog = frame.getByRole('dialog', { name: 'Remove Duplicates', exact: true })
      await dialog.getByRole('button', { name: 'OK', exact: true }).click()
      // Either the dialog closes on success or it keeps an alert explaining the refusal.
      await Promise.race([
        dialog.waitFor({ state: 'hidden', timeout: 60_000 }),
        dialog.getByRole('alert').waitFor({ timeout: 60_000 }),
      ])
      if (await dialog.getByRole('alert').count())
        assert.fail(await dialog.getByRole('alert').innerText())
      await expect(frame.getByRole('status')).toContainText('Removed 600 duplicate row(s).', {
        timeout: 60_000,
      })
      if (await dialog.getByRole('alert').count())
        assert.fail(await dialog.getByRole('alert').innerText())
      await document.bridge!.request('view.flush', {})
      assert.equal(document.engine.descriptor.revision, revision + 1)
      const sheet = document.engine.state.sheets[0]!
      assert.equal(sheet.cells.A601!.rawValue, 'Item 599')
      assert.equal(sheet.cells.A602!.rawValue ?? null, null)
      assert.equal(sheet.cells.C1201!.rawValue ?? null, null)
    } finally {
      if (api.getDocument(uri))
        await vscode.commands.executeCommand('workbench.action.files.save', uri)
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
      await browser.close()
    }
  })
})
