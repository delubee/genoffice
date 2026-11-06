import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import JSZip from 'jszip'
import { suite, test } from 'mocha'
import { chromium, expect, type Frame } from '@playwright/test'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'
import { gotoReference } from './goto-helper'

suite('Desktop Consolidate UI', () => {
  test('persists cross-sheet formulas with native history and an independent reopen', async function () {
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
      'consolidate-ui.xlsx',
    )
    await vscode.workspace.fs.writeFile(uri, await manualWorkbook())
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
            if (!(await page.title()).includes('consolidate-ui.xlsx')) continue
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
      await gotoReference(frame, 'Data!C1')
      const before = structuredClone(document.engine.state)
      const revision = document.engine.descriptor.revision
      await frame.locator('#ribbon-tab-data').click()
      await frame.getByRole('button', { name: 'Consolidate', exact: true }).click()
      const dialog = frame.getByRole('dialog', { name: 'Consolidate', exact: true })
      const fn = dialog.getByRole('button', { name: 'Function', exact: true })
      await fn.click()
      await expect(dialog.getByRole('option')).toHaveText(['Sum', 'Count', 'Average', 'Max', 'Min'])
      await dialog.getByRole('option', { name: 'Max', exact: true }).click()
      await expect(fn).toContainText('Max')
      await fn.click()
      await dialog.getByRole('option', { name: 'Sum', exact: true }).click()
      const reference = dialog.locator('input:not([type=checkbox])')
      await reference.fill('Data!A2')
      await dialog.getByRole('button', { name: 'Add', exact: true }).click()
      await reference.fill('Other!A1')
      await dialog.getByRole('button', { name: 'OK', exact: true }).click()
      await document.bridge!.request('view.flush', {})
      if (await dialog.getByRole('alert').count())
        assert.fail(await dialog.getByRole('alert').innerText())
      await dialog.waitFor({ state: 'hidden' })
      assert.equal(document.engine.descriptor.revision, revision + 1)
      assert.equal(document.engine.state.sheets[0]!.cells.C1!.formula, '=SUM(Data!A2,Other!A1)')
      const after = structuredClone(document.engine.state)
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, before)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, after)
      await vscode.commands.executeCommand('workbench.action.files.save')
      const zip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      assert.match(
        await zip.file('xl/worksheets/sheet1.xml')!.async('string'),
        /<f>SUM\(Data!A2,Other!A1\)<\/f>/,
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
      assert.equal(api.getDocument(uri), undefined, 'the previous document must be disposed')
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      const reopened = api.getDocument(uri)!
      await reopened.whenViewReady()
      assert.notEqual(reopened.sessionId, previousSession)
      assert.equal(reopened.engine.descriptor.revision, 0)
      assert.equal(reopened.engine.descriptor.dirty, false)
      const snapshot = reopened.snapshot()
      try {
        const savedCells = await snapshot.read(reopened.engine.state.sheets[0]!.id, 'A1:C2')
        assert.equal(savedCells.cells.find((cell) => cell.address === 'A2')!.rawValue, 12.5)
        assert.equal(
          savedCells.cells.find((cell) => cell.address === 'C1')!.formula,
          '=SUM(Data!A2,Other!A1)',
        )
        const other = await snapshot.read(reopened.engine.state.sheets[1]!.id, 'A1')
        assert.equal(other.cells[0]!.rawValue, 7)
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
