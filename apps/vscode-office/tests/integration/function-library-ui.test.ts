import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import { suite, test } from 'mocha'
import { chromium, type Frame } from '@playwright/test'
import JSZip from 'jszip'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'
import { gotoReference } from './goto-helper'

suite('Desktop Function Library UI', () => {
  test('inserts and edits shared function arguments with source previews, native history and independent reopen', async function () {
    const port = process.env.GENOFFICE_TEST_CDP_PORT
    if (!port) return this.skip()
    this.timeout(90_000)
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as { getDocument(uri: vscode.Uri): WorkbookDocument | undefined }
    const uri = vscode.Uri.joinPath(
      vscode.workspace.workspaceFolders![0]!.uri,
      'function-library-ui.xlsx',
    )
    await vscode.workspace.fs.writeFile(uri, await manualWorkbook())
    const browser = await chromium.connectOverCDP(`http://127.0.0.1:${port}`)
    const findFrame = async (): Promise<Frame> => {
      const started = Date.now()
      while (Date.now() - started < 20_000) {
        for (const context of browser.contexts())
          for (const page of context.pages()) {
            if (!(await page.title()).includes('function-library-ui.xlsx')) continue
            for (const frame of page.frames())
              if (
                await frame
                  .locator('.workbook-desktop-header')
                  .count()
                  .catch(() => 0)
              )
                return frame
          }
        await new Promise((resolve) => setTimeout(resolve, 100))
      }
      throw new Error('Function Library frame did not load')
    }
    try {
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      let document = api.getDocument(uri)!
      await document.whenViewReady()
      let frame = await findFrame()
      const goTo = async (reference: string) => {
        await gotoReference(frame, reference)
      }
      await goTo('C4')
      await frame.getByRole('tab', { name: 'Formulas', exact: true }).click()
      const before = structuredClone(document.engine.state)
      await frame.locator('.workbook-grid .univerjs-icon-fx-icon').locator('..').click()
      const catalog = frame.getByRole('dialog', { name: 'Insert Function', exact: true })
      await catalog.getByPlaceholder('Search functions…').fill('SUM')
      await catalog.getByRole('option', { name: /^SUM\b/ }).click()
      await catalog.getByRole('button', { name: 'OK', exact: true }).click()
      let args = frame.getByRole('dialog', { name: 'Function Arguments', exact: true })
      await args.getByRole('textbox').first().fill('A2,Other!A1')
      await args.locator('.fn-args-result').filter({ hasText: '19.5' }).waitFor()
      await args.getByRole('button', { name: 'OK', exact: true }).click()
      try {
        await args.waitFor({ state: 'hidden', timeout: 10000 })
      } catch {
        throw new Error('Function submission failed: ' + (await args.innerText()))
      }
      await document.bridge!.request('view.flush', {})
      const after = structuredClone(document.engine.state)
      assert.equal(after.sheets[0]!.cells.C4!.formula, '=SUM(A2,Other!A1)')
      assert.deepEqual(after.sheets[1]!.cells, before.sheets[1]!.cells)
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, before)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, after)
      await vscode.commands.executeCommand('workbench.action.files.save')
      const saved = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const xml = await saved.file('xl/worksheets/sheet1.xml')!.async('string')
      assert.match(xml, /<f>SUM\(A2,Other!A1\)<\/f>/)
      assert.equal(
        await saved.file('customXml/manual-sentinel.xml')!.async('string'),
        '<preserve id="manual-fixture">unchanged</preserve>',
      )
      const sessionId = document.sessionId
      await vscode.commands.executeCommand('workbench.action.closeActiveEditor')
      const closed = Date.now()
      while (api.getDocument(uri) && Date.now() - closed < 10000)
        await new Promise((resolve) => setTimeout(resolve, 50))
      assert.equal(api.getDocument(uri), undefined)
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      document = api.getDocument(uri)!
      await document.whenViewReady()
      assert.notEqual(document.sessionId, sessionId)
      frame = await findFrame()
      await goTo('C4')
      await frame.getByRole('tab', { name: 'Formulas', exact: true }).click()
      await frame
        .locator('.workbook-desktop-header')
        .getByRole('button', { name: 'Insert Function', exact: true })
        .click()
      // Existing calls skip the catalog and preload the desktop parameter fields.
      args = frame.getByRole('dialog', { name: 'Function Arguments', exact: true })
      await args.waitFor()
      assert.equal(
        await frame.getByRole('dialog', { name: 'Insert Function', exact: true }).count(),
        0,
      )
      assert.equal(await args.getByRole('textbox').nth(0).inputValue(), 'A2')
      assert.equal(await args.getByRole('textbox').nth(1).inputValue(), 'Other!A1')
      await args.getByRole('textbox').nth(0).fill('2')
      await args.getByRole('textbox').nth(1).fill('3')
      await args.getByRole('button', { name: 'OK', exact: true }).click()
      await args.waitFor({ state: 'hidden' })
      await document.bridge!.request('view.flush', {})
      assert.equal(document.engine.state.sheets[0]!.cells.C4!.formula, '=SUM(2, 3)')
      await vscode.commands.executeCommand('workbench.action.files.save')
      const edited = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      assert.match(
        await edited.file('xl/worksheets/sheet1.xml')!.async('string'),
        /<f>SUM\(2, 3\)<\/f>/,
      )
    } finally {
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
      await browser.close()
    }
  })
})
