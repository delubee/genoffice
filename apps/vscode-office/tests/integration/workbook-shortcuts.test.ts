import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import { suite, test } from 'mocha'
import { chromium, type Frame } from '@playwright/test'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'

suite('Workbook native shortcuts UI', () => {
  test('routes focused workbook shortcuts across ribbon tabs and leaves text editors on native Find', async function () {
    const port = process.env.GENOFFICE_TEST_CDP_PORT
    if (!port) return this.skip()
    this.timeout(90_000)
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as {
      getDocument(uri: vscode.Uri): WorkbookDocument | undefined
    }
    const uri = vscode.Uri.joinPath(vscode.workspace.workspaceFolders![0]!.uri, 'shortcuts-ui.xlsx')
    const original = await applyCellEditsToXlsx(await manualWorkbook(), [
      {
        sheetName: 'Data',
        row: 79,
        column: 0,
        writeValue: true,
        cell: { value: '00123' },
      },
    ])
    await vscode.workspace.fs.writeFile(uri, original.buffer)
    const browser = await chromium.connectOverCDP(`http://127.0.0.1:${port}`)
    try {
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      const document = api.getDocument(uri)!
      await document.whenViewReady()
      let frame: Frame | undefined
      const started = Date.now()
      while (!frame && Date.now() - started < 20_000) {
        for (const context of browser.contexts())
          for (const page of context.pages()) {
            if (!(await page.title()).includes('shortcuts-ui.xlsx')) continue
            for (const candidate of page.frames())
              if (
                await candidate
                  .locator('.workbook-desktop-header')
                  .count()
                  .catch(() => 0)
              )
                frame = candidate
          }
        if (!frame) await new Promise((resolve) => setTimeout(resolve, 100))
      }
      assert.ok(frame, 'workbook frame')
      const state = structuredClone(document.engine.state),
        revision = document.engine.descriptor.revision
      assert.equal(await vscode.commands.executeCommand('genoffice.workbook.goto'), true)
      const initialGoTo = frame.getByRole('dialog', { name: 'Go To', exact: true })
      await initialGoTo.getByRole('textbox', { name: 'Reference', exact: true }).fill('Data!A80')
      await initialGoTo.getByRole('button', { name: 'Go', exact: true }).click()
      await initialGoTo.waitFor({ state: 'hidden' })
      await frame.getByRole('tab', { name: 'Insert', exact: true }).click()
      const modifier = process.platform === 'darwin' ? 'Meta' : 'Control'
      await frame.page().keyboard.press(`${modifier}+f`)
      const find = frame.getByRole('dialog', { name: 'Find and Replace', exact: true })
      await find.waitFor({ state: 'visible', timeout: 10_000 })
      const query = find.getByRole('combobox', { name: 'Find what', exact: true })
      await query.fill('00123')
      await find.getByRole('status').filter({ hasText: '2 of 2' }).waitFor()
      await frame.page().keyboard.press(`${modifier}+f`)
      await frame.waitForFunction(
        (element) => window.document.activeElement === element,
        await query.elementHandle(),
      )
      const focused = await query.evaluate((element) => ({
        focused: window.document.activeElement === element,
        start: (element as HTMLInputElement).selectionStart,
        end: (element as HTMLInputElement).selectionEnd,
      }))
      assert.deepEqual(focused, { focused: true, start: 0, end: 5 })
      await frame.page().keyboard.press(`${modifier}+h`)
      await find.getByRole('combobox', { name: 'Replace with', exact: true }).waitFor()
      assert.equal(
        await find.getByRole('tab', { name: 'Replace', exact: true }).getAttribute('aria-selected'),
        'true',
      )
      await query.press('Escape')
      await find.waitFor({ state: 'hidden' })
      await frame.page().keyboard.press(`${modifier}+g`)
      const goTo = frame.getByRole('dialog', { name: 'Go To', exact: true })
      await goTo.waitFor({ state: 'visible' })
      await goTo.getByRole('textbox', { name: 'Reference', exact: true }).fill('Other!A1')
      await goTo.getByRole('button', { name: 'Go', exact: true }).click()
      await goTo.waitFor({ state: 'hidden' })
      await frame.waitForFunction(
        () => window.document.activeElement?.getAttribute('data-u-comp') === 'editor',
      )
      assert.equal(
        await frame.getByRole('tab', { name: 'Insert', exact: true }).getAttribute('aria-selected'),
        'true',
      )
      const selection = (await document.bridge!.request('view.selection', {})) as {
        selection: { sheetId: string; address: string }
      }
      assert.equal(selection.selection.sheetId, state.sheets[1]!.id)
      assert.match(selection.selection.address, /^A1(?::A1)?$/)
      await assert.rejects(
        document.bridge!.request('view.command', { command: 'execute', uri: uri.toString() }),
      )
      await assert.rejects(
        document.bridge!.request('view.command', { command: 'find', uri: uri.toString() }),
      )
      assert.equal(document.engine.descriptor.revision, revision)
      assert.deepEqual(document.engine.state, state)
      const text = await vscode.workspace.openTextDocument({ language: 'plaintext', content: '' })
      await vscode.window.showTextDocument(text)
      assert.equal(await vscode.commands.executeCommand('genoffice.workbook.find'), false)
      await frame.page().keyboard.press(`${modifier}+f`)
      await frame
        .page()
        .locator('.monaco-editor .find-widget')
        .first()
        .waitFor({ state: 'visible', timeout: 10_000 })
      assert.equal(await find.count(), 0)
      assert.deepEqual(document.engine.state, state)
    } finally {
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
      await browser.close()
    }
  })
})
