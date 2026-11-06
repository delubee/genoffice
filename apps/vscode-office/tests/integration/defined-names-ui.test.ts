import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import { suite, test } from 'mocha'
import { chromium, type Frame } from '@playwright/test'
import JSZip from 'jszip'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'

suite('Canonical Defined Names', () => {
  test('synchronizes agent name edits with Go To, native history and independent reopen', async function () {
    const port = process.env.GENOFFICE_TEST_CDP_PORT
    if (!port) return this.skip()
    this.timeout(90_000)
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as { getDocument(uri: vscode.Uri): WorkbookDocument | undefined }
    const uri = vscode.Uri.joinPath(
      vscode.workspace.workspaceFolders![0]!.uri,
      'defined-names-ui.xlsx',
    )
    await vscode.workspace.fs.writeFile(uri, await manualWorkbook())
    const browser = await chromium.connectOverCDP(`http://127.0.0.1:${port}`)
    const findFrame = async (): Promise<Frame> => {
      const started = Date.now()
      while (Date.now() - started < 20_000) {
        for (const context of browser.contexts())
          for (const page of context.pages()) {
            if (!(await page.title()).includes('defined-names-ui.xlsx')) continue
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
      await document.bridge!.request('view.flush', {})
      const before = structuredClone(document.engine.state)
      await document.engine.submit({
        schemaVersion: 1,
        sessionId: document.sessionId,
        expectedRevision: document.engine.descriptor.revision,
        transactionId: crypto.randomUUID(),
        origin: 'agent',
        summary: 'Define workbook and local names',
        operations: [
          { op: 'add_defined_name', name: 'Point', ref: 'Data!$A$2' },
          {
            op: 'add_defined_name',
            name: 'Point',
            ref: 'Other!$A$1',
            scopeSheetId: before.sheets[1]!.id,
          },
        ],
      })
      await document.bridge!.request('view.flush', {})
      const after = structuredClone(document.engine.state)
      assert.deepEqual(
        after.sheets.map((sheet) => sheet.cells),
        before.sheets.map((sheet) => sheet.cells),
      )
      await vscode.commands.executeCommand('genoffice.workbook.goto')
      let go = frame.getByRole('dialog', { name: 'Go To', exact: true })
      await go.getByRole('button', { name: /Point/ }).waitFor()
      await go.getByRole('textbox', { name: 'Reference', exact: true }).fill('Point')
      await go.getByRole('button', { name: 'Go', exact: true }).click()
      await go.waitFor({ state: 'hidden' })
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, before)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, after)
      await vscode.commands.executeCommand('workbench.action.files.save')
      const saved = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const xml = await saved.file('xl/workbook.xml')!.async('string')
      assert.match(xml, /<definedName name="Point">Data!\$A\$2<\/definedName>/)
      assert.match(xml, /<definedName name="Point" localSheetId="1">Other!\$A\$1<\/definedName>/)
      assert.equal(
        await saved.file('customXml/manual-sentinel.xml')!.async('string'),
        '<preserve id="manual-fixture">unchanged</preserve>',
      )
      const session = document.sessionId
      await vscode.commands.executeCommand('workbench.action.closeActiveEditor')
      const closed = Date.now()
      while (api.getDocument(uri) && Date.now() - closed < 10000)
        await new Promise((resolve) => setTimeout(resolve, 50))
      assert.equal(api.getDocument(uri), undefined)
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      document = api.getDocument(uri)!
      await document.whenViewReady()
      assert.notEqual(document.sessionId, session)
      frame = await findFrame()
      await vscode.commands.executeCommand('genoffice.workbook.goto')
      go = frame.getByRole('dialog', { name: 'Go To', exact: true })
      await go.getByRole('button', { name: /Point/ }).waitFor()
      await go.getByRole('button', { name: 'Cancel', exact: true }).click()
    } finally {
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
      await browser.close()
    }
  })
})
