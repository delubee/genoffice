import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import { suite, test } from 'mocha'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'

suite('Product workbook appearance', () => {
  test('defaults to Excel, applies persisted settings live and preserves unsaved state', async function () {
    this.timeout(60000)
    const root = vscode.workspace.workspaceFolders![0]!.uri
    const uri = vscode.Uri.joinPath(root, 'appearance.xlsx')
    const config = vscode.workspace.getConfiguration('genoffice.workbook', uri)
    const previous = config.inspect('appearance')?.workspaceFolderValue
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as { getDocument(uri: vscode.Uri): WorkbookDocument }
    try {
      await config.update('appearance', undefined, vscode.ConfigurationTarget.WorkspaceFolder)
      assert.equal(config.inspect('appearance')?.defaultValue, 'excel')
      await vscode.workspace.fs.writeFile(uri, await manualWorkbook())
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      const document = api.getDocument(uri)
      await document.whenViewReady()
      await document.bridge!.request('view.flush', {})
      const shown = () =>
        document.bridge!.request('view.presentation', {
          stateId: document.engine.state.stateId,
          range: 'A2',
        }) as Promise<{ appearance: string; darkMode: boolean }>
      assert.equal((await shown()).appearance, 'excel')
      assert.equal((await shown()).darkMode, false)
      await document.engine.submit({
        schemaVersion: 1,
        sessionId: document.sessionId,
        transactionId: 'appearance-edit',
        expectedRevision: 0,
        origin: 'manual',
        summary: 'Unsaved appearance regression',
        operations: [
          {
            op: 'set_cell',
            sheetId: document.engine.state.sheets[0]!.id,
            address: 'A2',
            value: 941,
          },
        ],
      })
      const state = JSON.stringify(document.engine.state)
      for (const appearance of ['vscode', 'excel']) {
        await config.update('appearance', appearance, vscode.ConfigurationTarget.WorkspaceFolder)
        const deadline = Date.now() + 10000
        while ((await shown()).appearance !== appearance && Date.now() < deadline)
          await new Promise((resolve) => setTimeout(resolve, 50))
        assert.equal((await shown()).appearance, appearance)
        assert.equal(JSON.stringify(document.engine.state), state)
        assert.equal(document.engine.descriptor.dirty, true)
      }
      await vscode.commands.executeCommand('workbench.action.files.save')
    } finally {
      await config.update('appearance', previous, vscode.ConfigurationTarget.WorkspaceFolder)
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri).then(undefined, () => {})
    }
  })
})
