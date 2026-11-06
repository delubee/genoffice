import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import JSZip from 'jszip'
import { randomUUID } from 'node:crypto'
import { suite, test } from 'mocha'
import { manualWorkbook } from '../fixtures/manual-workbook'
import { nativeOfficeActions } from '../../src/extension/office-actions'
import { captureWorkbookBinding } from '../../src/extension/commands/bind-selection'
import { flushCurrentEditor } from '../../src/extension/tools/current-state'
import type { WorkbookDocument } from '../../src/extension/workbook-document'
import type { WorkbookBindings } from '@genoffice/workbook-session/bindings'

suite('Native office menu host', () => {
  test('saves the explicit workbook through the native command and refuses an inactive target', async () => {
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as {
      getDocument(uri: vscode.Uri): WorkbookDocument | undefined
      bindings: WorkbookBindings
      revertForTests(uri: vscode.Uri): Promise<void>
    }
    const uri = vscode.Uri.joinPath(vscode.workspace.workspaceFolders![0]!.uri, 'office-menu.xlsx')
    await vscode.workspace.fs.writeFile(uri, await manualWorkbook())
    try {
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      const document = api.getDocument(uri)
      assert.ok(document)
      await document.whenViewReady()
      await flushCurrentEditor(document)
      const sessionId = document.sessionId
      const active = () => {
        const tab = vscode.window.tabGroups.activeTabGroup.activeTab?.input
        return tab instanceof vscode.TabInputCustom && tab.uri.toString() === uri.toString()
      }
      const handler = nativeOfficeActions({
        uri,
        active,
        available: () => api.getDocument(uri) === document && document.sessionId === sessionId,
        flush: () => flushCurrentEditor(document),
        bind: async (scope) => {
          const binding = await captureWorkbookBinding(api, uri, scope)
          return {
            kind: 'xlsx',
            scope,
            documentUri: binding.documentUri,
            bindingId: binding.bindingId,
            revision: binding.boundRevision,
            sheetId: binding.logicalSheetId,
            range: binding.frozenRange,
          }
        },
      })
      await document.engine.submit({
        schemaVersion: 1,
        sessionId,
        transactionId: randomUUID(),
        expectedRevision: document.engine.descriptor.revision,
        origin: 'manual',
        summary: 'Native menu save verification',
        operations: [
          {
            op: 'set_cell',
            sheetId: document.engine.state.sheets[0]!.id,
            address: 'A1',
            value: 83917,
          },
        ],
      })
      await handler({ action: 'save' })
      const saved = await vscode.workspace.fs.readFile(uri)
      const zip = await JSZip.loadAsync(saved)
      assert.match(await zip.file('xl/worksheets/sheet1.xml')!.async('string'), /83917/)
      assert.equal(
        await zip.file('customXml/manual-sentinel.xml')!.async('string'),
        '<preserve id="manual-fixture">unchanged</preserve>',
      )
      assert.ok((await vscode.commands.getCommands(true)).includes('workbench.action.chat.open'))
      await vscode.commands.executeCommand('workbench.action.files.newUntitledFile')
      await assert.rejects(handler({ action: 'save' }), /Activate this document/)
      assert.deepEqual(await vscode.workspace.fs.readFile(uri), saved)
    } finally {
      if (api.getDocument(uri)) await api.revertForTests(uri)
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
    }
  })
})
