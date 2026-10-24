import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import { suite, test } from 'mocha'
import { link, readFile, writeFile } from 'node:fs/promises'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'
import type { WorkbookBindings } from '@genoffice/workbook-session/bindings'

suite('Product session freshness and disk conflicts', () => {
  test('invalidates reverted epochs, deduplicates receipts, locks aliases and detects external writes', async () => {
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as {
      getDocument(uri: vscode.Uri): WorkbookDocument | undefined
      bindings: WorkbookBindings
      revertForTests(uri: vscode.Uri): Promise<void>
      operationStatus(sessionId: string, transactionId: string): Promise<any>
    }
    assert.equal(
      typeof api.revertForTests,
      'function',
      'destructive fixture hook is restricted to VS Code Test mode',
    )
    const root = vscode.workspace.workspaceFolders![0]!.uri,
      uri = vscode.Uri.joinPath(root, 'session-freshness.xlsx'),
      alias = vscode.Uri.joinPath(root, 'session-alias.xlsx')
    try {
      await vscode.workspace.fs.writeFile(uri, await manualWorkbook())
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      const document = api.getDocument(uri)!
      const lifecycle = () => document.engine.descriptor.lifecycle
      await document.whenViewReady()
      await document.bridge!.request('view.flush', {})
      const sessionId = document.sessionId,
        sheetId = document.engine.state.sheets[0]!.id,
        binding = api.bindings.bind(sessionId, sheetId, 'A2')
      const intent = {
        schemaVersion: 1,
        sessionId,
        transactionId: 'native-freshness',
        expectedRevision: 0,
        origin: 'manual',
        summary: 'Edit fixture',
        operations: [{ op: 'set_cell', sheetId, address: 'A2', value: 15 }],
      }
      const receipt = await document.engine.submit(intent)
      assert.equal(api.bindings.get(binding.bindingId).status, 'stale')
      assert.deepEqual(await document.engine.submit(intent), receipt)
      assert.equal(document.engine.descriptor.revision, 1)
      assert.equal((await api.operationStatus(sessionId, intent.transactionId)).status, 'committed')
      await vscode.commands.executeCommand('workbench.action.files.save')
      assert.equal(
        document.engine.descriptor.lifecycle,
        'ready',
        'the watcher must recognize the extension own save',
      )
      await api.revertForTests(uri)
      assert.notEqual(document.sessionId, sessionId)
      assert.equal(api.bindings.get(binding.bindingId).status, 'invalidated')
      await assert.rejects(api.operationStatus(sessionId, intent.transactionId), {
        code: 'session_closed',
      })
      await document.whenViewReady()
      await document.bridge!.request('view.flush', {})
      await link(uri.fsPath, alias.fsPath)
      await Promise.resolve(
        vscode.commands.executeCommand('vscode.openWith', alias, 'genoffice.workbook'),
      ).catch(() => {})
      assert.equal(
        api.getDocument(alias),
        undefined,
        'a physical alias cannot own another writable session',
      )
      const stateId = document.engine.state.stateId
      await writeFile(uri.fsPath, await readFile(uri.fsPath))
      const deadline = Date.now() + 5_000
      while (lifecycle() !== 'conflict' && Date.now() < deadline)
        await new Promise((resolve) => setTimeout(resolve, 25))
      assert.equal(document.engine.descriptor.lifecycle, 'conflict')
      assert.equal(
        document.engine.state.stateId,
        stateId,
        'disk notifications do not mutate the canonical root',
      )
      await assert.rejects(
        document.engine.saveWith(async () => {
          throw new Error('The conflicted save must not start writing')
        }),
        { code: 'external_change_conflict' },
      )
      await api.revertForTests(uri)
      assert.equal(document.engine.descriptor.lifecycle, 'ready')
    } finally {
      const document = api.getDocument(uri)
      if (document?.engine.descriptor.dirty) await api.revertForTests(uri)
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await document?.whenDisposed()
      await vscode.workspace.fs.delete(alias, { useTrash: false }).then(undefined, () => {})
      await vscode.workspace.fs.delete(uri, { useTrash: false }).then(undefined, () => {})
    }
  })
})
