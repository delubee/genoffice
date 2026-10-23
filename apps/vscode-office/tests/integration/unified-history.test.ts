import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import { suite, test } from 'mocha'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'

suite('Product unified history', () => {
  test('keeps mixed history across save and refuses a background workbook history request', async () => {
    const api = (await vscode.extensions
        .getExtension('genoffice.genoffice-workbook')!
        .activate()) as {
        getDocument(uri: vscode.Uri): WorkbookDocument | undefined
        revertForTests(uri: vscode.Uri): Promise<void>
      },
      root = vscode.workspace.workspaceFolders![0]!.uri,
      aUri = vscode.Uri.joinPath(root, 'native-history.xlsx'),
      bUri = vscode.Uri.joinPath(root, 'native-history-b.xlsx')
    const dirty = (uri: vscode.Uri) =>
      vscode.window.tabGroups.all
        .flatMap((group) => group.tabs)
        .find(
          (tab) =>
            tab.input instanceof vscode.TabInputCustom &&
            tab.input.uri.toString() === uri.toString(),
        )?.isDirty
    try {
      await vscode.workspace.fs.writeFile(aUri, await manualWorkbook())
      await vscode.workspace.fs.writeFile(bUri, await manualWorkbook())
      await vscode.commands.executeCommand('vscode.openWith', aUri, 'genoffice.workbook')
      const a = api.getDocument(aUri)!,
        sessionId = a.sessionId,
        basePath = a.original.path,
        sheetId = a.engine.state.sheets[0]!.id,
        otherId = a.engine.state.sheets[1]!.id
      await a.whenViewReady()
      await a.bridge!.request('view.flush', {})
      for (const [index, operations] of [
        [{ op: 'set_cell', sheetId, address: 'A2', value: 20 }],
        [
          { op: 'set_cell', sheetId, address: 'A2', value: 30 },
          { op: 'set_cell', sheetId: otherId, address: 'A1', value: 9 },
        ],
        [{ op: 'format_range', sheetId, range: 'A2', format: { bold: false } }],
      ].entries())
        await a.engine.submit({
          schemaVersion: 1,
          sessionId,
          transactionId: `native-history-${index}`,
          expectedRevision: index,
          origin: index === 1 ? 'agent' : 'manual',
          summary: 'History fixture',
          operations,
        })
      assert.equal(a.engine.history.length, 3)
      await vscode.commands.executeCommand('workbench.action.files.save')
      const savedId = a.engine.state.stateId
      assert.equal(dirty(aUri), false)
      await a.requestHistory('undo')
      assert.equal(a.engine.descriptor.revision, 4)
      assert.equal(a.engine.state.sheets[0]!.cells.A2!.style?.bold, true)
      assert.equal(dirty(aUri), true)
      await a.requestHistory('undo')
      assert.equal(a.engine.state.sheets[0]!.cells.A2!.rawValue, 20)
      assert.equal(
        a.engine.state.sheets[1]!.cells.A1,
        undefined,
        'one undo reverts the entire two-cell agent batch',
      )
      await a.requestHistory('redo')
      await a.requestHistory('redo')
      assert.equal(a.engine.descriptor.revision, 7)
      assert.equal(a.engine.state.stateId, savedId)
      assert.equal(a.engine.descriptor.dirty, false)
      assert.equal(dirty(aUri), false)
      await vscode.commands.executeCommand('vscode.openWith', bUri, 'genoffice.workbook')
      const b = api.getDocument(bUri)!
      await b.whenViewReady()
      await b.bridge!.request('view.flush', {})
      await b.engine.submit({
        schemaVersion: 1,
        sessionId: b.sessionId,
        transactionId: 'other-document',
        expectedRevision: 0,
        origin: 'manual',
        summary: 'Other document',
        operations: [
          { op: 'set_cell', sheetId: b.engine.state.sheets[0]!.id, address: 'A2', value: 99 },
        ],
      })
      await vscode.commands.executeCommand('workbench.action.files.save')
      const bState = b.engine.state.stateId
      await assert.rejects(a.requestHistory('undo'), { code: 'view_unavailable' })
      assert.equal(a.engine.descriptor.revision, 7)
      assert.equal(b.engine.state.stateId, bState)
      await vscode.commands.executeCommand('vscode.openWith', aUri, 'genoffice.workbook')
      await a.requestHistory('undo')
      assert.equal(a.engine.descriptor.revision, 8)
      assert.equal(dirty(aUri), true)
      assert.equal(
        b.engine.state.stateId,
        bState,
        'history requests must not follow the active editor',
      )
      await a.requestHistory('redo')
      assert.equal(a.engine.descriptor.revision, 9)
      assert.equal(a.engine.state.stateId, savedId)
      assert.equal(dirty(aUri), false)
      assert.equal(a.sessionId, sessionId)
      assert.equal(a.original.path, basePath)
    } finally {
      for (const uri of [aUri, bUri])
        if (api.getDocument(uri)?.engine.descriptor.dirty) {
          await api.revertForTests(uri)
          await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
          await api.getDocument(uri)!.whenViewReady()
          await vscode.commands.executeCommand('workbench.action.files.save')
        }
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      for (const uri of [aUri, bUri])
        await vscode.workspace.fs.delete(uri, { useTrash: false }).then(undefined, () => {})
    }
  })
})
