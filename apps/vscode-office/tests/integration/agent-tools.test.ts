import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import { suite, test } from 'mocha'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'

suite('Product native workbook tools', () => {
  test('reads unsaved cells, binds selection and edits the explicit hidden workbook through registered LM tools', async () => {
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as {
      getDocument(uri: vscode.Uri): WorkbookDocument | undefined
      revertForTests(uri: vscode.Uri): Promise<void>
    }
    const root = vscode.workspace.workspaceFolders![0]!.uri,
      aUri = vscode.Uri.joinPath(root, 'tools-a.xlsx'),
      bUri = vscode.Uri.joinPath(root, 'tools-b.xlsx')
    const invoke = async (name: string, input: object) => {
      const result = await vscode.lm.invokeTool(name, { input, toolInvocationToken: undefined })
      const content = result.content.find((item) => item instanceof vscode.LanguageModelTextPart)
      assert.ok(content instanceof vscode.LanguageModelTextPart)
      return JSON.parse(content.value)
    }
    try {
      for (const uri of [aUri, bUri])
        await vscode.workspace.fs.writeFile(uri, await manualWorkbook())
      await vscode.commands.executeCommand('vscode.openWith', aUri, 'genoffice.workbook')
      const a = api.getDocument(aUri)!
      await a.whenViewReady()
      await a.bridge!.request('view.flush', {})
      const sheetId = a.engine.state.sheets[0]!.id
      await a.engine.submit({
        schemaVersion: 1,
        sessionId: a.sessionId,
        transactionId: 'manual-before-read',
        expectedRevision: 0,
        origin: 'manual',
        summary: 'Unsaved manual value',
        operations: [{ op: 'set_cell', sheetId, address: 'A2', value: 25 }],
      })
      const selection = await vscode.commands.executeCommand<{
        documentUri: string
        frozenRange: string
      }>('genoffice.bindSelection')
      assert.equal(selection!.documentUri, aUri.toString())
      assert.ok(selection!.frozenRange)
      for (const name of ['context', 'read', 'aggregate', 'apply', 'operationStatus'])
        assert.ok(vscode.lm.tools.some((tool) => tool.name === `genoffice_workbook_${name}`))
      for (const name of ['export', 'prepareImport', 'applyImport'])
        assert.ok(vscode.lm.tools.some((tool) => tool.name === `genoffice_workbook_${name}`))
      const context = await invoke('genoffice_workbook_context', {
        schemaVersion: 1,
        documentUri: aUri.toString(),
        sheetId,
        range: 'A1:A3',
      })
      assert.equal(context.status, 'ok')
      const bindingId = context.data.binding.bindingId,
        stateBefore = a.engine.state.stateId
      const read = await invoke('genoffice_workbook_read', {
        schemaVersion: 1,
        bindingId,
        sheetId,
        range: 'A1:A3',
        expectedRevision: 1,
      })
      assert.equal(read.status, 'ok')
      assert.equal(read.data.cells[0].rawValue, '00123')
      assert.equal(read.data.cells[1].rawValue, 25)
      assert.equal(a.engine.state.stateId, stateBefore, 'reading does not commit an AI edit')
      const tooManyFormats = await invoke('genoffice_workbook_read', {
        schemaVersion: 1,
        bindingId,
        sheetId,
        range: 'A1:A201',
        includeFormats: true,
      })
      assert.equal(tooManyFormats.code, 'range_limit')
      const aggregate = await invoke('genoffice_workbook_aggregate', {
        schemaVersion: 1,
        bindingId,
        sheetId,
        range: 'A1:A3',
        expectedRevision: 1,
      })
      assert.equal(aggregate.data.sum, 25)
      assert.equal(aggregate.data.nonEmptyCount, 3)
      await vscode.commands.executeCommand('vscode.openWith', bUri, 'genoffice.workbook')
      const b = api.getDocument(bUri)!
      await b.whenViewReady()
      await b.bridge!.request('view.flush', {})
      const bRoot = b.engine.state.stateId
      const input = {
        schemaVersion: 1,
        bindingId,
        expectedRevision: 1,
        transactionId: 'native-tool-edit',
        summary: 'Set A2 through native tool',
        operations: [{ op: 'set_cell', sheetId, address: 'A2', value: 40 }],
      }
      const edited = await invoke('genoffice_workbook_apply', input)
      assert.equal(edited.status, 'ok', JSON.stringify(edited))
      assert.equal(edited.data.revision, 2)
      assert.equal(a.engine.state.sheets[0]!.cells.A2!.rawValue, 40)
      assert.equal(b.engine.state.stateId, bRoot)
      const retry = await invoke('genoffice_workbook_apply', input)
      assert.deepEqual(retry.data, edited.data)
      assert.equal(a.engine.descriptor.revision, 2)
      const status = await invoke('genoffice_workbook_operationStatus', {
        schemaVersion: 1,
        sessionId: a.sessionId,
        transactionId: input.transactionId,
      })
      assert.equal(status.data.status, 'committed')
      const stale = await invoke('genoffice_workbook_apply', {
        ...input,
        transactionId: 'stale-edit',
      })
      assert.equal(stale.code, 'revision_conflict')
      await vscode.commands.executeCommand('vscode.openWith', aUri, 'genoffice.workbook')
      await a.requestHistory('undo')
      assert.equal(a.engine.state.sheets[0]!.cells.A2!.rawValue, 25)
      assert.equal(a.engine.history.length, 2)
      const structureBinding = await invoke('genoffice_workbook_context', {
        schemaVersion: 1,
        documentUri: aUri.toString(),
        sheetId,
        range: 'A1:A3',
      })
      const structure = await invoke('genoffice_workbook_apply', {
        schemaVersion: 1,
        bindingId: structureBinding.data.binding.bindingId,
        expectedRevision: 3,
        transactionId: 'native-tools-structure',
        summary: 'Insert a row',
        operations: [{ op: 'insert_rows', sheetId, row: 2, count: 1 }],
      })
      assert.equal(structure.status, 'ok', JSON.stringify(structure))
      const frozen = await invoke('genoffice_workbook_context', { schemaVersion: 1, bindingId })
      assert.equal(frozen.data.binding.selectionStatus, 'selection_stale_structure')
      await invoke('genoffice_workbook_read', {
        schemaVersion: 1,
        bindingId,
        sheetId,
        range: 'A1:A4',
        expectedRevision: 4,
      })
      const staleSelection = await invoke('genoffice_workbook_apply', {
        ...input,
        expectedRevision: 4,
        transactionId: 'old-selection',
      })
      assert.equal(staleSelection.code, 'selection_stale_structure')
      const explicit = await invoke('genoffice_workbook_apply', {
        ...input,
        usesFrozenSelection: false,
        expectedRevision: 4,
        transactionId: 'new-explicit-range',
        operations: [{ op: 'set_cell', sheetId, address: 'A4', value: 41 }],
      })
      assert.equal(explicit.status, 'ok', JSON.stringify(explicit))
      assert.equal(a.engine.state.sheets[0]!.cells.A4!.rawValue, 41)
      await vscode.commands.executeCommand('workbench.action.files.save')
    } finally {
      for (const uri of [aUri, bUri]) {
        const document = api.getDocument(uri)
        if (document?.engine.descriptor.dirty) {
          await api.revertForTests(uri)
          await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
          await document.whenViewReady()
          await vscode.commands.executeCommand('workbench.action.files.save')
        }
      }
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      for (const uri of [aUri, bUri])
        await vscode.workspace.fs.delete(uri, { useTrash: false }).then(undefined, () => {})
    }
  })
})
