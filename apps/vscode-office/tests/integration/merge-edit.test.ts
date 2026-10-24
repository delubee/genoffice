import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import JSZip from 'jszip'
import { suite, test } from 'mocha'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'

suite('Merge editing', () => {
  test('unmerges original cells, merges a new range, and preserves geometry through history and save/reopen', async function () {
    this.timeout(60000)
    const uri = vscode.Uri.joinPath(vscode.workspace.workspaceFolders![0]!.uri, 'merge-edit.xlsx')
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as {
      getDocument(uri: vscode.Uri): WorkbookDocument
    }
    const original = { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 }
    const added = { startRow: 4, endRow: 5, startColumn: 2, endColumn: 3 }
    const geometry = async (document: WorkbookDocument) => {
      await document.bridge!.request('view.flush', {})
      const shown = (await document.bridge!.request('view.presentation', {
        stateId: document.engine.state.stateId,
        range: 'A1:D6',
      })) as { merges: unknown[] }
      return shown.merges
    }
    try {
      const fixture = await JSZip.loadAsync(await manualWorkbook())
      const xml = await fixture.file('xl/worksheets/sheet1.xml')!.async('string')
      fixture.file(
        'xl/worksheets/sheet1.xml',
        xml.replace(
          '</sheetData>',
          '</sheetData><mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>',
        ),
      )
      await vscode.workspace.fs.writeFile(uri, await fixture.generateAsync({ type: 'nodebuffer' }))
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      let document = api.getDocument(uri)
      await document.whenViewReady()
      assert.deepEqual(await geometry(document), [original])
      const sheetId = document.engine.state.sheets[0]!.id
      await document.engine.submit({
        schemaVersion: 1,
        sessionId: document.sessionId,
        transactionId: 'merge-edit',
        expectedRevision: 0,
        origin: 'agent',
        summary: 'Edit merged ranges',
        operations: [
          { op: 'unmerge_cells', sheetId, range: 'B1' },
          { op: 'set_cell', sheetId, address: 'B1', value: 'Unmerged value' },
          { op: 'set_cell', sheetId, address: 'C5', value: 'New heading' },
          { op: 'merge_cells', sheetId, range: 'C5:D6' },
        ],
      })
      assert.deepEqual(await geometry(document), [added])
      await document.engine.undo()
      assert.deepEqual(await geometry(document), [original])
      await document.engine.redo()
      assert.deepEqual(await geometry(document), [added])
      await vscode.commands.executeCommand('workbench.action.files.save')
      const saved = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const savedXml = await saved.file('xl/worksheets/sheet1.xml')!.async('string')
      assert.ok(savedXml.includes('<mergeCell ref="C5:D6"/>'))
      assert.ok(!savedXml.includes('<mergeCell ref="A1:B1"'))
      assert.equal(
        await saved.file('customXml/manual-sentinel.xml')!.async('string'),
        '<preserve id="manual-fixture">unchanged</preserve>',
      )
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      document = api.getDocument(uri)
      await document.whenViewReady()
      assert.deepEqual(await geometry(document), [added])
      const values = await document.bridge!.request('view.presentation', {
        stateId: document.engine.state.stateId,
        range: 'A1:D6',
      })
      assert.ok(JSON.stringify(values).includes('New heading'))
      assert.ok(JSON.stringify(values).includes('Unmerged value'))
    } finally {
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri).then(undefined, () => {})
    }
  })
})
