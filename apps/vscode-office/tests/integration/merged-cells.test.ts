import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import JSZip from 'jszip'
import { suite, test } from 'mocha'
import { mergedWorkbook, mergedRanges } from '../fixtures/merged-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'

suite('Product merged cells', () => {
  test('renders original merges, safely edits anchors and preserves merges after save and reopen', async function () {
    this.timeout(60000)
    const uri = vscode.Uri.joinPath(vscode.workspace.workspaceFolders![0]!.uri, 'merged.xlsx')
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as {
      getDocument(uri: vscode.Uri): WorkbookDocument
    }
    const shown = async (document: WorkbookDocument) => {
      await document.bridge!.request('view.flush', {})
      return document.bridge!.request('view.presentation', {
        stateId: document.engine.state.stateId,
        range: 'A1:D5',
      }) as Promise<{ merges: typeof mergedRanges; values: unknown[][] }>
    }
    try {
      await vscode.workspace.fs.writeFile(uri, await mergedWorkbook())
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      let document = api.getDocument(uri)
      await document.whenViewReady()
      assert.deepEqual((await shown(document)).merges, mergedRanges)
      const sheetId = document.engine.state.sheets[0]!.id
      for (const [index, origin] of (['manual', 'agent'] as const).entries()) {
        await document.engine.submit({
          schemaVersion: 1,
          sessionId: document.sessionId,
          transactionId: `merged-${origin}`,
          expectedRevision: index,
          origin,
          summary: 'Edit merged anchor',
          operations: [
            {
              op: 'set_cell',
              sheetId,
              address: index ? 'D3' : 'A1',
              value: index ? 'Vertical' : 'Updated heading',
            },
          ],
        })
        assert.deepEqual((await shown(document)).merges, mergedRanges)
      }
      const before = JSON.stringify(document.engine.state)
      await assert.rejects(
        document.engine.submit({
          schemaVersion: 1,
          sessionId: document.sessionId,
          transactionId: 'covered-write',
          expectedRevision: 2,
          origin: 'agent',
          summary: 'Reject covered cell',
          operations: [{ op: 'set_cell', sheetId, address: 'B1', value: 'Hidden' }],
        }),
        /merged_cell_target/,
      )
      assert.equal(JSON.stringify(document.engine.state), before)
      await vscode.commands.executeCommand('workbench.action.files.save')
      const zip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
      for (const ref of ['A1:C1', 'D3:D5', 'AX60:AZ62'])
        assert.ok(xml.includes(`<mergeCell ref="${ref}"`))
      assert.equal(
        await zip.file('customXml/manual-sentinel.xml')!.async('string'),
        '<preserve id="manual-fixture">unchanged</preserve>',
      )
      assert.ok(!/<c\b[^>]*\br="B1"/.test(xml), 'no hidden value is saved under the merge')
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      document = api.getDocument(uri)
      await document.whenViewReady()
      const reopened = await shown(document)
      assert.deepEqual(reopened.merges, mergedRanges)
      assert.equal(reopened.values[0]![0], 'Updated heading')
      assert.equal(reopened.values[2]![3], 'Vertical')
    } finally {
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri).then(undefined, () => {})
    }
  })
})
