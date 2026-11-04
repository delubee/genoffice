import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import JSZip from 'jszip'
import { suite, test } from 'mocha'
import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'

suite('Saved formula caches', () => {
  test('writes freshly calculated values for formulas that depend on an edit', async function () {
    this.timeout(90_000)
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as {
      getDocument(uri: vscode.Uri): WorkbookDocument | undefined
    }
    const uri = vscode.Uri.joinPath(
      vscode.workspace.workspaceFolders![0]!.uri,
      'save-calculation.xlsx',
    )
    const fixture = await applyCellEditsToXlsx(await manualWorkbook(), [
      { sheetName: 'Data', row: 0, column: 2, writeValue: true, cell: { value: 2 } },
      {
        sheetName: 'Data',
        row: 1,
        column: 2,
        writeValue: true,
        cell: { value: null, formula: '=C1*3' },
      },
      {
        sheetName: 'Other',
        row: 0,
        column: 1,
        writeValue: true,
        cell: { value: null, formula: '=Data!C2+A1' },
      },
    ])
    await vscode.workspace.fs.writeFile(uri, fixture.buffer)
    try {
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      const document = api.getDocument(uri)!
      await document.whenViewReady()
      await document.engine.submit({
        schemaVersion: 1,
        sessionId: document.sessionId,
        transactionId: 'edit-precedent',
        expectedRevision: document.engine.descriptor.revision,
        origin: 'manual',
        summary: 'Edit a precedent',
        operations: [
          { op: 'set_cell', sheetId: document.engine.state.sheets[0]!.id, address: 'C1', value: 5 },
        ],
      })
      // Tool and search reads see the host's results instead of `pending` formula cells.
      const snapshot = document.snapshot()
      try {
        const read = await snapshot.read(document.engine.state.sheets[0]!.id, 'C1:C2')
        const formula = read.cells.find((cell) => cell.address === 'C2')!
        assert.equal(formula.calculationState, 'ready')
        assert.equal(formula.cachedValue, 15)
        assert.equal(formula.rawValue, 15)
        const other = await snapshot.read(document.engine.state.sheets[1]!.id, 'B1')
        assert.equal(other.cells[0]!.cachedValue, 22)
      } finally {
        snapshot.release()
      }
      await vscode.commands.executeCommand('workbench.action.files.save')
      const zip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const data = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
      const other = await zip.file('xl/worksheets/sheet2.xml')!.async('string')
      assert.match(data, /<c\b[^>]*r="C2"[^>]*>\s*<f>C1\*3<\/f>\s*<v>15<\/v>/)
      // Other!A1 holds 7, so the dependent on another sheet is refreshed too.
      assert.match(other, /<c\b[^>]*r="B1"[^>]*>\s*<f>Data!C2\+A1<\/f>\s*<v>22<\/v>/)
    } finally {
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
    }
  })
})
