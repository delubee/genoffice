import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import JSZip from 'jszip'
import { suite, test } from 'mocha'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'

suite('Product worksheet dimensions', () => {
  test('loads original sizes, resizes with undo/redo, then saves and reopens exact sizes', async function () {
    this.timeout(60000)
    const uri = vscode.Uri.joinPath(vscode.workspace.workspaceFolders![0]!.uri, 'dimensions.xlsx')
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as { getDocument(uri: vscode.Uri): WorkbookDocument }
    const shown = async (document: WorkbookDocument) => {
      await document.bridge!.request('view.flush', {})
      return document.bridge!.request('view.presentation', {
        stateId: document.engine.state.stateId,
        range: 'A1:B3',
      }) as Promise<{ dimensions: { rows: { pixels: number }[]; columns: { pixels: number }[] } }>
    }
    try {
      const source = await JSZip.loadAsync(await manualWorkbook())
      const xml = (await source.file('xl/worksheets/sheet1.xml')!.async('string'))
        .replace(
          '<sheetData>',
          '<cols><col min="1" max="2" width="13" customWidth="1"/></cols><sheetData>',
        )
        .replace('<row r="2"', '<row ht="22" customHeight="1" r="2"')
      source.file('xl/worksheets/sheet1.xml', xml)
      await vscode.workspace.fs.writeFile(uri, await source.generateAsync({ type: 'nodebuffer' }))
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      let document = api.getDocument(uri)
      await document.whenViewReady()
      assert.equal((await shown(document)).dimensions.columns[1]!.pixels, 91)
      assert.ok(Math.abs((await shown(document)).dimensions.rows[1]!.pixels - (22 * 4) / 3) < 0.01)
      await document.engine.submit({
        schemaVersion: 1,
        sessionId: document.sessionId,
        transactionId: 'resize',
        expectedRevision: 0,
        origin: 'manual',
        summary: 'Resize dimensions',
        operations: [
          {
            op: 'set_col_width',
            sheetId: document.engine.state.sheets[0]!.id,
            column: 'B',
            count: 1,
            widthPx: 140,
          },
          {
            op: 'set_row_height',
            sheetId: document.engine.state.sheets[0]!.id,
            row: 2,
            count: 1,
            heightPoints: 30,
          },
        ],
      })
      assert.equal((await shown(document)).dimensions.columns[1]!.pixels, 140)
      assert.equal((await shown(document)).dimensions.rows[1]!.pixels, 40)
      await document.engine.undo()
      assert.equal((await shown(document)).dimensions.columns[1]!.pixels, 91)
      await document.engine.redo()
      assert.equal((await shown(document)).dimensions.columns[1]!.pixels, 140)
      await vscode.commands.executeCommand('workbench.action.files.save')
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      document = api.getDocument(uri)
      await document.whenViewReady()
      const reopened = await shown(document)
      assert.equal(reopened.dimensions.columns[1]!.pixels, 140)
      assert.equal(reopened.dimensions.rows[1]!.pixels, 40)
    } finally {
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri).then(undefined, () => {})
    }
  })
})
