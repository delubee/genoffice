import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import JSZip from 'jszip'
import { suite, test } from 'mocha'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'

suite('Product ribbon settings', () => {
  test('saves page, view, freeze, notes and filter settings into XLSX and restores the view', async function () {
    this.timeout(60000)
    const uri = vscode.Uri.joinPath(
      vscode.workspace.workspaceFolders![0]!.uri,
      'ribbon-settings.xlsx',
    )
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as { getDocument(uri: vscode.Uri): WorkbookDocument }
    const shown = async (document: WorkbookDocument) => {
      await document.bridge!.request('view.flush', {})
      return document.bridge!.request('view.presentation', {
        stateId: document.engine.state.stateId,
        range: 'A1:B3',
      }) as Promise<{
        view: { freeze: { ySplit: number }; zoom: number; gridlines: boolean }
        dimensions: { rows: { hidden: boolean }[] }
      }>
    }
    try {
      await vscode.workspace.fs.writeFile(uri, await manualWorkbook())
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      let document = api.getDocument(uri)
      await document.whenViewReady()
      await shown(document)
      const sheetId = document.engine.state.sheets[0]!.id
      await document.engine.submit({
        schemaVersion: 1,
        sessionId: document.sessionId,
        transactionId: 'ribbon-settings',
        expectedRevision: 0,
        origin: 'manual',
        summary: 'Ribbon settings',
        operations: [
          {
            op: 'set_page_setup',
            sheetId,
            orientation: 'landscape',
            paperSize: 8,
            margins: 'narrow',
            printArea: 'A1:B3',
            zoomScale: 125,
            showGridlines: false,
          },
          { op: 'set_freeze', sheetId, rows: 1, columns: 0 },
          { op: 'set_note', sheetId, address: 'A2', text: 'Check & approve' },
          { op: 'set_filter', sheetId, range: 'A1:A3' },
          { op: 'set_filter_criteria', sheetId, column: 'A', values: ['12.50'] },
        ],
      })
      assert.equal((await shown(document)).view.freeze.ySplit, 1)
      await document.engine.undo()
      assert.equal((await shown(document)).view.freeze.ySplit, 0)
      await document.engine.redo()
      await vscode.commands.executeCommand('workbench.action.files.save')
      const zip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
      assert.ok(xml.includes('orientation="landscape"'))
      assert.ok(xml.includes('paperSize="8"'))
      assert.ok(xml.includes('zoomScale="125"'))
      assert.ok(xml.includes('showGridLines="0"'))
      assert.ok(xml.includes('<autoFilter ref="A1:A3"'))
      assert.ok(xml.includes('val="12.50"'))
      const comments = Object.keys(zip.files).filter((name) => /comments[^/]*\.xml$/.test(name))
      assert.ok(comments.length)
      assert.ok((await zip.file(comments[0]!)!.async('string')).includes('Check &amp; approve'))
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      document = api.getDocument(uri)
      await document.whenViewReady()
      const reopened = await shown(document)
      assert.equal(reopened.view.freeze.ySplit, 1)
      assert.equal(reopened.view.zoom, 1.25)
      assert.equal(reopened.view.gridlines, false)
      assert.equal(reopened.dimensions.rows[2]!.hidden, true)
      const metadata = document.metadata as { sheets: { comments: { text: string }[] }[] }
      assert.equal(metadata.sheets[0]!.comments[0]!.text, 'Check & approve')
    } finally {
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri).then(undefined, () => {})
    }
  })
  test('sorts Chinese text consistently in the real Webview and host candidate', async function () {
    this.timeout(60000)
    const uri = vscode.Uri.joinPath(
      vscode.workspace.workspaceFolders![0]!.uri,
      'ribbon-chinese-sort.xlsx',
    )
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as { getDocument(uri: vscode.Uri): WorkbookDocument }
    try {
      await vscode.workspace.fs.writeFile(uri, await manualWorkbook())
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      const document = api.getDocument(uri)
      await document.whenViewReady()
      await document.bridge!.request('view.flush', {})
      const sheetId = document.engine.state.sheets[0]!.id
      await document.engine.submit({
        schemaVersion: 1,
        sessionId: document.sessionId,
        transactionId: 'chinese-data',
        expectedRevision: 0,
        origin: 'manual',
        summary: 'Chinese data',
        operations: [
          {
            op: 'set_range',
            sheetId,
            start: 'A2',
            values: [
              ['中', 2],
              ['阿', 1],
            ],
          },
        ],
      })
      await document.engine.submit({
        schemaVersion: 1,
        sessionId: document.sessionId,
        transactionId: 'chinese-sort',
        expectedRevision: 1,
        origin: 'agent',
        summary: 'Pinyin sort',
        operations: [
          {
            op: 'sort_range',
            sheetId,
            range: 'A2:B3',
            byColumn: 'A',
            order: 'asc',
            hasHeader: false,
          },
        ],
      })
      assert.equal(document.engine.state.sheets[0]!.cells.A2!.rawValue, '阿')
      assert.equal(document.engine.state.sheets[0]!.cells.B2!.rawValue, 1)
      const visible = (await document.bridge!.request('view.presentation', {
        stateId: document.engine.state.stateId,
        range: 'A2:B3',
      })) as { values: unknown[][] }
      assert.deepEqual(visible.values, [
        ['阿', 1],
        ['中', 2],
      ])
      await vscode.commands.executeCommand('workbench.action.files.save')
    } finally {
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri).then(undefined, () => {})
    }
  })
})
