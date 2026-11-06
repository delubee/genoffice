import * as assert from 'node:assert/strict'
import { promises as fs } from 'node:fs'
import * as vscode from 'vscode'
import { suite, test } from 'mocha'
import JSZip from 'jszip'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'

suite('Windows virtual-drive junction', () => {
  test('opens, edits, saves and reopens a junction when realpath reports UNKNOWN', async function () {
    if (process.platform !== 'win32') return this.skip()
    this.timeout(60000)
    const workspace = vscode.workspace.workspaceFolders![0]!.uri
    const target = vscode.Uri.joinPath(workspace, 'virtual-drive-target')
    const junction = vscode.Uri.joinPath(workspace, 'virtual-drive-junction')
    const uri = vscode.Uri.joinPath(junction, 'book.xlsx')
    const targetFile = vscode.Uri.joinPath(target, 'book.xlsx')
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as {
      getDocument(uri: vscode.Uri): WorkbookDocument
    }
    const realpath = fs.realpath
    let unsupportedCalls = 0
    try {
      await vscode.workspace.fs.createDirectory(target)
      await vscode.workspace.fs.writeFile(targetFile, await manualWorkbook())
      await fs.symlink(target.fsPath, junction.fsPath, 'junction')
      // Inject only the reported filesystem capability error. Reads, link
      // traversal, copy, atomic replacement and the real Webview remain real.
      fs.realpath = (async (...args: Parameters<typeof fs.realpath>) => {
        const path = String(args[0]).toLowerCase()
        if (
          path.startsWith(target.fsPath.toLowerCase()) ||
          path.startsWith(junction.fsPath.toLowerCase())
        ) {
          unsupportedCalls++
          throw Object.assign(new Error('UNKNOWN: unknown error, realpath'), {
            code: 'UNKNOWN',
            syscall: 'realpath',
          })
        }
        return Reflect.apply(realpath, fs, args)
      }) as typeof fs.realpath
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      let document = api.getDocument(uri)
      await document.whenViewReady()
      await document.bridge!.request('view.flush', {})
      await document.engine.submit({
        schemaVersion: 1,
        sessionId: document.sessionId,
        transactionId: 'junction-edit',
        expectedRevision: 0,
        origin: 'manual',
        summary: 'Edit a virtual-drive workbook',
        operations: [
          {
            op: 'set_cell',
            sheetId: document.engine.state.sheets[0]!.id,
            address: 'A1',
            value: 'junction saved',
          },
        ],
      })
      await vscode.commands.executeCommand('workbench.action.files.save')
      const zip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(targetFile))
      assert.ok(
        (await zip.file('xl/worksheets/sheet1.xml')!.async('string')).includes('junction saved'),
      )
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      document = api.getDocument(uri)
      await document.whenViewReady()
      await document.bridge!.request('view.flush', {})
      const shown = await document.bridge!.request('view.presentation', {
        stateId: document.engine.state.stateId,
        range: 'A1',
      })
      assert.ok(JSON.stringify(shown).includes('junction saved'))
      assert.ok(unsupportedCalls > 0)
    } finally {
      fs.realpath = realpath
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await fs.rm(junction.fsPath, { recursive: true, force: true })
      await fs.rm(target.fsPath, { recursive: true, force: true })
    }
  })
})
