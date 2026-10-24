import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import { suite, test } from 'mocha'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'
import { saveWorkbookDocument } from '../../src/extension/save'
import { XlsxSidecarClient } from '@genoffice/workbook-session/sidecar/client'
import { join } from 'node:path'
import { chmod } from 'node:fs/promises'
suite('Product workbook editor', () => {
  test('opens a real Webview, edits through its candidate, marks dirty and saves', async () => {
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as { getDocument(uri: vscode.Uri): WorkbookDocument }
    const uri = vscode.Uri.joinPath(
      vscode.workspace.workspaceFolders![0]!.uri,
      'manual-editor.xlsx',
    )
    const copyUri = vscode.Uri.joinPath(
      vscode.workspace.workspaceFolders![0]!.uri,
      'manual-copy.xlsx',
    )
    let writerClient: XlsxSidecarClient | undefined
    let document: WorkbookDocument | undefined
    const saveCancellation = new vscode.CancellationTokenSource()
    try {
      await vscode.workspace.fs.writeFile(uri, await manualWorkbook())
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      document = api.getDocument(uri)
      assert.ok(document)
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      assert.equal(api.getDocument(uri), document, 'opening the same URI reuses its document')
      await document.whenViewReady()
      assert.deepEqual(await document.bridge!.request('view.flush', {}), { ready: true })
      const sheetId = document.engine.state.sheets[0]!.id
      await document.engine.submit({
        schemaVersion: 1,
        sessionId: document.sessionId,
        transactionId: 'native-manual',
        expectedRevision: 0,
        origin: 'manual',
        summary: 'Manual native test',
        operations: [{ op: 'set_cell', sheetId, address: 'A2', value: 42.5 }],
      })
      assert.equal(document.engine.state.sheets[0]!.cells.A2!.rawValue, 42.5)
      assert.equal(
        document.engine.projectionError,
        undefined,
        'the visible grid must acknowledge the committed state',
      )
      assert.equal(document.engine.descriptor.dirty, true)
      assert.equal(vscode.window.tabGroups.activeTabGroup.activeTab?.isDirty, true)
      const sourceBytes = await vscode.workspace.fs.readFile(uri)
      writerClient = new XlsxSidecarClient(
        join(
          vscode.extensions.getExtension('genoffice.genoffice-workbook')!.extensionPath,
          'native',
          `${process.platform}-${process.arch}`,
          process.platform === 'win32' ? 'xlsx-sidecar.exe' : 'xlsx-sidecar',
        ),
      )
      const savedStateId = document.engine.descriptor.savedStateId
      await saveWorkbookDocument(document, copyUri, writerClient, saveCancellation.token, true)
      assert.deepEqual(
        await vscode.workspace.fs.readFile(uri),
        sourceBytes,
        'Save As leaves original source bytes intact',
      )
      assert.equal(
        document.engine.descriptor.savedStateId,
        savedStateId,
        'Save As does not mark the source clean',
      )
      assert.equal(document.engine.descriptor.dirty, true)
      const copy = (await writerClient.open(copyUri.fsPath)) as {
        sessionId: string
        sheets: { id: string }[]
      }
      const range = (await writerClient.readRange({
        sessionId: copy.sessionId,
        sheetId: copy.sheets[0]!.id,
        range: { startRow: 1, endRow: 1, startColumn: 0, endColumn: 0 },
      })) as { cells: { value: unknown }[] }
      assert.equal(range.cells[0]!.value, 42.5, 'Save As includes current unsaved edits')
      await writerClient.close(copy.sessionId)
      await vscode.commands.executeCommand('workbench.action.files.save')
      assert.equal(document.engine.descriptor.dirty, false)
      assert.equal(vscode.window.tabGroups.activeTabGroup.activeTab?.isDirty, false)
      await vscode.commands.executeCommand('vscode.openWith', copyUri, 'genoffice.workbook')
      const copiedDocument = api.getDocument(copyUri)
      assert.ok(copiedDocument)
      assert.notEqual(
        copiedDocument.sessionId,
        document.sessionId,
        'opening Save As creates a new session',
      )
      await copiedDocument.whenViewReady()
      assert.deepEqual(await copiedDocument.bridge!.request('view.flush', {}), { ready: true })
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await copiedDocument.whenDisposed()
      await chmod(copyUri.fsPath, 0o444)
      await vscode.commands.executeCommand('vscode.openWith', copyUri, 'genoffice.workbook')
      const readOnly = api.getDocument(copyUri)
      await readOnly.whenViewReady()
      assert.ok(readOnly.engine.state.sheets.every((sheet) => sheet.readOnly))
      await assert.rejects(
        saveWorkbookDocument(readOnly, copyUri, writerClient, saveCancellation.token),
        { code: 'save_failed' },
      )
      const rootBefore = readOnly.engine.state.stateId
      await assert.rejects(
        readOnly.engine.submit({
          schemaVersion: 1,
          sessionId: readOnly.sessionId,
          transactionId: 'readonly-attempt',
          expectedRevision: 0,
          origin: 'manual',
          summary: 'Rejected readonly edit',
          operations: [
            {
              op: 'set_cell',
              sheetId: readOnly.engine.state.sheets[0]!.id,
              address: 'A2',
              value: 999,
            },
          ],
        }),
      )
      assert.equal(readOnly.engine.state.stateId, rootBefore)
    } finally {
      saveCancellation.dispose()
      writerClient?.stop()
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await document?.whenDisposed()
      if (document) assert.equal(document.engine.descriptor.lifecycle, 'disposed')
      await vscode.workspace.fs.delete(uri, { useTrash: false })
      await chmod(copyUri.fsPath, 0o600).catch(() => {})
      await vscode.workspace.fs.delete(copyUri, { useTrash: false }).then(undefined, () => {})
    }
  })
})
