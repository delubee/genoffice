import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import JSZip from 'jszip'
import { suite, test } from 'mocha'
import { randomUUID } from 'node:crypto'
import { BorderStyleTypes } from '@univerjs/core'
import { manualWorkbook } from '../fixtures/manual-workbook'
import { ribbonIntents } from '@genoffice/sheets-editor/ribbon-intents'
import { toSelectionFormat } from '../../../sheets/src/renderer/selection-format'
import type { WorkbookDocument } from '../../src/extension/workbook-document'

suite('Desktop Home rich formatting', () => {
  test('projects desktop font/alignment/border settings, saves XLSX and uses native undo', async () => {
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as {
      getDocument(uri: vscode.Uri): WorkbookDocument
      revertForTests(uri: vscode.Uri): Promise<void>
    }
    const uri = vscode.Uri.joinPath(
      vscode.workspace.workspaceFolders![0]!.uri,
      'home-rich-format.xlsx',
    )
    await vscode.workspace.fs.writeFile(uri, await manualWorkbook())
    try {
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      const document = api.getDocument(uri)
      await document.whenViewReady()
      await document.bridge!.request('view.flush', {})
      const before = structuredClone(document.engine.state)
      const target = {
        sheetId: before.sheets[0]!.id,
        address: 'A1:C3',
        format: toSelectionFormat({}, 'General'),
        doubleUnderline: false,
      }
      await document.engine.submit({
        schemaVersion: 1,
        sessionId: document.sessionId,
        expectedRevision: document.engine.descriptor.revision,
        transactionId: randomUUID(),
        origin: 'manual',
        summary: 'Desktop Home rich format',
        operations: [
          'underline:double',
          'superscript',
          'align:distributed',
          'rotate:-45',
          'indent:2',
          'cellprot:locked-off:hidden-on',
          'border:outer:#123456:double',
        ].flatMap((command) => ribbonIntents(command, target)),
      })
      const presentation = (await document.bridge!.request('view.presentation', {
        stateId: document.engine.state.stateId,
        range: 'A1:C3',
      })) as {
        styles: Record<
          string,
          {
            ul?: { t: number }
            va?: number
            tr?: { a: number }
            pd?: { l: number }
            bd?: { t?: { s: number }; l?: { s: number }; b?: { s: number }; r?: { s: number } }
          }
        >
      }
      assert.equal(presentation.styles.A1!.ul?.t, 10)
      assert.equal(presentation.styles.A1!.va, 3)
      assert.equal(presentation.styles.A1!.tr?.a, -45)
      assert.equal(presentation.styles.A1!.pd?.l, 16)
      assert.equal(presentation.styles.A1!.bd?.t?.s, BorderStyleTypes.DOUBLE)
      assert.equal(presentation.styles.A1!.bd?.l?.s, BorderStyleTypes.DOUBLE)
      assert.equal(presentation.styles.C3!.bd?.b?.s, BorderStyleTypes.DOUBLE)
      assert.equal(presentation.styles.C3!.bd?.r?.s, BorderStyleTypes.DOUBLE)
      await vscode.commands.executeCommand('workbench.action.files.save')
      const zip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const styles = await zip.file('xl/styles.xml')!.async('string')
      assert.match(styles, /<u val="double"/)
      assert.match(styles, /<vertAlign val="superscript"/)
      assert.match(styles, /textRotation="135"/)
      assert.match(styles, /indent="2"/)
      assert.match(styles, /<protection[^>]*locked="0"[^>]*hidden="1"/)
      assert.equal(
        await zip.file('customXml/manual-sentinel.xml')!.async('string'),
        '<preserve id="manual-fixture">unchanged</preserve>',
      )
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state.sheets[0]!.cells, before.sheets[0]!.cells)
      await vscode.commands.executeCommand('redo')
      assert.equal(document.engine.state.sheets[0]!.cells.A1!.style!.underlineStyle, 'double')
    } finally {
      if (api.getDocument(uri)) await api.revertForTests(uri)
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
    }
  })
})
