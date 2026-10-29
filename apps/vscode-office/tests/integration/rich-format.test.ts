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
          { op: 'set_cell', sheetId: target.sheetId, address: 'A1', value: 341504.77 },
          { op: 'set_cell', sheetId: target.sheetId, address: 'B1', value: 0.12 },
          ...[
            'underline:double',
            'superscript',
            'align:distributed',
            'rotate:-45',
            'indent:2',
            'cellprot:locked-off:hidden-on',
            'border:outer:#123456:double',
          ].flatMap((command) => ribbonIntents(command, target)),
          { op: 'format_range', sheetId: target.sheetId, range: 'B1', format: { indent: 4 } },
        ],
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
      // Relative commands use each selected cell's own formatting and form
      // one native history entry even when they expand into several operations.
      const richState = structuredClone(document.engine.state.sheets[0]!.cells)
      const relativeTarget = {
        ...target,
        address: 'A1:B1',
        cellIndents: [
          { address: 'A1', indent: 2 },
          { address: 'B1', indent: 4 },
        ],
        decimalCells: [
          { row: 0, col: 0, pattern: 'General', raw: { v: richState.A1!.rawValue, t: 2 } },
          { row: 0, col: 1, pattern: '0.00%', raw: { v: richState.B1!.rawValue, t: 2 } },
        ],
      }
      await document.engine.submit({
        schemaVersion: 1,
        sessionId: document.sessionId,
        expectedRevision: document.engine.descriptor.revision,
        transactionId: randomUUID(),
        origin: 'manual',
        summary: 'Relative desktop formats',
        operations: ['indent-step:1', 'decimal-inc'].flatMap((command) =>
          ribbonIntents(command, relativeTarget),
        ),
      })
      assert.equal(document.engine.state.sheets[0]!.cells.A1!.style!.indent, 3)
      assert.equal(document.engine.state.sheets[0]!.cells.B1!.style!.indent, 5)
      assert.equal(document.engine.state.sheets[0]!.cells.A1!.style!.numberFormat, '0.000')
      assert.equal(document.engine.state.sheets[0]!.cells.B1!.style!.numberFormat, '0.000%')
      await vscode.commands.executeCommand('workbench.action.files.save')
      const relativeZip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const relativeStyles = await relativeZip.file('xl/styles.xml')!.async('string')
      assert.match(relativeStyles, /indent="3"/)
      assert.match(relativeStyles, /indent="5"/)
      assert.match(relativeStyles, /formatCode="0.000%"/)
      const relativePresentation = (await document.bridge!.request('view.presentation', {
        stateId: document.engine.state.stateId,
        range: 'A1:B1',
      })) as { styles: Record<string, { pd?: { l?: number }; n?: { pattern?: string } }> }
      assert.equal(relativePresentation.styles.A1?.pd?.l, 24)
      assert.equal(relativePresentation.styles.B1?.pd?.l, 40)
      assert.equal(relativePresentation.styles.B1?.n?.pattern, '0.000%')
      const beforeFill = structuredClone(document.engine.state.sheets[0]!.cells)
      await document.engine.submit({
        schemaVersion: 1,
        sessionId: document.sessionId,
        expectedRevision: document.engine.descriptor.revision,
        transactionId: randomUUID(),
        origin: 'manual',
        summary: 'Desktop formatting fill',
        operations: [
          ...ribbonIntents('fill-down', { ...target, address: 'A1:A2' }),
          ...ribbonIntents('fill-left', { ...target, address: 'C2:D2' }),
        ],
      })
      assert.equal(document.engine.state.sheets[0]!.cells.A2!.styleReset, true)
      assert.equal(document.engine.state.sheets[0]!.cells.A2!.style!.indent, 3)
      assert.equal(document.engine.state.sheets[0]!.cells.C2!.styleReset, true)
      assert.deepEqual(document.engine.state.sheets[0]!.cells.C2!.style, {})
      const fillPresentation = (await document.bridge!.request('view.presentation', {
        stateId: document.engine.state.stateId,
        range: 'A1:C3',
      })) as {
        styles: Record<
          string,
          { pd?: { l?: number }; ul?: { s?: number }; n?: { pattern?: string } }
        >
      }
      assert.equal(fillPresentation.styles.A2?.pd?.l, 24)
      assert.equal(fillPresentation.styles.A2?.n?.pattern, '0.000')
      assert.equal(fillPresentation.styles.C2?.ul?.s, 0)
      await vscode.commands.executeCommand('workbench.action.files.save')
      const fillZip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const fillSheet = await fillZip.file('xl/worksheets/sheet1.xml')!.async('string')
      const fillStyles = await fillZip.file('xl/styles.xml')!.async('string')
      const cellStyle = (address: string) => {
        const cell = new RegExp('<c\\b[^>]*\\br="' + address + '"[^>]*>').exec(fillSheet)?.[0]
        assert.ok(cell, 'saved cell ' + address)
        const styleIndex = Number(/\bs="(\d+)"/.exec(cell)?.[1] ?? 0)
        const xfs = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/
          .exec(fillStyles)![1]!
          .match(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g)!
        return xfs[styleIndex]!
      }
      assert.match(cellStyle('A2'), /indent="3"/)
      assert.match(cellStyle('A2'), /fillId="0"/)
      assert.match(cellStyle('C2'), /fontId="0"/)
      assert.match(cellStyle('C2'), /numFmtId="0"/)
      assert.doesNotMatch(cellStyle('C2'), /indent=|<protection|<alignment/)
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state.sheets[0]!.cells, beforeFill)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state.sheets[0]!.cells.C2!.style, {})
      await vscode.commands.executeCommand('undo')
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state.sheets[0]!.cells, richState)
      await vscode.commands.executeCommand('redo')
      assert.equal(document.engine.state.sheets[0]!.cells.B1!.style!.indent, 5)
      await vscode.commands.executeCommand('undo')
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state.sheets[0]!.cells, before.sheets[0]!.cells)
      await vscode.commands.executeCommand('redo')
      assert.equal(document.engine.state.sheets[0]!.cells.A1!.style!.underlineStyle, 'double')
      await document.engine.submit({
        schemaVersion: 1,
        sessionId: document.sessionId,
        expectedRevision: document.engine.descriptor.revision,
        transactionId: randomUUID(),
        origin: 'manual',
        summary: 'Seed clear targets',
        operations: [
          { op: 'set_note', sheetId: target.sheetId, address: 'A1', text: 'remove note' },
          { op: 'set_note', sheetId: target.sheetId, address: 'B1', text: 'keep note' },
          {
            op: 'set_hyperlink',
            sheetId: target.sheetId,
            address: 'A1',
            target: 'https://example.com/remove',
          },
          {
            op: 'set_hyperlink',
            sheetId: target.sheetId,
            address: 'B1',
            target: 'https://example.com/keep',
          },
        ],
      })
      await vscode.commands.executeCommand('workbench.action.files.save')
      const linkedZip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      assert.match(await linkedZip.file('xl/worksheets/sheet1.xml')!.async('string'), /<hyperlinks/)
      const beforeClear = structuredClone(document.engine.state.sheets[0]!)
      await document.engine.submit({
        schemaVersion: 1,
        sessionId: document.sessionId,
        expectedRevision: document.engine.descriptor.revision,
        transactionId: randomUUID(),
        origin: 'manual',
        summary: 'Desktop Clear All',
        operations: ribbonIntents('clear-all', {
          ...target,
          address: 'A1',
          noteAddresses: ['A1'],
          hyperlinkAddresses: ['A1'],
        }),
      })
      assert.equal(document.engine.state.sheets[0]!.cells.A1!.rawValue, null)
      await vscode.commands.executeCommand('workbench.action.files.save')
      const clearedZip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const clearedSheet = await clearedZip.file('xl/worksheets/sheet1.xml')!.async('string')
      assert.doesNotMatch(clearedSheet, /<hyperlink\b[^>]*ref="A1"/)
      assert.match(clearedSheet, /<hyperlink\b[^>]*ref="B1"/)
      const commentEntry = Object.keys(clearedZip.files).find((path) =>
        /comments.*\.xml$/i.test(path),
      )!
      const comments = await clearedZip.file(commentEntry)!.async('string')
      assert.doesNotMatch(comments, /remove note/)
      assert.match(comments, /keep note/)
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state.sheets[0]!, beforeClear)
      await vscode.commands.executeCommand('redo')
      assert.equal(document.engine.state.sheets[0]!.cells.A1!.rawValue, null)
      const beforeSizes = structuredClone(document.engine.state.sheets[0]!)
      await document.engine.submit({
        schemaVersion: 1,
        sessionId: document.sessionId,
        expectedRevision: document.engine.descriptor.revision,
        transactionId: randomUUID(),
        origin: 'manual',
        summary: 'Desktop axis sizes',
        operations: [
          ...ribbonIntents('row-height:25.3', { ...target, address: 'B2:C3' }),
          ...ribbonIntents('col-width:10', { ...target, address: 'B2:C3', maximumDigitWidth: 8 }),
        ],
      })
      await vscode.commands.executeCommand('workbench.action.files.save')
      const sizedZip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const sizedSheet = await sizedZip.file('xl/worksheets/sheet1.xml')!.async('string')
      assert.match(sizedSheet, /<row\b[^>]*r="2"[^>]*ht="25.5"/)
      assert.match(sizedSheet, /<col\b[^>]*min="2"[^>]*width="10"/)
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state.sheets[0]!, beforeSizes)
      await vscode.commands.executeCommand('redo')
      await vscode.commands.executeCommand('workbench.action.files.save')
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      const reopened = api.getDocument(uri)
      await reopened.whenViewReady()
      const sizes = (await reopened.bridge!.request('view.presentation', {
        stateId: reopened.engine.state.stateId,
        range: 'B2:C3',
      })) as { dimensions: { rows: { pixels: number }[]; columns: { pixels: number }[] } }
      assert.deepEqual(
        sizes.dimensions.rows.map((row) => row.pixels),
        [34, 34],
      )
      assert.deepEqual(
        sizes.dimensions.columns.map((column) => column.pixels),
        [80, 80],
      )
      await reopened.engine.submit({
        schemaVersion: 1,
        sessionId: reopened.sessionId,
        expectedRevision: reopened.engine.descriptor.revision,
        transactionId: randomUUID(),
        origin: 'manual',
        summary: 'Zero axis sizes',
        operations: [
          ...ribbonIntents('row-height:0', {
            ...target,
            sheetId: reopened.engine.state.sheets[0]!.id,
            address: 'B2',
          }),
          ...ribbonIntents('col-width:0', {
            ...target,
            sheetId: reopened.engine.state.sheets[0]!.id,
            address: 'B2',
            maximumDigitWidth: 8,
          }),
        ],
      })
      await vscode.commands.executeCommand('workbench.action.files.save')
      const zeroZip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const zeroSheet = await zeroZip.file('xl/worksheets/sheet1.xml')!.async('string')
      assert.match(zeroSheet, /<row\b[^>]*r="2"[^>]*ht="0"/)
      assert.match(zeroSheet, /<col\b[^>]*min="2"[^>]*width="0"/)
    } finally {
      if (api.getDocument(uri)) await api.revertForTests(uri)
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
    }
  })
})
