import {
  measureNormalFontMdw,
  type NormalFontDescriptor,
} from '@genoffice/xlsx-gateway/domain/normal-font-metrics'
import { columnWidthToPixels } from '@genoffice/xlsx-gateway/domain/sheet-dimensions'
import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import { suite, test } from 'mocha'
import { chromium, type Frame } from '@playwright/test'
import JSZip from 'jszip'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'

suite('Desktop Paste Special UI', () => {
  test('uses the shared dialog across sheets with native history, selective fields and saved column widths', async function () {
    const port = process.env.GENOFFICE_TEST_CDP_PORT
    if (!port) return this.skip()
    this.timeout(90_000)
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as { getDocument(uri: vscode.Uri): WorkbookDocument | undefined }
    const uri = vscode.Uri.joinPath(
      vscode.workspace.workspaceFolders![0]!.uri,
      'paste-special-ui.xlsx',
    )
    await vscode.workspace.fs.writeFile(uri, await manualWorkbook())
    const browser = await chromium.connectOverCDP(`http://127.0.0.1:${port}`)
    try {
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      let document = api.getDocument(uri)!
      await document.whenViewReady()
      let frame: Frame | undefined
      const started = Date.now()
      while (!frame && Date.now() - started < 20_000) {
        for (const context of browser.contexts())
          for (const page of context.pages()) {
            if (!(await page.title()).includes('paste-special-ui.xlsx')) continue
            for (const candidate of page.frames())
              if (
                await candidate
                  .locator('.workbook-desktop-header')
                  .count()
                  .catch(() => 0)
              )
                frame = candidate
          }
        if (!frame) await new Promise((resolve) => setTimeout(resolve, 100))
      }
      assert.ok(frame)
      const sheetId = document.engine.state.sheets[0]!.id,
        otherId = document.engine.state.sheets[1]!.id
      await document.engine.submit({
        schemaVersion: 1,
        sessionId: document.sessionId,
        expectedRevision: document.engine.descriptor.revision,
        transactionId: crypto.randomUUID(),
        origin: 'manual',
        summary: 'Paste Special source and target',
        operations: [
          {
            op: 'paste_cells',
            sheetId,
            start: 'A80',
            cells: [
              [
                {
                  kind: 'number',
                  rawValue: 2,
                  numberFormat: '0.00',
                  style: { bold: true, numberFormat: '0.00' },
                },
                { kind: 'blank', rawValue: null },
              ],
              [
                { kind: 'number', rawValue: 4 },
                { kind: 'number', rawValue: 0 },
              ],
            ],
          },
          {
            op: 'paste_cells',
            sheetId: otherId,
            start: 'D5',
            cells: [
              [
                {
                  kind: 'number',
                  rawValue: 10,
                  style: {
                    italic: true,
                    fillColor: '#CCEEFF',
                    borderLeft: { style: 'double', color: '#123456' },
                  },
                },
                { kind: 'number', rawValue: 20 },
              ],
              [
                { kind: 'number', rawValue: 30 },
                { kind: 'number', rawValue: 40 },
              ],
            ],
          },
          { op: 'set_col_width', sheetId, column: 'A', widthPx: 80 },
          { op: 'set_col_width', sheetId, column: 'B', widthPx: 120 },
        ],
      })
      await frame.evaluate(() => {
        let items: ClipboardItem[] = []
        Object.defineProperty(navigator, 'clipboard', {
          configurable: true,
          value: {
            write: async (values: ClipboardItem[]) => {
              items = values
            },
            read: async () => items,
            readText: async () => (items[0] ? (await items[0].getType('text/plain')).text() : ''),
          },
        })
      })
      const navigate = async (reference: string) => {
        assert.equal(await vscode.commands.executeCommand('genoffice.workbook.goto'), true)
        const dialog = frame!.getByRole('dialog', { name: 'Go To', exact: true })
        await dialog.getByRole('textbox', { name: 'Reference', exact: true }).fill(reference)
        await dialog.getByRole('button', { name: 'Go', exact: true }).click()
        await dialog.waitFor({ state: 'hidden' })
      }
      const choose = async (name: string) => {
        await frame!
          .locator('.workbook-desktop-header')
          .getByRole('button', { name: 'Paste Special', exact: true })
          .click()
        await frame!.getByRole('option', { name, exact: true }).click()
      }
      await navigate('Data!A80:B81')
      await frame
        .locator('.workbook-desktop-header')
        .getByRole('button', { name: /^Copy / })
        .click()
      await document.bridge!.request('view.flush', {})
      const before = structuredClone(document.engine.state)
      await navigate('Other!D5')
      await choose('Paste Special…')
      const dialog = frame.getByRole('dialog', { name: 'Paste Special', exact: true })
      await dialog.getByRole('radio', { name: 'Values', exact: true }).check()
      await dialog.getByRole('radio', { name: 'Add', exact: true }).check()
      await dialog.getByRole('checkbox', { name: 'Transpose', exact: true }).check()
      await dialog.getByRole('checkbox', { name: 'Skip blanks', exact: true }).check()
      await dialog.getByRole('button', { name: 'OK', exact: true }).click()
      await dialog.waitFor({ state: 'hidden' })
      await document.bridge!.request('view.flush', {})
      const after = structuredClone(document.engine.state)
      const target = after.sheets.find((s) => s.id === otherId)!.cells
      assert.deepEqual(
        ['D5', 'E5', 'D6', 'E6'].map((address) => target[address]!.rawValue),
        [12, 24, 30, 40],
      )
      assert.equal(target.D5!.style!.fillColor, '#CCEEFF')
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, before)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, after)
      await choose('Paste Special…')
      await dialog.getByRole('radio', { name: 'Values and number formats', exact: true }).check()
      await dialog.getByRole('button', { name: 'OK', exact: true }).click()
      await dialog.waitFor({ state: 'hidden' })
      await document.bridge!.request('view.flush', {})
      const selected = document.engine.state.sheets.find((s) => s.id === otherId)!.cells.D5!
      assert.equal(selected.rawValue, 2)
      assert.equal(selected.numberFormat, '0.00')
      assert.equal(selected.style!.italic, true)
      assert.equal(selected.style!.fillColor, '#CCEEFF')
      assert.equal(selected.style!.borderLeft!.style, 'double')
      await choose('Column Widths Only')
      await document.bridge!.request('view.flush', {})
      const widths = document.engine.state.sheets.find((s) => s.id === otherId)!.features!
        .colWidths as Record<string, number>
      assert.equal(widths.D, 80)
      assert.equal(widths.E, 120)
      await vscode.commands.executeCommand('workbench.action.files.save')
      const zip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const xml = await zip.file('xl/worksheets/sheet2.xml')!.async('string')
      assert.match(xml, /<col\b[^>]*min="4"[^>]*max="4"/)
      const mdw = measureNormalFontMdw(
        (document.engine.state.features?.originalMetadata ?? {}) as NormalFontDescriptor,
      )
      for (const [column, pixels] of [
        [4, 80],
        [5, 120],
      ]) {
        const record = [...xml.matchAll(/<col\b[^>]*>/g)].find((match) => {
          const low = Number(/\bmin="(\d+)"/.exec(match[0])?.[1]),
            high = Number(/\bmax="(\d+)"/.exec(match[0])?.[1])
          return low <= column! && high >= column!
        })
        assert.ok(record)
        assert.equal(
          columnWidthToPixels(Number(/\bwidth="([^"]+)"/.exec(record[0])?.[1]), mdw),
          pixels,
        )
      }

      assert.match(await zip.file('xl/styles.xml')!.async('string'), /<left style="double"/)
      assert.equal(
        await zip.file('customXml/manual-sentinel.xml')!.async('string'),
        '<preserve id="manual-fixture">unchanged</preserve>',
      )
      const sessionId = document.sessionId
      await vscode.commands.executeCommand('workbench.action.closeActiveEditor')
      const closed = Date.now()
      while (api.getDocument(uri) && Date.now() - closed < 10_000)
        await new Promise((resolve) => setTimeout(resolve, 50))
      assert.equal(api.getDocument(uri), undefined)
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      document = api.getDocument(uri)!
      await document.whenViewReady()
      assert.notEqual(document.sessionId, sessionId)
      const snapshot = document.snapshot()
      try {
        const read = await snapshot.read(document.engine.state.sheets[1]!.id, 'D5:E6')
        assert.equal(read.cells.find((cell) => cell.address === 'D5')!.rawValue, 2)
        assert.equal(read.cells.find((cell) => cell.address === 'D5')!.numberFormat, '0.00')
        assert.equal(read.cells.find((cell) => cell.address === 'D5')!.style!.italic, true)
        assert.equal(
          read.cells.find((cell) => cell.address === 'D5')!.style!.borderLeft!.style,
          'double',
        )
      } finally {
        await snapshot.release()
      }
    } finally {
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
      await browser.close()
    }
  })
})
