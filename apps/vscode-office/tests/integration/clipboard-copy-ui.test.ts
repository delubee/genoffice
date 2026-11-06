import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import { suite, test } from 'mocha'
import { chromium, type Frame } from '@playwright/test'
import JSZip from 'jszip'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'

suite('Desktop Copy UI', () => {
  test('copies typed cells across sheets, retains the snapshot after edits, undoes and saves then reopens', async function () {
    const port = process.env.GENOFFICE_TEST_CDP_PORT
    if (!port) return this.skip()
    this.timeout(90_000)
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as {
      getDocument(uri: vscode.Uri): WorkbookDocument | undefined
    }
    const uri = vscode.Uri.joinPath(
      vscode.workspace.workspaceFolders![0]!.uri,
      'clipboard-copy-ui.xlsx',
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
            if (!(await page.title()).includes('clipboard-copy-ui.xlsx')) continue
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
      const sheetId = document.engine.state.sheets[0]!.id
      const otherId = document.engine.state.sheets[1]!.id
      await document.engine.submit({
        schemaVersion: 1,
        sessionId: document.sessionId,
        expectedRevision: document.engine.descriptor.revision,
        transactionId: crypto.randomUUID(),
        origin: 'manual',
        summary: 'Typed clipboard source',
        operations: [
          {
            op: 'paste_cells',
            sheetId,
            start: 'A80',
            cells: [
              [
                {
                  kind: 'string',
                  rawValue: '=literal',
                  style: {
                    bold: true,
                    borderLeft: { style: 'mediumDashDotDot', color: '#123456' },
                  },
                },
                { kind: 'error', rawValue: '#DIV/0!' },
                { kind: 'date', rawValue: 45000, dateSystem: '1900', numberFormat: 'yyyy-mm-dd' },
                {
                  kind: 'number',
                  rawValue: 2,
                  formula: '=A80&$A$1',
                  calculationState: 'ready',
                  cachedValue: 2,
                },
              ],
            ],
          },
        ],
      })
      // Keep native ClipboardItem serialization deterministic inside the VS Code
      // iframe. Actual OS clipboard read/write is covered by browser acceptance.
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
      await navigate('Data!A80:D80')
      const revision = document.engine.descriptor.revision
      await frame
        .locator('.workbook-desktop-header')
        .getByRole('button', { name: /^Copy / })
        .click()
      await document.bridge!.request('view.flush', {})
      assert.equal(document.engine.descriptor.revision, revision)
      assert.equal(
        await frame
          .locator('.workbook-error')
          .allTextContents()
          .then((values) => values.join('')),
        '',
      )
      assert.ok(
        await frame.evaluate(async () =>
          (await navigator.clipboard.readText()).includes('=literal'),
        ),
        'native copy serialized the source',
      )
      await document.engine.submit({
        schemaVersion: 1,
        sessionId: document.sessionId,
        expectedRevision: revision,
        transactionId: crypto.randomUUID(),
        origin: 'manual',
        summary: 'Change source after copy',
        operations: [{ op: 'set_cell', sheetId, address: 'A80', value: 'Changed source' }],
      })
      const beforePaste = structuredClone(document.engine.state)
      await navigate('Other!D5')
      await frame
        .locator('.workbook-desktop-header')
        .getByRole('button', { name: 'Paste', exact: true })
        .click()
      await document.bridge!.request('view.flush', {})
      const afterPaste = structuredClone(document.engine.state)
      const cells = afterPaste.sheets.find((s) => s.id === otherId)!.cells
      assert.equal(cells.D5!.rawValue, '=literal')
      assert.equal(cells.D5!.formula, undefined)
      assert.equal(cells.D5!.style!.borderLeft!.style, 'mediumDashDotDot')
      assert.equal(cells.E5!.kind, 'error')
      assert.equal(cells.F5!.kind, 'date')
      assert.equal(cells.G5!.formula, '=D5&$A$1')
      assert.equal(cells.G5!.cachedValue, undefined)
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, beforePaste)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, afterPaste)
      await vscode.commands.executeCommand('workbench.action.files.save')
      const zip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const xml = await zip.file('xl/worksheets/sheet2.xml')!.async('string')
      assert.match(xml, /<c\b[^>]*r="E5"[^>]*t="e"[^>]*><v>#DIV\/0!<\/v>/)
      assert.match(xml, /<f>D5&amp;\$A\$1<\/f>/)
      assert.match(await zip.file('xl/styles.xml')!.async('string'), /mediumDashDotDot/)
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
        const read = await snapshot.read(document.engine.state.sheets[1]!.id, 'D5:G5')
        assert.equal(read.cells.find((c) => c.address === 'D5')!.rawValue, '=literal')
        assert.equal(read.cells.find((c) => c.address === 'D5')!.formula, undefined)
        assert.equal(read.cells.find((c) => c.address === 'E5')!.kind, 'error')
        assert.equal(read.cells.find((c) => c.address === 'F5')!.kind, 'date')
        assert.equal(read.cells.find((c) => c.address === 'G5')!.formula, '=D5&$A$1')
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
