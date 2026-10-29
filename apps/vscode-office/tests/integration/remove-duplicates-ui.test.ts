import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import JSZip from 'jszip'
import { suite, test } from 'mocha'
import { chromium, expect, type Frame } from '@playwright/test'
import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'

suite('Desktop Remove Duplicates UI', () => {
  test('uses the shared dialog to compact duplicate rows atomically and preserves formulas, types and source content', async function () {
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
      'remove-duplicates-ui.xlsx',
    )
    const values = [
      ['Name', 'Quantity', 'Flag'],
      ['North', 7, true],
      ['north', 7, true],
      ['South', 12.5, false],
      ['south', 12.5, false],
    ]
    const fixture = await applyCellEditsToXlsx(
      await manualWorkbook(),
      [
        ...values.flatMap((row, r) =>
          row.map((value, c) => ({
            sheetName: 'Data',
            row: r,
            column: c,
            writeValue: true,
            cell: {
              value,
              ...(r === 1 && c === 1 ? { formula: '=Other!A1' } : {}),
              ...(r === 3 && c === 2 ? { formula: '=FALSE()' } : {}),
            },
            ...(r === 2 && c === 0 ? { style: { fillColor: '#CCEEFF' } } : {}),
          })),
        ),
        {
          sheetName: 'Data',
          row: 1,
          column: 4,
          writeValue: true,
          cell: { value: 45658 },
          style: { numberFormat: 'yyyy-mm-dd' },
        },
      ],
      [],
      [],
      undefined,
      [],
      [],
      [],
      [],
      [],
      null,
      [],
      [],
      [
        // Known source results: Other!A1 = 7 and FALSE() = false.
        {
          sheetName: 'Data',
          cells: [
            { row: 1, column: 1, value: 7 },
            { row: 3, column: 2, value: false },
          ],
        },
      ],
    )
    await vscode.workspace.fs.writeFile(uri, fixture.buffer)
    const browser = await chromium.connectOverCDP(`http://127.0.0.1:${port}`)
    try {
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      const document = api.getDocument(uri)!
      await document.whenViewReady()
      const initial = document.snapshot()
      try {
        const source = await initial.read(document.engine.state.sheets[0]!.id, 'B2:C4')
        assert.equal(source.cells.find((cell) => cell.address === 'B2')!.calculationState, 'ready')
        assert.equal(source.cells.find((cell) => cell.address === 'B2')!.cachedValue, 7)
        assert.equal(source.cells.find((cell) => cell.address === 'C4')!.calculationState, 'ready')
        assert.equal(source.cells.find((cell) => cell.address === 'C4')!.cachedValue, false)
      } finally {
        initial.release()
      }
      let frame: Frame | undefined
      const deadline = Date.now() + 20_000
      while (!frame && Date.now() < deadline) {
        for (const context of browser.contexts())
          for (const page of context.pages()) {
            if (!(await page.title()).includes('remove-duplicates-ui.xlsx')) continue
            for (const candidate of page.frames())
              if (
                await candidate
                  .locator('#ribbon-tab-data')
                  .count()
                  .catch(() => 0)
              )
                frame = candidate
          }
        if (!frame) await new Promise((resolve) => setTimeout(resolve, 100))
      }
      assert.ok(frame)
      await vscode.commands.executeCommand('notifications.clearAll')
      await vscode.commands.executeCommand('workbench.action.closeAuxiliaryBar')
      await vscode.commands.executeCommand('workbench.action.closeSidebar')
      await vscode.commands.executeCommand('workbench.action.closePanel')
      await vscode.commands.executeCommand('genoffice.workbook.goto')
      const go = frame.getByRole('dialog', { name: 'Go To', exact: true })
      await go.getByRole('textbox', { name: 'Reference', exact: true }).fill('Data!A1:C5')
      await go.getByRole('button', { name: 'Go', exact: true }).click()
      await go.waitFor({ state: 'hidden' })
      const before = structuredClone(document.engine.state)
      const revision = document.engine.descriptor.revision
      await frame.locator('#ribbon-tab-data').click()
      await frame.getByRole('button', { name: 'Remove Duplicates', exact: true }).click()
      const dialog = frame.getByRole('dialog', { name: 'Remove Duplicates', exact: true })
      await expect(
        dialog.getByRole('checkbox', { name: 'My data has headers', exact: true }),
      ).toBeChecked()
      await dialog.getByRole('button', { name: 'OK', exact: true }).click()
      await document.bridge!.request('view.flush', {})
      if (await dialog.getByRole('alert').count())
        assert.fail(await dialog.getByRole('alert').innerText())
      await dialog.waitFor({ state: 'hidden' })
      assert.equal(document.engine.descriptor.revision, revision + 1)
      await expect(frame.getByRole('status')).toContainText('Removed 2 duplicate row(s).')
      const after = structuredClone(document.engine.state)
      const sheet = after.sheets[0]!
      assert.equal(sheet.cells.A3!.rawValue, 'South')
      assert.equal(sheet.cells.B3!.rawValue, 12.5)
      assert.equal(sheet.cells.C3!.rawValue, false)
      assert.equal(sheet.cells.C3!.kind, 'boolean')
      assert.equal(sheet.cells.C3!.formula, undefined)
      for (const address of ['A4', 'B4', 'C4', 'A5', 'B5', 'C5'])
        assert.equal(sheet.cells[address]!.rawValue, null)
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, before)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, after)
      await vscode.commands.executeCommand('workbench.action.files.save')
      const zip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
      assert.match(xml, /<c\b[^>]*r="B2"[^>]*>[\s\S]*?<f>Other!A1<\/f>/)
      assert.doesNotMatch(xml, /<f>FALSE\(\)<\/f>/)
      assert.equal(
        await zip.file('customXml/manual-sentinel.xml')!.async('string'),
        '<preserve id="manual-fixture">unchanged</preserve>',
      )
      const previousSession = document.sessionId
      await vscode.commands.executeCommand('workbench.action.closeActiveEditor')
      const deadlineClose = Date.now() + 10000
      while (api.getDocument(uri) && Date.now() < deadlineClose)
        await new Promise((resolve) => setTimeout(resolve, 50))
      assert.equal(api.getDocument(uri), undefined)
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      const reopened = api.getDocument(uri)!
      await reopened.whenViewReady()
      assert.notEqual(reopened.sessionId, previousSession)
      assert.equal(reopened.engine.descriptor.revision, 0)
      assert.equal(reopened.engine.descriptor.dirty, false)
      const snapshot = reopened.snapshot()
      try {
        const result = await snapshot.read(reopened.engine.state.sheets[0]!.id, 'A1:E5')
        const cells = new Map(result.cells.map((cell) => [cell.address, cell]))
        assert.equal(cells.get('A1')!.rawValue, 'Name')
        assert.equal(cells.get('B2')!.formula, '=Other!A1')
        assert.equal(cells.get('A3')!.rawValue, 'South')
        assert.equal(cells.get('A3')!.style?.fillColor, '#CCEEFF')
        assert.equal(cells.get('B3')!.rawValue, 12.5)
        assert.equal(cells.get('C3')!.rawValue, false)
        assert.equal(cells.get('C3')!.kind, 'boolean')
        assert.equal(cells.get('C3')!.formula, undefined)
        for (const address of ['A4', 'B4', 'C4', 'A5', 'B5', 'C5'])
          assert.equal(cells.get(address)!.rawValue, null)
        assert.equal(cells.get('E2')!.rawValue, 45658)
        assert.equal(cells.get('E2')!.kind, 'date')
        assert.equal(cells.get('E2')!.numberFormat, 'yyyy-mm-dd')
        const other = await snapshot.read(reopened.engine.state.sheets[1]!.id, 'A1')
        assert.equal(other.cells[0]!.rawValue, 7)
      } finally {
        snapshot.release()
      }
    } finally {
      if (api.getDocument(uri))
        await vscode.commands.executeCommand('workbench.action.files.save', uri)
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
      await browser.close()
    }
  })
})
