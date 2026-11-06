import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import JSZip from 'jszip'
import { suite, test } from 'mocha'
import { chromium, type Frame } from '@playwright/test'
import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'
import { gotoReference } from './goto-helper'

suite('Desktop Text to Columns UI', () => {
  test('splits sparse non-A rows through the shared menu and preserves references across history and reopen', async function () {
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
      'text-to-columns-ui.xlsx',
    )
    const fixture = await applyCellEditsToXlsx(await manualWorkbook(), [
      { sheetName: 'Data', row: 3, column: 1, writeValue: true, cell: { value: 'left,right' } },
      { sheetName: 'Data', row: 5, column: 1, writeValue: true, cell: { value: 'north,south' } },
      {
        sheetName: 'Data',
        row: 0,
        column: 0,
        writeValue: true,
        cell: { value: 0, formula: '=D4' },
      },
    ])
    await vscode.workspace.fs.writeFile(uri, fixture.buffer)
    const browser = await chromium.connectOverCDP(`http://127.0.0.1:${port}`)
    try {
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      const document = api.getDocument(uri)!
      await document.whenViewReady()
      let frame: Frame | undefined
      const deadline = Date.now() + 20_000
      while (!frame && Date.now() < deadline) {
        for (const context of browser.contexts())
          for (const page of context.pages()) {
            if (!(await page.title()).includes('text-to-columns-ui.xlsx')) continue
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
      await frame.page().setViewportSize({ width: 1200, height: 900 })
      await gotoReference(frame, 'Data!B4:B6')
      const before = structuredClone(document.engine.state)
      const revision = document.engine.descriptor.revision
      await frame.locator('#ribbon-tab-data').click()
      await frame.getByRole('button', { name: 'Text to Columns', exact: true }).click()
      await frame.getByRole('option', { name: 'Split by Comma', exact: true }).click()
      await document.bridge!.request('view.flush', {})
      if (await frame.getByRole('alert').count())
        assert.fail(await frame.getByRole('alert').first().innerText())
      assert.equal(document.engine.descriptor.revision, revision + 1)
      const after = structuredClone(document.engine.state)
      assert.deepEqual(after.structuralLog, before.structuralLog)
      assert.equal(after.sheets[0]!.cells.B4!.rawValue, 'left')
      assert.equal(after.sheets[0]!.cells.C4!.rawValue, 'right')
      assert.equal(after.sheets[0]!.cells.B6!.rawValue, 'north')
      assert.equal(after.sheets[0]!.cells.C6!.rawValue, 'south')
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, before)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, after)
      await vscode.commands.executeCommand('workbench.action.files.save')
      const zip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      assert.match(await zip.file('xl/worksheets/sheet1.xml')!.async('string'), /<f>D4<\/f>/)
      assert.equal(
        await zip.file('customXml/manual-sentinel.xml')!.async('string'),
        '<preserve id="manual-fixture">unchanged</preserve>',
      )
      const previousSession = document.sessionId
      await vscode.commands.executeCommand('workbench.action.closeActiveEditor')
      const closeDeadline = Date.now() + 10_000
      while (api.getDocument(uri) && Date.now() < closeDeadline)
        await new Promise((resolve) => setTimeout(resolve, 50))
      assert.equal(api.getDocument(uri), undefined)
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      const reopened = api.getDocument(uri)!
      await reopened.whenViewReady()
      assert.notEqual(reopened.sessionId, previousSession)
      const snapshot = reopened.snapshot()
      try {
        const result = await snapshot.read(reopened.engine.state.sheets[0]!.id, 'A1:D6')
        const cells = new Map(result.cells.map((cell) => [cell.address, cell]))
        assert.equal(cells.get('A1')!.formula, '=D4')
        assert.equal(cells.get('C4')!.rawValue, 'right')
        assert.equal(cells.get('C6')!.rawValue, 'south')
        assert.equal(cells.get('B5')?.rawValue ?? null, null)
        assert.equal(cells.get('A2')!.rawValue, 12.5)
        assert.equal(cells.get('A2')!.style?.fillColor, '#CCEEFF')
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
