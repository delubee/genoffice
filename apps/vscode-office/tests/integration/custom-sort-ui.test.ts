import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import JSZip from 'jszip'
import { suite, test } from 'mocha'
import { chromium, type Frame } from '@playwright/test'
import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'
import { gotoReference } from './goto-helper'

suite('Desktop Custom Sort UI', () => {
  test('sorts by two keys while preserving typed payloads and styles across history/save/reopen', async function () {
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
      'custom-sort-ui.xlsx',
    )
    const values = [
      ['Group', 'Amount', 'Payload'],
      ['a', 30, '00123'],
      ['a', 10, true],
      ['b', 20, 123],
    ]
    const fixture = await applyCellEditsToXlsx(
      await manualWorkbook(),
      values.flatMap((row, r) =>
        row.map((value, c) => ({
          sheetName: 'Data',
          row: r,
          column: c,
          writeValue: true,
          cell: { value },
          ...(r === 1 && c === 2
            ? { style: { bold: true, fillColor: '#CCEEFF', numberFormat: '@' } }
            : {}),
        })),
      ),
    )
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
            if (!(await page.title()).includes('custom-sort-ui.xlsx')) continue
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
      await gotoReference(frame, 'Data!A1:C4')
      const before = structuredClone(document.engine.state)
      const revision = document.engine.descriptor.revision
      await frame.locator('#ribbon-tab-data').click()
      await frame
        .locator('.workbook-shared-sort-filter')
        .getByRole('button', { name: 'Sort', exact: true })
        .click()
      await frame.getByRole('option', { name: 'Custom Sort…', exact: true }).click()
      const dialog = frame.locator('.sort-dialog')
      await dialog.locator('.sort-level').nth(1).getByRole('button').first().click()
      await frame.locator('[role=option][data-value="1"]').click()
      await dialog.getByRole('button', { name: 'OK', exact: true }).click()
      await dialog.waitFor({ state: 'hidden' })
      await document.bridge!.request('view.flush', {})
      assert.equal(document.engine.descriptor.revision, revision + 1)
      const cells = document.engine.state.sheets[0]!.cells
      assert.equal(cells.B2!.rawValue, 10)
      assert.equal(cells.C2!.rawValue, true)
      assert.equal(cells.C2!.kind, 'boolean')
      assert.equal(cells.C3!.rawValue, '00123')
      assert.equal(cells.C3!.kind, 'string')
      assert.equal(cells.C3!.style!.bold, true)
      const after = structuredClone(document.engine.state)
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, before)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, after)
      await vscode.commands.executeCommand('workbench.action.files.save')
      const zip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      assert.equal(
        await zip.file('customXml/manual-sentinel.xml')!.async('string'),
        '<preserve id="manual-fixture">unchanged</preserve>',
      )
      const previousSession = document.sessionId
      await vscode.commands.executeCommand('workbench.action.closeActiveEditor')
      const closeDeadline = Date.now() + 10_000
      while (api.getDocument(uri) && Date.now() < closeDeadline)
        await new Promise((resolve) => setTimeout(resolve, 50))
      assert.equal(api.getDocument(uri), undefined, 'the previous document must be disposed')
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      const reopened = api.getDocument(uri)!
      await reopened.whenViewReady()
      assert.notEqual(reopened.sessionId, previousSession)
      assert.equal(reopened.engine.descriptor.revision, 0)
      assert.equal(reopened.engine.descriptor.dirty, false)
      const restored = await reopened.snapshot().read(reopened.engine.state.sheets[0]!.id, 'A1:C4')
      const savedCells = new Map(restored.cells.map((cell) => [cell.address, cell]))
      assert.equal(savedCells.get('A1')!.rawValue, 'Group')
      assert.equal(savedCells.get('B2')!.rawValue, 10)
      assert.equal(savedCells.get('C2')!.rawValue, true)
      assert.equal(savedCells.get('C2')!.kind, 'boolean')
      assert.equal(savedCells.get('C3')!.rawValue, '00123')
      assert.equal(savedCells.get('C3')!.kind, 'string')
      assert.equal(savedCells.get('C3')!.style!.bold, true)
      assert.notEqual(savedCells.get('C2')!.style?.bold, true)
    } finally {
      if (api.getDocument(uri))
        await vscode.commands.executeCommand('workbench.action.files.save', uri)
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
      await browser.close()
    }
  })
})
