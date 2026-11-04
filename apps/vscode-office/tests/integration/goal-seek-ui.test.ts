import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import JSZip from 'jszip'
import { suite, test } from 'mocha'
import { chromium, type Frame } from '@playwright/test'
import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'
import { gotoReference } from './goto-helper'

suite('Desktop Goal Seek UI', () => {
  test('solves a cross-sheet target in one revision with native history and independent reopen', async function () {
    const port = process.env.GENOFFICE_TEST_CDP_PORT
    if (!port) return this.skip()
    this.timeout(90_000)
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as {
      getDocument(uri: vscode.Uri): WorkbookDocument | undefined
    }
    const uri = vscode.Uri.joinPath(vscode.workspace.workspaceFolders![0]!.uri, 'goal-seek-ui.xlsx')
    const fixture = await applyCellEditsToXlsx(await manualWorkbook(), [
      {
        sheetName: 'Data',
        row: 0,
        column: 1,
        writeValue: true,
        cell: { value: null, formula: '=A2*3+Other!A1' },
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
            if (!(await page.title()).includes('goal-seek-ui.xlsx')) continue
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
      await gotoReference(frame, 'Data!B1')
      const before = structuredClone(document.engine.state)
      const revision = document.engine.descriptor.revision
      await frame.locator('#ribbon-tab-data').click()
      await frame.getByRole('button', { name: 'What-If Analysis', exact: true }).click()
      await frame.getByRole('option', { name: 'Goal Seek', exact: true }).click()
      const dialog = frame.getByRole('dialog', { name: 'Goal Seek', exact: true })
      assert.equal(
        await dialog.getByRole('textbox', { name: 'Set cell', exact: true }).inputValue(),
        'B1',
      )
      await dialog.getByRole('textbox', { name: 'To value', exact: true }).fill('22')
      await dialog.getByRole('textbox', { name: 'By changing cell', exact: true }).fill('A2')
      await dialog.getByRole('button', { name: 'Solve', exact: true }).click()
      await dialog.getByText('Found a solution:', { exact: false }).waitFor()
      await dialog.getByRole('button', { name: 'Close', exact: true }).click()
      await document.bridge!.request('view.flush', {})
      assert.equal(document.engine.descriptor.revision, revision + 1)
      assert.ok(Math.abs(Number(document.engine.state.sheets[0]!.cells.A2!.rawValue) - 5) < 0.001)
      const after = structuredClone(document.engine.state)
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, before)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, after)
      await vscode.commands.executeCommand('workbench.action.files.save')
      const zip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      assert.match(
        await zip.file('xl/worksheets/sheet1.xml')!.async('string'),
        /<f>A2\*3\+Other!A1<\/f>/,
      )
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
      const savedCells = await reopened
        .snapshot()
        .read(reopened.engine.state.sheets[0]!.id, 'A2:C2')
      assert.ok(
        Math.abs(Number(savedCells.cells.find((cell) => cell.address === 'A2')!.rawValue) - 5) <
          0.001,
      )
      const result = await reopened.snapshot().read(reopened.engine.state.sheets[0]!.id, 'B1')
      assert.equal(result.cells[0]!.formula, '=A2*3+Other!A1')
      const other = await reopened.snapshot().read(reopened.engine.state.sheets[1]!.id, 'A1')
      assert.equal(other.cells[0]!.rawValue, 7)
    } finally {
      if (api.getDocument(uri))
        await vscode.commands.executeCommand('workbench.action.files.save', uri)
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
      await browser.close()
    }
  })
})
