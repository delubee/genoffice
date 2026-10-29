import { pickOutlineCommand } from '../helpers/ribbon-menu'
import { workbookFileSchema } from '@genoffice/xlsx-gateway/shared/workbook-protocol'
import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import JSZip from 'jszip'
import { suite, test } from 'mocha'
import { chromium, expect, type Frame } from '@playwright/test'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'

suite('Desktop Outline UI', () => {
  test('groups and ungroups nested axes with summary settings, native history and reopen', async function () {
    const port = process.env.GENOFFICE_TEST_CDP_PORT
    if (!port) return this.skip()
    this.timeout(90_000)
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as {
      getDocument(uri: vscode.Uri): WorkbookDocument | undefined
    }
    const uri = vscode.Uri.joinPath(vscode.workspace.workspaceFolders![0]!.uri, 'outline-ui.xlsx')
    await vscode.workspace.fs.writeFile(uri, await manualWorkbook())
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
            if (!(await page.title()).includes('outline-ui.xlsx')) continue
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
      await go.getByRole('textbox', { name: 'Reference', exact: true }).fill('Data!A2:A3')
      await go.getByRole('button', { name: 'Go', exact: true }).click()
      await go.waitFor({ state: 'hidden' })
      const before = structuredClone(document.engine.state)
      const revision = document.engine.descriptor.revision
      await frame.locator('#ribbon-tab-data').click()
      for (const name of ['Group Rows', 'Group Rows', 'Ungroup Rows']) {
        await pickOutlineCommand(frame, name)
        await document.bridge!.request('view.flush', {})
      }
      await pickOutlineCommand(frame, '✓ Summary rows below detail')
      await document.bridge!.request('view.flush', {})
      await frame.getByRole('button', { name: 'Group', exact: true }).click()
      await expect(
        frame.getByRole('option', { name: 'Summary rows below detail', exact: true }),
      ).toBeVisible()
      await frame.getByRole('button', { name: 'Group', exact: true }).press('Escape')
      for (const name of ['Group Columns', 'Group Columns', 'Ungroup Columns']) {
        await pickOutlineCommand(frame, name)
        await document.bridge!.request('view.flush', {})
      }
      await pickOutlineCommand(frame, '✓ Summary columns to right of detail')
      await document.bridge!.request('view.flush', {})
      await frame.getByRole('button', { name: 'Group', exact: true }).click()
      await expect(
        frame.getByRole('option', { name: 'Summary columns to right of detail', exact: true }),
      ).toBeVisible()
      await frame.getByRole('button', { name: 'Group', exact: true }).press('Escape')
      assert.equal(document.engine.descriptor.revision, revision + 8)
      const edits = document.engine.state.sheets[0]!.features!.outlineEdits as { level?: number }[]
      assert.deepEqual(
        edits.slice(0, 3).map((edit) => edit.level),
        [1, 2, 1],
      )
      assert.deepEqual(
        edits.slice(4, 7).map((edit) => edit.level),
        [1, 2, 1],
      )
      assert.deepEqual(document.engine.state.sheets[0]!.cells, before.sheets[0]!.cells)
      const after = structuredClone(document.engine.state)
      for (let i = 0; i < 8; i++) await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, before)
      for (let i = 0; i < 8; i++) await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, after)
      await vscode.commands.executeCommand('workbench.action.files.save')
      const zip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
      assert.match(xml, /<row\b[^>]*r="2"[^>]*outlineLevel="1"/)
      assert.match(xml, /<row\b[^>]*r="3"[^>]*outlineLevel="1"/)
      assert.match(xml, /summaryBelow="0"/)
      assert.match(xml, /summaryRight="0"/)
      assert.match(xml, /<col\b[^>]*min="1"[^>]*max="1"[^>]*outlineLevel="1"/)
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
      const snapshot = reopened.snapshot()
      try {
        const savedCells = await snapshot.read(reopened.engine.state.sheets[0]!.id, 'A2:C2')
        assert.equal(savedCells.cells.find((cell) => cell.address === 'A2')!.rawValue, 12.5)
      } finally {
        snapshot.release()
      }
      let importedFrame: Frame | undefined
      const reopenDeadline = Date.now() + 20_000
      while (!importedFrame && Date.now() < reopenDeadline) {
        for (const context of browser.contexts())
          for (const page of context.pages()) {
            if (!(await page.title()).includes('outline-ui.xlsx')) continue
            for (const candidate of page.frames())
              if (
                await candidate
                  .locator('#ribbon-tab-data')
                  .count()
                  .catch(() => 0)
              )
                importedFrame = candidate
          }
        if (!importedFrame) await new Promise((resolve) => setTimeout(resolve, 100))
      }
      assert.ok(importedFrame)
      await importedFrame.locator('#ribbon-tab-data').click()
      await importedFrame.getByRole('button', { name: 'Group', exact: true }).click()
      await expect(
        importedFrame.getByRole('option', { name: 'Summary rows below detail', exact: true }),
      ).toBeVisible()
      await expect(
        importedFrame.getByRole('option', {
          name: 'Summary columns to right of detail',
          exact: true,
        }),
      ).toBeVisible()
      await importedFrame.getByRole('button', { name: 'Group', exact: true }).press('Escape')
      const metadata = workbookFileSchema.parse(reopened.metadata)
      assert.equal(metadata.sheets[0]!.outlineLevelRow, 1)
      assert.equal(metadata.sheets[0]!.outlineSummaryBelow, false)
      assert.equal(metadata.sheets[0]!.outlineLevelCol, 1)
      assert.equal(metadata.sheets[0]!.outlineSummaryRight, false)
    } finally {
      if (api.getDocument(uri))
        await vscode.commands.executeCommand('workbench.action.files.save', uri)
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
      await browser.close()
    }
  })
})
