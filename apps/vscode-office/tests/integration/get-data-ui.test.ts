import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import JSZip from 'jszip'
import { suite, test } from 'mocha'
import { chromium, type Frame } from '@playwright/test'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'
import { gotoReference } from './goto-helper'

suite('Desktop Get Data UI', () => {
  test('imports a CSV through the shared group as one revision and keeps it across history and reopen', async function () {
    const port = process.env.GENOFFICE_TEST_CDP_PORT
    if (!port) return this.skip()
    this.timeout(90_000)
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as {
      getDocument(uri: vscode.Uri): WorkbookDocument | undefined
    }
    const uri = vscode.Uri.joinPath(vscode.workspace.workspaceFolders![0]!.uri, 'get-data-ui.xlsx')
    await vscode.workspace.fs.writeFile(uri, new Uint8Array(await manualWorkbook()))
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
            if (!(await page.title()).includes('get-data-ui.xlsx')) continue
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
      await gotoReference(frame, 'Data!E3:F4')
      await frame.locator('#ribbon-tab-data').click()

      // Refresh All mirrors the desktop message for a workbook without pivots.
      const revision = document.engine.descriptor.revision
      await frame.getByRole('button', { name: 'Refresh All', exact: true }).click()
      await frame
        .getByRole('alert')
        .filter({ hasText: /no pivot tables/i })
        .waitFor()
      assert.equal(document.engine.descriptor.revision, revision)

      // A cancelled pick writes nothing.
      let chooser = frame.page().waitForEvent('filechooser')
      await frame.getByRole('button', { name: 'From Text/CSV', exact: true }).click()
      await (await chooser).setFiles([])
      await new Promise((resolve) => setTimeout(resolve, 300))
      assert.equal(document.engine.descriptor.revision, revision)

      const before = structuredClone(document.engine.state)
      chooser = frame.page().waitForEvent('filechooser')
      await frame.getByRole('button', { name: 'From Text/CSV', exact: true }).click()
      await (
        await chooser
      ).setFiles({
        name: 'orders.csv',
        mimeType: 'text/csv',
        buffer: Buffer.from('﻿item;qty;note\nalpha;3;"x;y"\nbeta;4.5;\n', 'utf8'),
      })
      await frame.getByRole('status').filter({ hasText: '3 rows × 3 columns' }).waitFor()
      await document.bridge!.request('view.flush', {})
      if (await frame.getByRole('alert').count())
        assert.fail(await frame.getByRole('alert').first().innerText())
      assert.equal(document.engine.descriptor.revision, revision + 1)
      const after = structuredClone(document.engine.state)
      const cells = after.sheets[0]!.cells
      assert.equal(cells.E3!.rawValue, 'item')
      assert.equal(cells.F4!.rawValue, 3)
      assert.equal(cells.G4!.rawValue, 'x;y')
      assert.equal(cells.F5!.rawValue, 4.5)
      assert.equal(cells.G5?.rawValue ?? null, null)
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
      assert.equal(api.getDocument(uri), undefined)
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      const reopened = api.getDocument(uri)!
      await reopened.whenViewReady()
      assert.notEqual(reopened.sessionId, previousSession)
      const snapshot = reopened.snapshot()
      try {
        const result = await snapshot.read(reopened.engine.state.sheets[0]!.id, 'A1:G5')
        const saved = new Map(result.cells.map((cell) => [cell.address, cell]))
        assert.equal(saved.get('E3')!.rawValue, 'item')
        assert.equal(saved.get('F5')!.rawValue, 4.5)
        assert.equal(saved.get('G4')!.rawValue, 'x;y')
        assert.equal(saved.get('A2')!.rawValue, 12.5)
        assert.equal(saved.get('A2')!.style?.fillColor, '#CCEEFF')
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
