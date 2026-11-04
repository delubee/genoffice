import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import JSZip from 'jszip'
import { suite, test } from 'mocha'
import { chromium, type Browser, type Frame } from '@playwright/test'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'
import { gotoReference } from './goto-helper'

type Api = { getDocument(uri: vscode.Uri): WorkbookDocument | undefined }

async function viewOf(browser: Browser, name: string) {
  let frame: Frame | undefined
  const deadline = Date.now() + 20_000
  while (!frame && Date.now() < deadline) {
    for (const context of browser.contexts())
      for (const page of context.pages()) {
        if (!(await page.title()).includes(name)) continue
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
  return frame
}

async function goTo(frame: Frame, reference: string) {
  await gotoReference(frame, reference)
}

async function settle(document: WorkbookDocument, revision: number) {
  const deadline = Date.now() + 20_000
  while (document.engine.descriptor.revision < revision && Date.now() < deadline)
    await new Promise((resolve) => setTimeout(resolve, 50))
  assert.equal(document.engine.descriptor.revision, revision)
}

async function reopen(api: Api, uri: vscode.Uri, previous: string) {
  await vscode.commands.executeCommand('workbench.action.closeActiveEditor')
  const deadline = Date.now() + 10_000
  while (api.getDocument(uri) && Date.now() < deadline)
    await new Promise((resolve) => setTimeout(resolve, 50))
  await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
  const document = api.getDocument(uri)!
  await document.whenViewReady()
  assert.notEqual(document.sessionId, previous)
  return document
}

suite('Desktop View tab UI', () => {
  test('applies saved sheet-view settings and display-only toggles through the shared View band', async function () {
    const port = process.env.GENOFFICE_TEST_CDP_PORT
    if (!port) return this.skip()
    this.timeout(150_000)
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as Api
    const name = 'view-tab-ui.xlsx'
    const uri = vscode.Uri.joinPath(vscode.workspace.workspaceFolders![0]!.uri, name)
    await vscode.workspace.fs.writeFile(uri, new Uint8Array(await manualWorkbook()))
    const browser = await chromium.connectOverCDP(`http://127.0.0.1:${port}`)
    try {
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      let document = api.getDocument(uri)!
      await document.whenViewReady()
      const frame = await viewOf(browser, name)
      await vscode.commands.executeCommand('notifications.clearAll')
      await vscode.commands.executeCommand('workbench.action.closeAuxiliaryBar')
      await vscode.commands.executeCommand('workbench.action.closeSidebar')
      await vscode.commands.executeCommand('workbench.action.closePanel')
      await frame.page().setViewportSize({ width: 1200, height: 900 })
      await goTo(frame, 'Data!B3')
      await frame.locator('#ribbon-tab-view').click()
      const header = frame.locator('.workbook-desktop-header')
      const show = header.locator('.check-column-view')
      const choose = async (menu: string, option: string) => {
        await header.getByRole('button', { name: menu, exact: true }).click()
        await frame.getByRole('option', { name: option, exact: true }).click()
      }
      const before = structuredClone(document.engine.state)
      let revision = document.engine.descriptor.revision
      await choose('Zoom', '150%')
      await settle(document, ++revision)
      await choose('Freeze Panes', 'Freeze Top Row')
      await settle(document, ++revision)
      await show.getByRole('button', { name: /Right-to-Left/ }).click()
      await settle(document, ++revision)
      await show.getByRole('button', { name: /Gridlines/ }).click()
      await settle(document, ++revision)
      const page = document.engine.state.sheets[0]!.features?.pageSetupEdits as Record<
        string,
        unknown
      >
      assert.deepEqual(
        { zoom: page.zoomScale, rtl: page.rightToLeft, grid: page.showGridlines },
        { zoom: 150, rtl: true, grid: false },
      )
      assert.equal(
        await show
          .getByRole('button', { name: /Right-to-Left/ })
          .locator('.check-box')
          .innerText(),
        '✓',
      )
      // Display-only toggles report status and create no revision.
      await show.getByRole('button', { name: /Formula Bar/ }).click()
      await frame.getByRole('status').filter({ hasText: 'Formula bar hidden' }).waitFor()
      assert.equal(await frame.locator('[data-u-comp="formula-bar"]').first().isVisible(), false)
      await show.getByRole('button', { name: /Formula Bar/ }).click()
      await show.getByRole('button', { name: /Highlight active row/ }).click()
      await header.getByRole('button', { name: /^Page Break Preview/ }).click()
      await frame
        .getByRole('status')
        .filter({ hasText: /Page Break Preview/i })
        .waitFor()
      await header.getByRole('button', { name: /^Normal/ }).click()
      assert.equal(document.engine.descriptor.revision, revision)
      if (await frame.getByRole('alert').count())
        assert.fail(await frame.getByRole('alert').first().innerText())
      const after = structuredClone(document.engine.state)
      for (let step = 0; step < 4; step++) await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, before)
      for (let step = 0; step < 4; step++) await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, after)
      await vscode.commands.executeCommand('workbench.action.files.save')
      const zip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
      const view = /<sheetView\b[^>]*>/.exec(xml)?.[0] ?? ''
      assert.match(view, /zoomScale="150"/)
      assert.match(view, /rightToLeft="1"/)
      assert.match(view, /showGridLines="0"/)
      assert.match(xml, /<pane\b[^>]*ySplit="1"/)
      document = await reopen(api, uri, document.sessionId)
      const metadata = document.engine.state.sheets[0]!.features?.originalMetadata as Record<
        string,
        unknown
      >
      assert.equal(metadata.zoomScale, 150)
      assert.equal(metadata.rightToLeft, true)
      assert.equal(metadata.showGridLines, false)
    } finally {
      if (api.getDocument(uri))
        await vscode.commands.executeCommand('workbench.action.files.save', uri)
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
      await browser.close()
    }
  })
})
