import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import JSZip from 'jszip'
import { suite, test } from 'mocha'
import { chromium, type Browser, type Frame } from '@playwright/test'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'

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
  await vscode.commands.executeCommand('genoffice.workbook.goto')
  const go = frame.getByRole('dialog', { name: 'Go To', exact: true })
  await go.getByRole('textbox', { name: 'Reference', exact: true }).fill(reference)
  await go.getByRole('button', { name: 'Go', exact: true }).click()
  await go.waitFor({ state: 'hidden' })
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

suite('Desktop Insert icons and checkbox UI', () => {
  test('inserts a rendered icon picture and a checkbox rule through the shared Insert band', async function () {
    const port = process.env.GENOFFICE_TEST_CDP_PORT
    if (!port) return this.skip()
    this.timeout(150_000)
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as Api
    const name = 'insert-icons-ui.xlsx'
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
      await goTo(frame, 'Data!D4')
      await frame.locator('#ribbon-tab-insert').click()
      const header = frame.locator('.workbook-desktop-header')
      let revision = document.engine.descriptor.revision
      await header
        .getByRole('button', { name: /^Icons/ })
        .first()
        .click()
      const icons = frame.getByRole('dialog', { name: 'Icons', exact: true })
      await icons.locator('.icons-grid button').first().click()
      await settle(document, ++revision)
      await icons.waitFor({ state: 'hidden' }).catch(async () => {
        await icons.getByRole('button', { name: 'Close', exact: true }).click()
      })
      const visuals = document.engine.state.sheets[0]!.features?.visuals as { kind: string }[]
      assert.ok(visuals?.some((visual) => visual.kind === 'image'))
      await goTo(frame, 'Data!F2:F3')
      await frame.locator('#ribbon-tab-insert').click()
      await header
        .getByRole('button', { name: /^Checkbox/ })
        .first()
        .click()
      await settle(document, ++revision)
      if (await frame.getByRole('alert').count())
        assert.fail(await frame.getByRole('alert').first().innerText())
      await vscode.commands.executeCommand('workbench.action.files.save')
      const zip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      assert.ok(Object.keys(zip.files).some((part) => /^xl\/media\/.+\.png$/.test(part)))
      assert.match(
        await zip.file('xl/worksheets/sheet1.xml')!.async('string'),
        /<dataValidation\b[^>]*sqref="F2:F3"/,
      )
      document = await reopen(api, uri, document.sessionId)
    } finally {
      if (api.getDocument(uri))
        await vscode.commands.executeCommand('workbench.action.files.save', uri)
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
      await browser.close()
    }
  })
})
