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

suite('Desktop Review tab UI', () => {
  test('protects ranges, sheet and workbook through the shared Review band and verifies the password after reopen', async function () {
    const port = process.env.GENOFFICE_TEST_CDP_PORT
    if (!port) return this.skip()
    this.timeout(240_000)
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as Api
    const name = 'review-tab-ui.xlsx'
    const uri = vscode.Uri.joinPath(vscode.workspace.workspaceFolders![0]!.uri, name)
    await vscode.workspace.fs.writeFile(uri, new Uint8Array(await manualWorkbook()))
    const browser = await chromium.connectOverCDP(`http://127.0.0.1:${port}`)
    try {
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      let document = api.getDocument(uri)!
      await document.whenViewReady()
      let frame = await viewOf(browser, name)
      await vscode.commands.executeCommand('notifications.clearAll')
      await vscode.commands.executeCommand('workbench.action.closeAuxiliaryBar')
      await vscode.commands.executeCommand('workbench.action.closeSidebar')
      await vscode.commands.executeCommand('workbench.action.closePanel')
      await frame.page().setViewportSize({ width: 1200, height: 900 })
      await goTo(frame, 'Data!B2')
      await frame.locator('#ribbon-tab-review').click()
      let header = frame.locator('.workbook-desktop-header')
      const before = structuredClone(document.engine.state)
      let revision = document.engine.descriptor.revision

      await header.getByRole('button', { name: /^Workbook Statistics/ }).click()
      await frame
        .getByRole('status')
        .filter({ hasText: /2 sheet/ })
        .waitFor()
      assert.equal(document.engine.descriptor.revision, revision)

      await header.getByRole('button', { name: /^Allow Edit Ranges/ }).click()
      const ranges = frame.getByRole('dialog', { name: 'Allow Edit Ranges', exact: true })
      await ranges.getByRole('textbox', { name: 'Title', exact: true }).fill('Inputs')
      await ranges.getByRole('textbox', { name: 'Range', exact: true }).fill('B2:C3')
      await ranges.getByRole('button', { name: 'Add', exact: true }).click()
      await ranges.getByRole('button', { name: 'OK', exact: true }).click()
      await settle(document, ++revision)

      await header.getByRole('button', { name: /^Protect Sheet/ }).click()
      const protect = frame.getByRole('dialog', { name: 'Protect Sheet', exact: true })
      const secrets = protect.locator('input[type="password"]')
      await secrets.nth(0).fill('open sesame')
      await secrets.nth(1).fill('open sesame')
      await protect.getByRole('button', { name: 'OK', exact: true }).click()
      await protect.waitFor({ state: 'hidden', timeout: 60_000 })
      await settle(document, ++revision)
      await header.getByRole('button', { name: /^Unprotect Sheet/ }).waitFor()

      await header.getByRole('button', { name: /^Protect Workbook/ }).click()
      await settle(document, ++revision)
      await header.getByRole('button', { name: /^Unprotect Workbook/ }).waitFor()
      if (await frame.getByRole('alert').count())
        assert.fail(await frame.getByRole('alert').first().innerText())
      const after = structuredClone(document.engine.state)
      for (let step = 0; step < 3; step++) await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, before)
      for (let step = 0; step < 3; step++) await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, after)

      await vscode.commands.executeCommand('workbench.action.files.save')
      const zip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
      assert.match(sheetXml, /<sheetProtection\b[^>]*hashValue="/)
      assert.match(sheetXml, /<protectedRange\b[^>]*sqref="B2:C3"/)
      assert.match(
        await zip.file('xl/workbook.xml')!.async('string'),
        /<workbookProtection\b[^>]*lockStructure="1"/,
      )

      document = await reopen(api, uri, document.sessionId)
      frame = await viewOf(browser, name)
      await frame.locator('#ribbon-tab-review').click()
      header = frame.locator('.workbook-desktop-header')
      revision = document.engine.descriptor.revision
      await header.getByRole('button', { name: /^Unprotect Sheet/ }).click()
      const unprotect = frame.getByRole('dialog', { name: 'Unprotect Sheet', exact: true })
      await unprotect.locator('input[type="password"]').fill('wrong')
      await unprotect.getByRole('button', { name: 'OK', exact: true }).click()
      await unprotect
        .getByText('The password you supplied is not correct.')
        .waitFor({ timeout: 60_000 })
      assert.equal(document.engine.descriptor.revision, revision)
      await unprotect.locator('input[type="password"]').fill('open sesame')
      await unprotect.getByRole('button', { name: 'OK', exact: true }).click()
      await unprotect.waitFor({ state: 'hidden', timeout: 60_000 })
      await settle(document, revision + 1)
      assert.equal(document.engine.state.sheets[0]!.protected, false)
    } finally {
      if (api.getDocument(uri))
        await vscode.commands.executeCommand('workbench.action.files.save', uri)
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
      await browser.close()
    }
  })

  test('edits notes through the shared Notes menu and Univer note popup with history, save and reopen', async function () {
    const port = process.env.GENOFFICE_TEST_CDP_PORT
    if (!port) return this.skip()
    this.timeout(180_000)
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as Api
    const name = 'review-notes-ui.xlsx'
    const uri = vscode.Uri.joinPath(vscode.workspace.workspaceFolders![0]!.uri, name)
    await vscode.workspace.fs.writeFile(uri, new Uint8Array(await manualWorkbook()))
    const browser = await chromium.connectOverCDP(`http://127.0.0.1:${port}`)
    try {
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      let document = api.getDocument(uri)!
      await document.whenViewReady()
      let frame = await viewOf(browser, name)
      await vscode.commands.executeCommand('notifications.clearAll')
      await vscode.commands.executeCommand('workbench.action.closeAuxiliaryBar')
      await vscode.commands.executeCommand('workbench.action.closeSidebar')
      await vscode.commands.executeCommand('workbench.action.closePanel')
      await frame.page().setViewportSize({ width: 1200, height: 900 })
      await goTo(frame, 'Data!C2')
      await frame.locator('#ribbon-tab-review').click()
      const header = frame.locator('.workbook-desktop-header')
      const before = structuredClone(document.engine.state)
      const revision = document.engine.descriptor.revision
      await header.getByRole('button', { name: 'Notes', exact: true }).click()
      await frame.getByRole('option', { name: 'New Note', exact: true }).click()
      const editor = frame.locator('textarea').last()
      await editor.waitFor()
      await editor.fill('Check this total')
      await new Promise((resolve) => setTimeout(resolve, 600))
      assert.equal(document.engine.descriptor.revision, revision)
      await goTo(frame, 'Data!E8')
      await settle(document, revision + 1)
      const edits = document.engine.state.sheets[0]!.features?.noteEdits as Record<string, string>
      assert.equal(edits.C2, 'Check this total')
      const after = structuredClone(document.engine.state)
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, before)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, after)
      await vscode.commands.executeCommand('workbench.action.files.save')
      document = await reopen(api, uri, document.sessionId)
      const comments = (
        document.engine.state.sheets[0]!.features?.originalMetadata as {
          comments?: { row: number; column: number; text: string }[]
        }
      ).comments
      assert.ok(
        comments?.some(
          (comment) =>
            comment.row === 1 && comment.column === 2 && comment.text === 'Check this total',
        ),
      )
      frame = await viewOf(browser, name)
      await goTo(frame, 'Data!C2')
      await frame.locator('#ribbon-tab-review').click()
      const reopened = document.engine.descriptor.revision
      await frame
        .locator('.workbook-desktop-header')
        .getByRole('button', { name: /^Delete/ })
        .first()
        .click()
      await settle(document, reopened + 1)
      assert.equal(
        (document.engine.state.sheets[0]!.features?.noteEdits as Record<string, null>).C2,
        null,
      )
    } finally {
      if (api.getDocument(uri))
        await vscode.commands.executeCommand('workbench.action.files.save', uri)
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
      await browser.close()
    }
  })

  test('posts, replies to, resolves and saves threaded comments through the shared Comments pane', async function () {
    const port = process.env.GENOFFICE_TEST_CDP_PORT
    if (!port) return this.skip()
    this.timeout(180_000)
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as Api
    const name = 'review-comments-ui.xlsx'
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
      await goTo(frame, 'Data!B2')
      await frame.locator('#ribbon-tab-review').click()
      const header = frame.locator('.workbook-desktop-header')
      const before = structuredClone(document.engine.state)
      let revision = document.engine.descriptor.revision
      await header.getByRole('button', { name: /^New Comment/ }).click()
      const pane = frame.locator('aside.threaded-comments-pane')
      await pane.waitFor()
      await pane.locator('textarea').first().fill('Is this total right?')
      await pane.getByRole('button', { name: 'Post', exact: true }).click()
      await settle(document, ++revision)
      const threads = () =>
        document.engine.state.sheets[0]!.features?.threadEdits as Record<
          string,
          { text: string; done: boolean; replies: { text: string }[] }
        >
      assert.equal(threads().B2!.text, 'Is this total right?')
      await pane.locator('textarea').last().fill('Yes, checked.')
      await pane.getByRole('button', { name: 'Reply', exact: true }).click()
      await settle(document, ++revision)
      assert.equal(threads().B2!.replies[0]!.text, 'Yes, checked.')
      await pane.getByRole('button', { name: 'Resolve', exact: true }).click()
      await settle(document, ++revision)
      assert.equal(threads().B2!.done, true)
      const after = structuredClone(document.engine.state)
      for (let step = 0; step < 3; step++) await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, before)
      for (let step = 0; step < 3; step++) await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, after)
      await vscode.commands.executeCommand('workbench.action.files.save')
      document = await reopen(api, uri, document.sessionId)
      const comment = (
        document.engine.state.sheets[0]!.features?.originalMetadata as {
          comments: {
            row: number
            column: number
            thread?: { done: boolean; replies: unknown[] }
          }[]
        }
      ).comments.find((entry) => entry.row === 1 && entry.column === 1)
      assert.equal(comment?.thread?.done, true)
      assert.equal(comment?.thread?.replies.length, 1)
    } finally {
      if (api.getDocument(uri))
        await vscode.commands.executeCommand('workbench.action.files.save', uri)
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
      await browser.close()
    }
  })
})
