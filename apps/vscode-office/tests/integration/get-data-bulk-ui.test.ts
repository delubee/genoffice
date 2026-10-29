import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import { suite, test } from 'mocha'
import { chromium, type Browser, type Frame } from '@playwright/test'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'

type Api = { getDocument(uri: vscode.Uri): WorkbookDocument | undefined }

async function openWorkbook(browser: Browser, api: Api, name: string) {
  const uri = vscode.Uri.joinPath(vscode.workspace.workspaceFolders![0]!.uri, name)
  await vscode.workspace.fs.writeFile(uri, new Uint8Array(await manualWorkbook()))
  await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
  const document = api.getDocument(uri)!
  await document.whenViewReady()
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
  await vscode.commands.executeCommand('notifications.clearAll')
  await vscode.commands.executeCommand('workbench.action.closeAuxiliaryBar')
  await vscode.commands.executeCommand('workbench.action.closeSidebar')
  await vscode.commands.executeCommand('workbench.action.closePanel')
  await frame.page().setViewportSize({ width: 1200, height: 900 })
  await frame.locator('#ribbon-tab-data').click()
  return { uri, document, frame }
}

async function closeAndReopen(api: Api, uri: vscode.Uri, previousSession: string) {
  await vscode.commands.executeCommand('workbench.action.closeActiveEditor')
  const closeDeadline = Date.now() + 10_000
  while (api.getDocument(uri) && Date.now() < closeDeadline)
    await new Promise((resolve) => setTimeout(resolve, 50))
  assert.equal(api.getDocument(uri), undefined)
  await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
  const reopened = api.getDocument(uri)!
  await reopened.whenViewReady()
  assert.notEqual(reopened.sessionId, previousSession)
  return reopened
}

suite('Desktop Get Data bulk UI', () => {
  test('imports a CSV above the ordinary batch ceiling as one revision and one history entry', async function () {
    const port = process.env.GENOFFICE_TEST_CDP_PORT
    if (!port) return this.skip()
    this.timeout(120_000)
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as Api
    const browser = await chromium.connectOverCDP(`http://127.0.0.1:${port}`)
    let uri: vscode.Uri | undefined
    try {
      const opened = await openWorkbook(browser, api, 'get-data-bulk-csv.xlsx')
      uri = opened.uri
      const { document, frame } = opened
      await vscode.commands.executeCommand('genoffice.workbook.goto')
      const go = frame.getByRole('dialog', { name: 'Go To', exact: true })
      await go.getByRole('textbox', { name: 'Reference', exact: true }).fill('Data!C2')
      await go.getByRole('button', { name: 'Go', exact: true }).click()
      await go.waitFor({ state: 'hidden' })
      await frame.locator('#ribbon-tab-data').click()
      // 700 rows × 12 columns = 8,400 cells: five bounded stages, one commit.
      const lines = Array.from({ length: 700 }, (_, row) =>
        Array.from({ length: 12 }, (_, column) => (column === 0 ? `r${row}` : row * 100 + column)),
      )
      const before = structuredClone(document.engine.state)
      const revision = document.engine.descriptor.revision
      const chooser = frame.page().waitForEvent('filechooser')
      await frame.getByRole('button', { name: 'From Text/CSV', exact: true }).click()
      await (
        await chooser
      ).setFiles({
        name: 'bulk.csv',
        mimeType: 'text/csv',
        buffer: Buffer.from(lines.map((line) => line.join(',')).join('\n'), 'utf8'),
      })
      await frame
        .getByRole('status')
        .filter({ hasText: '700 rows × 12 columns' })
        .waitFor({ timeout: 60_000 })
      await document.bridge!.request('view.flush', {})
      if (await frame.getByRole('alert').count())
        assert.fail(await frame.getByRole('alert').first().innerText())
      assert.equal(document.engine.descriptor.revision, revision + 1)
      const after = structuredClone(document.engine.state)
      const cells = after.sheets[0]!.cells
      assert.equal(cells.C2!.rawValue, 'r0')
      assert.equal(cells.N701!.rawValue, 69911)
      assert.equal(cells.D500!.rawValue, 49801)
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, before)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, after)
      await vscode.commands.executeCommand('workbench.action.files.save')
      const reopened = await closeAndReopen(api, uri, document.sessionId)
      const snapshot = reopened.snapshot()
      try {
        const result = await snapshot.read(reopened.engine.state.sheets[0]!.id, 'C700:N701')
        const saved = new Map(result.cells.map((cell) => [cell.address, cell]))
        assert.equal(saved.get('C701')!.rawValue, 'r699')
        assert.equal(saved.get('N701')!.rawValue, 69911)
      } finally {
        snapshot.release()
      }
    } finally {
      if (uri && api.getDocument(uri))
        await vscode.commands.executeCommand('workbench.action.files.save', uri)
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      if (uri) await vscode.workspace.fs.delete(uri)
      await browser.close()
    }
  })

  test('merges a picked workbook through the host dialog as one revision and keeps it after reopen', async function () {
    const port = process.env.GENOFFICE_TEST_CDP_PORT
    if (!port) return this.skip()
    this.timeout(120_000)
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as Api
    const files = vscode.workspace.getConfiguration('files')
    await files.update('simpleDialog.enable', true, vscode.ConfigurationTarget.Global)
    const sourceUri = vscode.Uri.joinPath(
      vscode.workspace.workspaceFolders![0]!.uri,
      'merge-source.xlsx',
    )
    await vscode.workspace.fs.writeFile(sourceUri, new Uint8Array(await manualWorkbook()))
    const browser = await chromium.connectOverCDP(`http://127.0.0.1:${port}`)
    let uri: vscode.Uri | undefined
    try {
      const opened = await openWorkbook(browser, api, 'get-data-merge.xlsx')
      uri = opened.uri
      const { document, frame } = opened
      const before = structuredClone(document.engine.state)
      const revision = document.engine.descriptor.revision
      await frame.getByRole('button', { name: 'Merge Workbooks', exact: true }).click()
      // files.simpleDialog renders the host's showOpenDialog as workbench quick input.
      const input = frame.page().locator('.quick-input-widget input.input')
      await input.waitFor()
      await input.fill(sourceUri.fsPath)
      await input.press('Enter')
      await frame
        .getByRole('status')
        .filter({ hasText: 'Merged 2 sheets from 1 files' })
        .waitFor({ timeout: 60_000 })
      if (await frame.getByRole('alert').count())
        assert.fail(await frame.getByRole('alert').first().innerText())
      assert.equal(document.engine.descriptor.revision, revision + 1)
      const after = structuredClone(document.engine.state)
      assert.deepEqual(
        after.sheets.map((sheet) => sheet.name),
        ['Data', 'Other', 'Data (2)', 'Other (2)'],
      )
      const merged = after.sheets[2]!
      assert.equal(merged.cells.A1!.rawValue, '00123')
      assert.equal(merged.cells.A2!.rawValue, 12.5)
      assert.equal(merged.cells.A2!.style?.fillColor, '#CCEEFF')
      assert.equal(after.sheets[3]!.cells.B1!.formula, undefined)
      assert.deepEqual(after.sheets[0]!.cells, before.sheets[0]!.cells)
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, before)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, after)
      await vscode.commands.executeCommand('workbench.action.files.save')
      const reopened = await closeAndReopen(api, uri, document.sessionId)
      assert.deepEqual(
        reopened.engine.state.sheets.map((sheet) => sheet.name),
        ['Data', 'Other', 'Data (2)', 'Other (2)'],
      )
      const snapshot = reopened.snapshot()
      try {
        const result = await snapshot.read(reopened.engine.state.sheets[2]!.id, 'A1:A2')
        const saved = new Map(result.cells.map((cell) => [cell.address, cell]))
        assert.equal(saved.get('A1')!.rawValue, '00123')
        assert.equal(saved.get('A2')!.rawValue, 12.5)
        assert.equal(saved.get('A2')!.style?.fillColor, '#CCEEFF')
      } finally {
        snapshot.release()
      }
    } finally {
      await files.update('simpleDialog.enable', undefined, vscode.ConfigurationTarget.Global)
      if (uri && api.getDocument(uri))
        await vscode.commands.executeCommand('workbench.action.files.save', uri)
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      if (uri) await vscode.workspace.fs.delete(uri)
      await vscode.workspace.fs.delete(sourceUri)
      await browser.close()
    }
  })
})
