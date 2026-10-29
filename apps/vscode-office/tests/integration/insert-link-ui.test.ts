import * as assert from 'node:assert/strict'
import { randomUUID } from 'node:crypto'
import * as vscode from 'vscode'
import JSZip from 'jszip'
import { suite, test } from 'mocha'
import { chromium, type Frame } from '@playwright/test'
import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'

suite('Desktop Insert Link UI', () => {
  test('edits original links through the shared dialog with cancel, validation, save and native history', async function () {
    const port = process.env.GENOFFICE_TEST_CDP_PORT
    if (!port) return this.skip()
    this.timeout(90_000)
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as {
      getDocument(uri: vscode.Uri): WorkbookDocument | undefined
      revertForTests(uri: vscode.Uri): Promise<void>
    }
    const uri = vscode.Uri.joinPath(
      vscode.workspace.workspaceFolders![0]!.uri,
      'insert-link-ui.xlsx',
    )
    const original = await applyCellEditsToXlsx(
      await manualWorkbook(),
      [],
      [],
      [],
      undefined,
      [],
      [
        {
          sheetName: 'Data',
          edits: [{ row: 1, column: 0, target: 'https://example.org/original' }],
        },
      ],
    )
    await vscode.workspace.fs.writeFile(uri, original.buffer)
    const browser = await chromium.connectOverCDP(`http://127.0.0.1:${port}`)
    try {
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      const document = api.getDocument(uri)!
      await document.whenViewReady()
      let frame: Frame | undefined
      const started = Date.now()
      while (!frame && Date.now() - started < 20_000) {
        for (const context of browser.contexts())
          for (const page of context.pages()) {
            if (!(await page.title()).includes('insert-link-ui.xlsx')) continue
            for (const candidate of page.frames())
              if (
                await candidate
                  .locator('.workbook-desktop-header')
                  .count()
                  .catch(() => 0)
              )
                frame = candidate
          }
        if (!frame) await new Promise((resolve) => setTimeout(resolve, 100))
      }
      assert.ok(frame)
      await vscode.commands.executeCommand('genoffice.workbook.goto')
      const goTo = frame.getByRole('dialog', { name: 'Go To', exact: true })
      await goTo.getByRole('textbox', { name: 'Reference', exact: true }).fill('Data!A2')
      await goTo.getByRole('button', { name: 'Go', exact: true }).click()
      await goTo.waitFor({ state: 'hidden' })
      await frame.getByRole('tab', { name: 'Insert', exact: true }).click()
      const insert = frame.locator('.workbook-desktop-header')
      assert.equal(await insert.locator('.ribbon-group').count(), 10)
      const before = structuredClone(document.engine.state)
      const beforeRevision = document.engine.descriptor.revision
      const open = () => insert.locator('[data-tip="Link"]').click()
      await open()
      const edit = frame.getByRole('dialog', { name: 'Edit Link', exact: true })
      await edit.waitFor({ state: 'visible' })
      assert.equal(await edit.getByRole('textbox').inputValue(), 'https://example.org/original')
      await edit.getByRole('textbox').fill('https://example.org/cancelled')
      await edit.getByRole('button', { name: 'Cancel', exact: true }).click()
      assert.deepEqual(document.engine.state, before)
      await open()
      await edit.getByRole('textbox').fill('javascript:alert(1)')
      await edit.getByRole('button', { name: 'OK', exact: true }).click()
      await frame.getByRole('alert').filter({ hasText: 'Invalid hyperlink' }).waitFor()
      assert.equal(await edit.isVisible(), true)
      assert.equal(document.engine.descriptor.revision, beforeRevision)
      await edit.getByRole('textbox').fill('example.org/report')
      await edit.getByRole('button', { name: 'OK', exact: true }).click()
      await edit.waitFor({ state: 'hidden' })
      await document.bridge!.request('view.flush', {})
      assert.equal(document.engine.descriptor.revision, beforeRevision + 1)
      assert.equal(
        (document.engine.state.sheets[0]!.features!.hyperlinkEdits as Record<string, unknown>).A2,
        'https://example.org/report',
      )
      const after = structuredClone(document.engine.state)
      await vscode.commands.executeCommand('workbench.action.files.save')
      const saved = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      assert.match(
        await saved.file('xl/worksheets/_rels/sheet1.xml.rels')!.async('string'),
        /https:\/\/example.org\/report/,
      )
      const initialZip = await JSZip.loadAsync(original.buffer)
      const initialXml = await initialZip.file('xl/worksheets/sheet1.xml')!.async('string')
      const savedXml = await saved.file('xl/worksheets/sheet1.xml')!.async('string')
      assert.equal(
        savedXml.match(/<sheetData[\s\S]*?<\/sheetData>/)![0],
        initialXml.match(/<sheetData[\s\S]*?<\/sheetData>/)![0],
      )
      assert.equal(
        await saved.file('customXml/manual-sentinel.xml')!.async('string'),
        '<preserve id="manual-fixture">unchanged</preserve>',
      )
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, before)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, after)
      await open()
      assert.equal(await edit.getByRole('textbox').inputValue(), 'https://example.org/report')
      await edit.getByRole('button', { name: 'Remove Link', exact: true }).click()
      await edit.waitFor({ state: 'hidden' })
      await document.bridge!.request('view.flush', {})
      await vscode.commands.executeCommand('workbench.action.files.save')
      const removedZip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      assert.doesNotMatch(
        await removedZip.file('xl/worksheets/sheet1.xml')!.async('string'),
        /<hyperlink\b[^>]*ref="A2"/,
      )
      await open()
      const add = frame.getByRole('dialog', { name: 'Insert Link', exact: true })
      await add.getByRole('textbox').fill('Other!A1')
      await add.getByRole('button', { name: 'OK', exact: true }).click()
      await add.waitFor({ state: 'hidden' })
      await document.bridge!.request('view.flush', {})
      await vscode.commands.executeCommand('workbench.action.files.save')
      const internalZip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      assert.match(
        await internalZip.file('xl/worksheets/sheet1.xml')!.async('string'),
        /<hyperlink\b[^>]*ref="A2"[^>]*location="Other!A1"/,
      )
      await open()
      await edit.getByRole('textbox').fill('https://example.org/stale')
      await document.engine.submit({
        schemaVersion: 1,
        sessionId: document.sessionId,
        transactionId: randomUUID(),
        expectedRevision: document.engine.descriptor.revision,
        origin: 'manual',
        summary: 'Concurrent edit while link dialog is open',
        operations: [
          { op: 'set_note', sheetId: before.sheets[0]!.id, address: 'A1', text: 'concurrent' },
        ],
      })
      const concurrentRevision = document.engine.descriptor.revision
      await edit.getByRole('button', { name: 'OK', exact: true }).click()
      await frame.getByRole('alert').filter({ hasText: 'document changed' }).waitFor()
      assert.equal(await edit.isVisible(), true)
      assert.equal(document.engine.descriptor.revision, concurrentRevision)
      assert.equal(
        (document.engine.state.sheets[0]!.features!.hyperlinkEdits as Record<string, unknown>).A2,
        '#Other!A1',
      )
      await edit.getByRole('button', { name: 'Cancel', exact: true }).click()
      await vscode.commands.executeCommand('workbench.action.files.save')
      const previousSession = document.sessionId
      await vscode.commands.executeCommand('workbench.action.closeActiveEditor')
      const closingStarted = Date.now()
      while (api.getDocument(uri) && Date.now() - closingStarted < 10_000)
        await new Promise((resolve) => setTimeout(resolve, 50))
      assert.equal(
        api.getDocument(uri),
        undefined,
        'old custom document is disposed before reopening',
      )
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      const reopened = api.getDocument(uri)!
      await reopened.whenViewReady()
      assert.notEqual(reopened.sessionId, previousSession)
      assert.equal(reopened.engine.state.sheets[0]!.features?.hyperlinkEdits, undefined)
      let reopenedFrame: Frame | undefined
      const reopenStarted = Date.now()
      while (!reopenedFrame && Date.now() - reopenStarted < 20_000) {
        for (const context of browser.contexts())
          for (const page of context.pages()) {
            if (!(await page.title()).includes('insert-link-ui.xlsx')) continue
            for (const candidate of page.frames())
              if (
                await candidate
                  .locator('.workbook-desktop-header')
                  .count()
                  .catch(() => 0)
              )
                reopenedFrame = candidate
          }
        if (!reopenedFrame) await new Promise((resolve) => setTimeout(resolve, 100))
      }
      assert.ok(reopenedFrame)
      await vscode.commands.executeCommand('genoffice.workbook.goto')
      const reopenedGoTo = reopenedFrame.getByRole('dialog', { name: 'Go To', exact: true })
      await reopenedGoTo.getByRole('textbox', { name: 'Reference', exact: true }).fill('Data!A2')
      await reopenedGoTo.getByRole('button', { name: 'Go', exact: true }).click()
      await reopenedGoTo.waitFor({ state: 'hidden' })
      await reopenedFrame.getByRole('tab', { name: 'Insert', exact: true }).click()
      await reopenedFrame.locator('.workbook-desktop-header [data-tip="Link"]').click()
      const reopenedLink = reopenedFrame.getByRole('dialog', { name: 'Edit Link', exact: true })
      assert.equal(await reopenedLink.getByRole('textbox').inputValue(), '#Other!A1')
      await reopenedLink.getByRole('button', { name: 'Cancel', exact: true }).click()
      const presented = (await reopened.bridge!.request('view.presentation', {
        stateId: reopened.engine.state.stateId,
        range: 'A2',
      })) as { values: unknown[][] }
      assert.deepEqual(presented.values, [[12.5]])
      assert.equal(reopened.engine.descriptor.revision, 0)
      assert.equal(reopened.engine.descriptor.dirty, false)
    } finally {
      if (api.getDocument(uri))
        await vscode.commands.executeCommand('workbench.action.files.save', uri)
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
      await browser.close()
    }
  })
})
