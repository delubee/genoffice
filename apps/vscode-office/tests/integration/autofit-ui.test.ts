import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import JSZip from 'jszip'
import { suite, test } from 'mocha'
import { chromium, type Frame } from '@playwright/test'
import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'

suite('Desktop Home AutoFit UI', () => {
  test('measures original offscreen cells from real menu clicks, saves and uses native history', async function () {
    const port = process.env.GENOFFICE_TEST_CDP_PORT
    if (!port) return this.skip()
    this.timeout(90_000)
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as {
      getDocument(uri: vscode.Uri): WorkbookDocument | undefined
      revertForTests(uri: vscode.Uri): Promise<void>
    }
    const uri = vscode.Uri.joinPath(vscode.workspace.workspaceFolders![0]!.uri, 'autofit-ui.xlsx')
    const original = await applyCellEditsToXlsx(await manualWorkbook(), [
      {
        sheetName: 'Data',
        row: 79,
        column: 0,
        writeValue: true,
        cell: {
          value:
            'Original offscreen text must participate in the native AutoFit column measurement',
        },
      },
      {
        sheetName: 'Data',
        row: 0,
        column: 3,
        writeValue: true,
        cell: {
          value:
            'Wrapped source outside the selected column should increase the whole first row height',
        },
        style: { wrapText: true, fontSize: 14, fontFamily: 'Calibri' },
      },
    ])
    const inputZip = await JSZip.loadAsync(original.buffer)
    const workbookXml = await inputZip.file('xl/workbook.xml')!.async('string')
    inputZip.file(
      'xl/workbook.xml',
      workbookXml.replace(
        '</sheets>',
        '</sheets><definedNames><definedName name="GoTarget">Data!$A$80</definedName></definedNames>',
      ),
    )
    await vscode.workspace.fs.writeFile(uri, await inputZip.generateAsync({ type: 'nodebuffer' }))
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
            if (!(await page.title()).includes('autofit-ui.xlsx')) continue
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
      assert.ok(frame, 'native workbook webview frame')
      const selection = (await document.bridge!.request('view.selection', {})) as {
        selection: { address: string }
      }
      assert.match(selection.selection.address, /^A1(?::A1)?$/)
      const home = frame.locator('.workbook-desktop-header')
      const fit = async (name: string) => {
        await home.getByRole('button', { name: 'Format', exact: true }).click()
        await frame!.getByRole('option', { name, exact: true }).click()
      }
      const waitRevision = async (revision: number) => {
        const start = Date.now()
        while (document.engine.descriptor.revision < revision && Date.now() - start < 20_000)
          await new Promise((resolve) => setTimeout(resolve, 50))
        assert.equal(document.engine.descriptor.revision, revision)
        await document.bridge!.request('view.flush', {})
        assert.equal(await frame!.getByRole('alert').count(), 0)
      }
      const initial = structuredClone(document.engine.state.sheets[0]!)
      const revision = document.engine.descriptor.revision
      await fit('AutoFit Row Height')
      await waitRevision(revision + 1)
      const rowHeight = document.engine.state.sheets[0]!.features!.rowHeights as Record<
        string,
        number
      >
      assert.ok(rowHeight['1']! > 30, 'wrapped D1 measured for selection A1')
      const beforeColumn = structuredClone(document.engine.state.sheets[0]!)
      await fit('AutoFit Column Width')
      await waitRevision(revision + 2)
      const widths = document.engine.state.sheets[0]!.features!.colWidths as Record<string, number>
      assert.ok(widths.A! > 200, 'original A80 measured outside initial viewport')
      assert.equal(document.engine.state.sheets[0]!.cells.A80, undefined)
      await vscode.commands.executeCommand('workbench.action.files.save')
      const zip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
      assert.match(sheet, new RegExp('<row\\b[^>]*r="1"[^>]*ht="' + rowHeight['1'] + '"'))
      assert.match(sheet, /<col\b[^>]*min="1"[^>]*width="[1-9]/)
      assert.match(sheet, /Original offscreen text/)
      assert.equal(
        await zip.file('customXml/manual-sentinel.xml')!.async('string'),
        '<preserve id="manual-fixture">unchanged</preserve>',
      )
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state.sheets[0]!, beforeColumn)
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state.sheets[0]!, initial)
      await vscode.commands.executeCommand('redo')
      await vscode.commands.executeCommand('redo')
      assert.ok(
        (document.engine.state.sheets[0]!.features!.colWidths as Record<string, number>).A! > 200,
      )
      const navigationRevision = document.engine.descriptor.revision
      const navigationState = structuredClone(document.engine.state)
      await home.getByRole('button', { name: /Go To/ }).click()
      const dialog = frame.getByRole('dialog', { name: 'Go To', exact: true })
      await dialog.getByRole('textbox', { name: 'Reference', exact: true }).fill('Other!A1')
      await dialog.getByRole('button', { name: 'Go', exact: true }).click()
      await dialog.waitFor({ state: 'hidden' })
      let navigation = (await document.bridge!.request('view.selection', {})) as {
        selection: { sheetId: string; address: string }
      }
      assert.equal(navigation.selection.sheetId, document.engine.state.sheets[1]!.id)
      await home.getByRole('button', { name: /Go To/ }).click()
      await dialog.getByRole('button', { name: /GoTarget/ }).dblclick()
      await dialog.waitFor({ state: 'hidden' })
      navigation = (await document.bridge!.request('view.selection', {})) as {
        selection: { sheetId: string; address: string }
      }
      assert.equal(navigation.selection.sheetId, document.engine.state.sheets[0]!.id)
      assert.match(navigation.selection.address, /^A80(?::A80)?$/)
      assert.equal(document.engine.descriptor.revision, navigationRevision)
      assert.deepEqual(document.engine.state, navigationState)

      await home.getByRole('button', { name: 'Find', exact: true }).click()
      const find = frame.getByRole('dialog', { name: 'Find and Replace', exact: true })
      await find.waitFor({ state: 'visible', timeout: 5_000 })
      await find.getByRole('combobox', { name: 'Look in', exact: true }).selectOption('formula')
      const queryInput = find.getByRole('combobox', { name: 'Find what', exact: true })
      await queryInput.fill('Original offscreen')
      await find.getByRole('status').filter({ hasText: '1 of 1' }).waitFor()
      assert.equal(
        await queryInput.evaluate((element) => window.document.activeElement === element),
        true,
        'live query navigation keeps the input focused',
      )
      await find.getByRole('button', { name: 'Find All', exact: true }).click()
      await find.getByRole('cell', { name: 'A80', exact: true }).waitFor()
      assert.match(await find.getByRole('status').innerText(), /1 of 1/)
      assert.equal(document.engine.descriptor.revision, navigationRevision)
      assert.deepEqual(document.engine.state, navigationState)
      await find.getByRole('tab', { name: 'Replace', exact: true }).click()
      await find
        .getByRole('combobox', { name: 'Replace with', exact: true })
        .fill('Updated offscreen')
      const replaceAll = find.getByRole('button', { name: 'Replace All', exact: true })
      await replaceAll.waitFor({ state: 'visible' })
      const replacingStarted = Date.now()
      while (!(await replaceAll.isEnabled()) && Date.now() - replacingStarted < 20_000)
        await new Promise((resolve) => setTimeout(resolve, 50))
      assert.ok(await replaceAll.isEnabled(), 'complete query enables replacement')
      await replaceAll.click()
      await waitRevision(navigationRevision + 1)
      assert.equal(
        document.engine.state.sheets[0]!.cells.A80!.rawValue,
        'Updated offscreen text must participate in the native AutoFit column measurement',
      )
      await find.locator('.fr-actions').getByRole('button', { name: 'Close', exact: true }).click()
      await find.waitFor({ state: 'hidden' })
      await vscode.commands.executeCommand('workbench.action.files.save')
      const replacedZip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const replacedXml = await replacedZip.file('xl/worksheets/sheet1.xml')!.async('string')
      assert.match(replacedXml, /Updated offscreen text/)
      assert.doesNotMatch(replacedXml, /Original offscreen text/)
      assert.equal(
        await replacedZip.file('customXml/manual-sentinel.xml')!.async('string'),
        '<preserve id="manual-fixture">unchanged</preserve>',
      )
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, navigationState)
      await vscode.commands.executeCommand('redo')
      assert.equal(
        document.engine.state.sheets[0]!.cells.A80!.rawValue,
        'Updated offscreen text must participate in the native AutoFit column measurement',
      )
    } finally {
      if (api.getDocument(uri)) await api.revertForTests(uri)
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
      await browser.close()
    }
  })
})
