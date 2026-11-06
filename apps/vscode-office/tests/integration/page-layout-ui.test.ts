import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import { suite, test } from 'mocha'
import { chromium, type Frame } from '@playwright/test'
import JSZip from 'jszip'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'
import { gotoReference } from './goto-helper'

suite('Desktop Page Layout UI', () => {
  test('echoes source settings and saves shared page controls with native history and unchanged cell XML', async function () {
    const port = process.env.GENOFFICE_TEST_CDP_PORT
    if (!port) return this.skip()
    this.timeout(90_000)
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as { getDocument(uri: vscode.Uri): WorkbookDocument | undefined }
    const uri = vscode.Uri.joinPath(
      vscode.workspace.workspaceFolders![0]!.uri,
      'page-layout-ui.xlsx',
    )
    const original = await JSZip.loadAsync(await manualWorkbook())
    let sheet = await original.file('xl/worksheets/sheet1.xml')!.async('string')
    const data = /<sheetData\b[^>]*>[\s\S]*?<\/sheetData>/.exec(sheet)![0]
    sheet = sheet.replace(
      /<worksheet\b[^>]*>/,
      (match) => match + '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>',
    )
    sheet = sheet.replace(
      '</worksheet>',
      '<printOptions gridLines="1" headings="0"/><pageSetup paperSize="5" orientation="landscape" fitToWidth="2" fitToHeight="1"/><rowBreaks count="2" manualBreakCount="2"><brk id="2" min="0" max="16383" man="1"/><brk id="10" min="0" max="16383" man="1"/></rowBreaks><colBreaks count="1" manualBreakCount="1"><brk id="1" min="0" max="1048575" man="1"/></colBreaks></worksheet>',
    )
    original.file('xl/worksheets/sheet1.xml', sheet)
    let workbook = await original.file('xl/workbook.xml')!.async('string')
    const names =
      '<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">Data!$B$2:$E$10</definedName><definedName name="_xlnm.Print_Titles" localSheetId="0">Data!$1:$3</definedName></definedNames>'
    workbook = workbook.includes('<calcPr')
      ? workbook.replace('<calcPr', names + '<calcPr')
      : workbook.replace('</workbook>', names + '</workbook>')
    original.file('xl/workbook.xml', workbook)
    await vscode.workspace.fs.writeFile(uri, await original.generateAsync({ type: 'uint8array' }))
    const browser = await chromium.connectOverCDP(`http://127.0.0.1:${port}`)
    const findFrame = async (): Promise<Frame> => {
      const started = Date.now()
      while (Date.now() - started < 20_000) {
        for (const context of browser.contexts())
          for (const page of context.pages()) {
            if (!(await page.title()).includes('page-layout-ui.xlsx')) continue
            for (const frame of page.frames())
              if (
                await frame
                  .locator('.workbook-desktop-header')
                  .count()
                  .catch(() => 0)
              )
                return frame
          }
        await new Promise((resolve) => setTimeout(resolve, 100))
      }
      throw new Error('Page Layout frame did not load')
    }
    try {
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      let document = api.getDocument(uri)!
      await document.whenViewReady()
      let frame = await findFrame()
      await gotoReference(frame, 'B4:C50')
      await frame.getByRole('tab', { name: 'Page Layout', exact: true }).click()
      let header = frame.locator('.workbook-desktop-header')
      assert.equal(
        await header.getByRole('button', { name: 'Themes', exact: true }).isDisabled(),
        true,
      )
      assert.match(
        (await header
          .getByRole('button', { name: 'Orientation', exact: true })
          .locator('..')
          .locator('..')
          .getAttribute('data-tip')) ?? '',
        /Landscape/,
      )
      const beforeCells = structuredClone(document.engine.state.sheets[0]!.cells)
      const choose = async (name: string, option: string) => {
        await header.getByRole('button', { name, exact: true }).click()
        await frame.getByRole('option', { name: option, exact: true }).click()
        await document.bridge!.request('view.flush', {})
      }
      await choose('Orientation', 'Portrait')
      await choose('Margins', 'Narrow')
      await choose('Size', 'A3')
      await choose('Print Area', 'Set Print Area')
      await choose('Print Titles', 'Repeat Selected Rows')
      await choose('Fit to width', '3 pages')
      assert.equal(
        await header
          .getByRole('button', { name: 'Fit to width', exact: true })
          .getAttribute('aria-expanded'),
        'false',
      )
      const grid = header.locator('.check-column').filter({ hasText: 'Gridlines' })
      await grid.getByRole('button', { name: /Print/ }).click()
      await document.bridge!.request('view.flush', {})
      await grid.getByRole('button', { name: /View/ }).click()
      await document.bridge!.request('view.flush', {})
      const headings = header.locator('.check-column').filter({ hasText: 'Headings' })
      await headings.getByRole('button', { name: /Print/ }).click()
      await document.bridge!.request('view.flush', {})
      const beforeBreak = structuredClone(document.engine.state)
      await choose('Breaks', 'Insert Page Break')
      const after = structuredClone(document.engine.state)
      const page = after.sheets[0]!.features!.pageSetupEdits as any
      assert.equal(page.orientation, 'portrait')
      assert.equal(page.paperSize, 8)
      assert.equal(page.margins, 'narrow')
      assert.equal(page.printArea, 'B4:C50')
      assert.equal(page.printTitles, '4:24')
      assert.equal(page.fitToWidth, 3)
      assert.equal(page.fitToHeight, 0)
      assert.deepEqual(page.rowBreaks, [2, 3, 10])
      assert.deepEqual(page.colBreaks, [1])
      assert.equal(page.printGridlines, false)
      assert.equal(page.showGridlines, false)
      assert.equal(page.printHeadings, true)
      assert.deepEqual(after.sheets[0]!.cells, beforeCells)
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, beforeBreak)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, after)
      await vscode.commands.executeCommand('workbench.action.files.save')
      const saved = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const xml = await saved.file('xl/worksheets/sheet1.xml')!.async('string')
      assert.equal(/<sheetData\b[^>]*>[\s\S]*?<\/sheetData>/.exec(xml)![0], data)
      assert.match(xml, /orientation="portrait"/)
      assert.match(xml, /paperSize="8"/)
      assert.match(xml, /fitToWidth="3"/)
      assert.match(xml, /showGridLines="0"/)
      assert.match(xml, /headings="1"/)
      assert.doesNotMatch(/<printOptions\b[^>]*>/.exec(xml)?.[0] ?? '', /gridLines="1"/)
      const savedWorkbook = await saved.file('xl/workbook.xml')!.async('string')
      assert.match(savedWorkbook, /\$B\$4:\$C\$50/)
      assert.match(savedWorkbook, /\$4:\$24/)
      assert.equal(
        await saved.file('customXml/manual-sentinel.xml')!.async('string'),
        '<preserve id="manual-fixture">unchanged</preserve>',
      )
      const id = document.sessionId
      await vscode.commands.executeCommand('workbench.action.closeActiveEditor')
      const closed = Date.now()
      while (api.getDocument(uri) && Date.now() - closed < 10_000)
        await new Promise((resolve) => setTimeout(resolve, 50))
      assert.equal(api.getDocument(uri), undefined)
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      document = api.getDocument(uri)!
      await document.whenViewReady()
      assert.notEqual(document.sessionId, id)
      frame = await findFrame()
      await frame.getByRole('tab', { name: 'Page Layout', exact: true }).click()
      header = frame.locator('.workbook-desktop-header')
      assert.match(
        (await header
          .getByRole('button', { name: 'Orientation', exact: true })
          .locator('..')
          .locator('..')
          .getAttribute('data-tip')) ?? '',
        /Portrait/,
      )
      assert.match(
        (await header.getByRole('button', { name: 'Fit to width', exact: true }).textContent()) ??
          '',
        /3 pages/,
      )
    } finally {
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
      await browser.close()
    }
  })
})
