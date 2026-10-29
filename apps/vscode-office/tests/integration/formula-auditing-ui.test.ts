import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import { suite, test } from 'mocha'
import { chromium, type Frame } from '@playwright/test'
import JSZip from 'jszip'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'

suite('Desktop Formula Auditing UI', () => {
  for (const protectedSheet of [false, true]) {
    test(`traces source cells, checks errors and watches values with native formula-view history and independent reopen (protected=${protectedSheet})`, async function () {
      const port = process.env.GENOFFICE_TEST_CDP_PORT
      if (!port) return this.skip()
      this.timeout(90_000)
      const api = (await vscode.extensions
        .getExtension('genoffice.genoffice-workbook')!
        .activate()) as { getDocument(uri: vscode.Uri): WorkbookDocument | undefined }
      const uri = vscode.Uri.joinPath(
        vscode.workspace.workspaceFolders![0]!.uri,
        'formula-auditing-ui.xlsx',
      )
      const fixture = await JSZip.loadAsync(await manualWorkbook())
      fixture.file(
        'xl/worksheets/sheet1.xml',
        (await fixture.file('xl/worksheets/sheet1.xml')!.async('string')).replace(
          '</sheetData>',
          '<row r="100001"><c r="C100001"><f>A2+1</f><v>13.5</v></c><c r="D100001" t="e"><f>1/0</f><v>#DIV/0!</v></c></row></sheetData>',
        ),
      )
      const protection = '<sheetProtection sheet="1" objects="1" scenarios="1" password="ABCD"/>'
      if (protectedSheet)
        fixture.file(
          'xl/worksheets/sheet1.xml',
          (await fixture.file('xl/worksheets/sheet1.xml')!.async('string')).replace(
            '</sheetData>',
            '</sheetData>' + protection,
          ),
        )
      await vscode.workspace.fs.writeFile(uri, await fixture.generateAsync({ type: 'uint8array' }))
      const browser = await chromium.connectOverCDP(`http://127.0.0.1:${port}`)
      const findFrame = async (): Promise<Frame> => {
        const started = Date.now()
        while (Date.now() - started < 20_000) {
          for (const context of browser.contexts())
            for (const page of context.pages()) {
              if (!(await page.title()).includes('formula-auditing-ui.xlsx')) continue
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
        throw new Error('Function Library frame did not load')
      }
      try {
        await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
        let document = api.getDocument(uri)!
        await document.whenViewReady()
        let frame = await findFrame()
        await document.bridge!.request('view.flush', {})
        const goTo = async (reference: string) => {
          await vscode.commands.executeCommand('genoffice.workbook.goto')
          const dialog = frame.getByRole('dialog', { name: 'Go To', exact: true })
          await dialog.getByRole('textbox', { name: 'Reference', exact: true }).fill(reference)
          await dialog.getByRole('button', { name: 'Go', exact: true }).click()
          await dialog.waitFor({ state: 'hidden' })
        }
        await goTo('C100001')
        await frame.getByRole('tab', { name: 'Formulas', exact: true }).click()
        const ribbon = frame.locator('.workbook-desktop-header')
        const before = structuredClone(document.engine.state)
        await ribbon.getByRole('button', { name: 'Trace Precedents', exact: true }).click()
        await frame.locator('.trace-arrow-layer').waitFor()
        await ribbon.getByRole('button', { name: 'Remove Arrows', exact: true }).click()
        await frame.locator('.trace-arrow-layer').waitFor({ state: 'detached' })
        await goTo('A2')
        await ribbon.getByRole('button', { name: 'Trace Dependents', exact: true }).click()
        await frame.getByRole('status').filter({ hasText: 'Traced 1 dependent' }).waitFor()
        await ribbon.getByRole('button', { name: 'Remove Arrows', exact: true }).click()
        await ribbon.getByRole('button', { name: 'Watch Window', exact: true }).click()
        const watch = frame.getByRole('dialog', { name: 'Watch Window', exact: true })
        await watch.getByRole('button', { name: 'Add watch', exact: true }).click()
        await watch.getByRole('cell', { name: '12.50', exact: true }).waitFor()
        await watch.locator('.watch-close').click()
        await ribbon.getByRole('button', { name: 'Error Checking', exact: true }).click()
        await frame.getByRole('status').filter({ hasText: 'D100001' }).waitFor()
        let selected = (await document.bridge!.request('view.selection', {})) as {
          selection?: { address?: string }
        }
        const started = Date.now()
        while (selected.selection?.address !== 'D100001:D100001' && Date.now() - started < 10000) {
          await new Promise((resolve) => setTimeout(resolve, 50))
          selected = (await document.bridge!.request('view.selection', {})) as typeof selected
        }
        assert.equal(selected.selection?.address, 'D100001:D100001')
        await document.bridge!.request('view.flush', {})
        assert.deepEqual(document.engine.state, before)
        await ribbon.getByRole('button', { name: 'Show Formulas', exact: true }).click()
        try {
          await frame
            .getByRole('status')
            .filter({ hasText: 'Showing formulas' })
            .waitFor({ timeout: 10000 })
        } catch {
          throw new Error(
            'Formula view failed: ' +
              (await frame.getByRole('alert').allTextContents()).join('; ') +
              ' status=' +
              (await frame.getByRole('status').allTextContents()).join('; ') +
              ' state=' +
              JSON.stringify({
                readonly: (
                  document.engine.state.features?.originalMetadata as { readOnly?: boolean }
                )?.readOnly,
                sheet: document.engine.state.sheets.map((sheet) => ({
                  readOnly: sheet.readOnly,
                  protected: sheet.protected,
                  view: sheet.features?.pageSetupEdits,
                })),
              }),
          )
        }
        await document.bridge!.request('view.flush', {})
        const after = structuredClone(document.engine.state)
        assert.equal(
          after.sheets[0]!.features?.pageSetupEdits &&
            (after.sheets[0]!.features.pageSetupEdits as { showFormulas?: boolean }).showFormulas,
          true,
        )
        assert.deepEqual(
          after.sheets.map((sheet) => sheet.cells),
          before.sheets.map((sheet) => sheet.cells),
        )
        await vscode.commands.executeCommand('undo')
        assert.deepEqual(document.engine.state, before)
        await vscode.commands.executeCommand('redo')
        assert.deepEqual(document.engine.state, after)
        await vscode.commands.executeCommand('workbench.action.files.save')
        const saved = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
        const xml = await saved.file('xl/worksheets/sheet1.xml')!.async('string')
        assert.match(xml, /showFormulas="1"/)
        if (protectedSheet) assert.ok(xml.includes(protection))
        assert.ok(xml.includes('<f>A2+1</f>'))
        assert.ok(xml.includes('<f>1/0</f>'))
        assert.equal(
          await saved.file('customXml/manual-sentinel.xml')!.async('string'),
          '<preserve id="manual-fixture">unchanged</preserve>',
        )
        const session = document.sessionId
        await vscode.commands.executeCommand('workbench.action.closeActiveEditor')
        const closed = Date.now()
        while (api.getDocument(uri) && Date.now() - closed < 10000)
          await new Promise((resolve) => setTimeout(resolve, 50))
        assert.equal(api.getDocument(uri), undefined)
        await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
        document = api.getDocument(uri)!
        await document.whenViewReady()
        assert.notEqual(document.sessionId, session)
        frame = await findFrame()
        const presentation = (await document.bridge!.request('view.presentation', {
          stateId: document.engine.state.stateId,
          range: 'A1:B2',
        })) as { view: { showFormulas: boolean } }
        assert.equal(presentation.view.showFormulas, true)
      } finally {
        await vscode.commands.executeCommand('workbench.action.closeAllEditors')
        await vscode.workspace.fs.delete(uri)
        await browser.close()
      }
    })
  }
})
