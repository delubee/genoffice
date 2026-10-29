import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import { suite, test } from 'mocha'
import { chromium, type Frame } from '@playwright/test'
import JSZip from 'jszip'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'

suite('Desktop Name Manager UI', () => {
  test('renames and removes source references through shared UI, native history and independent reopen', async function () {
    const port = process.env.GENOFFICE_TEST_CDP_PORT
    if (!port) return this.skip()
    this.timeout(90_000)
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as { getDocument(uri: vscode.Uri): WorkbookDocument | undefined }
    const uri = vscode.Uri.joinPath(
      vscode.workspace.workspaceFolders![0]!.uri,
      'name-manager-ui.xlsx',
    )
    const fixture = await JSZip.loadAsync(await manualWorkbook())
    const hidden =
      '<definedName name="HiddenPrivate" hidden="1">"private-formula-secret"</definedName>'
    const print =
      '<definedName name="_xlnm.Print_Area" localSheetId="0">Data!$A$1:$D$3</definedName>'
    fixture.file(
      'xl/workbook.xml',
      (await fixture.file('xl/workbook.xml')!.async('string')).replace(
        '</sheets>',
        '</sheets><definedNames><definedName name="Revenue">Data!$A$2</definedName>' +
          hidden +
          print +
          '</definedNames>',
      ),
    )
    fixture.file(
      'xl/worksheets/sheet1.xml',
      (await fixture.file('xl/worksheets/sheet1.xml')!.async('string')).replace(
        '</sheetData>',
        '<row r="100001"><c r="C100001"><f>Revenue+1</f><v>999</v></c></row></sheetData>',
      ),
    )
    await vscode.workspace.fs.writeFile(uri, await fixture.generateAsync({ type: 'uint8array' }))
    const browser = await chromium.connectOverCDP(`http://127.0.0.1:${port}`)
    const findFrame = async (): Promise<Frame> => {
      const started = Date.now()
      while (Date.now() - started < 20_000) {
        for (const context of browser.contexts())
          for (const page of context.pages()) {
            if (!(await page.title()).includes('name-manager-ui.xlsx')) continue
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
      const openManager = async () => {
        await frame.getByRole('tab', { name: 'Formulas', exact: true }).click()
        await frame
          .locator('.workbook-desktop-header')
          .getByRole('button', { name: 'Name Manager', exact: true })
          .click()
        return frame.getByRole('dialog', { name: 'Name Manager', exact: true })
      }
      const before = structuredClone(document.engine.state)
      assert.ok(!JSON.stringify(before).includes('private-formula-secret'))
      let dialog = await openManager()
      assert.equal(
        await dialog.getByRole('cell', { name: 'HiddenPrivate', exact: true }).count(),
        0,
      )
      await dialog.getByRole('textbox', { name: 'Name', exact: true }).fill('HiddenPrivate')
      await dialog.getByRole('textbox', { name: 'Refers to', exact: true }).fill('Data!$A$2')
      await dialog.getByRole('button', { name: 'Add', exact: true }).click()
      await dialog.getByRole('alert').waitFor()
      assert.deepEqual(document.engine.state, before)
      await dialog.getByRole('cell', { name: 'Revenue', exact: true }).click()
      await dialog.getByRole('textbox', { name: 'Name', exact: true }).fill('Sales')
      await dialog.getByRole('button', { name: 'Update', exact: true }).click()
      await dialog.getByRole('cell', { name: 'Sales', exact: true }).waitFor()
      await dialog.getByRole('button', { name: 'Close', exact: true }).click()
      await document.bridge!.request('view.flush', {})
      const after = structuredClone(document.engine.state)
      assert.deepEqual(
        after.sheets.map((s) => s.cells),
        before.sheets.map((s) => s.cells),
      )
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, before)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, after)
      dialog = await openManager()
      await dialog
        .getByRole('cell', { name: 'Sales', exact: true })
        .locator('..')
        .getByRole('button', { name: 'Delete name', exact: true })
        .click()
      await dialog.getByRole('cell', { name: 'Sales', exact: true }).waitFor({ state: 'detached' })
      await dialog.getByRole('textbox', { name: 'Name', exact: true }).fill('Point')
      await dialog.getByRole('textbox', { name: 'Refers to', exact: true }).fill('Data!$A$2')
      await dialog.getByRole('button', { name: 'Add', exact: true }).click()
      await dialog.getByRole('cell', { name: 'Point', exact: true }).waitFor()
      await dialog.getByRole('button', { name: 'Close', exact: true }).click()
      await vscode.commands.executeCommand('workbench.action.files.save')
      const saved = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const xml = await saved.file('xl/workbook.xml')!.async('string')
      assert.match(xml, /<definedName name="Point">Data!\$A\$2<\/definedName>/)
      assert.ok(xml.includes(hidden))
      assert.ok(xml.includes(print))
      const cells = await saved.file('xl/worksheets/sheet1.xml')!.async('string')
      assert.ok(cells.includes('<f>#REF!+1</f>'))
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
      await vscode.commands.executeCommand('genoffice.workbook.goto')
      const go = frame.getByRole('dialog', { name: 'Go To', exact: true })
      await go.getByRole('button', { name: /Point/ }).waitFor()
      await go.getByRole('button', { name: 'Cancel', exact: true }).click()
    } finally {
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
      await browser.close()
    }
  })
})
