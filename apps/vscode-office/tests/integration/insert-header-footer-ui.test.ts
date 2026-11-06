import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import JSZip from 'jszip'
import { suite, test } from 'mocha'
import { chromium, type Frame } from '@playwright/test'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'

suite('Desktop Insert Header Footer UI', () => {
  test('edits source print sections and preserves variants through native history, save and reopen', async function () {
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
      'insert-header-footer-ui.xlsx',
    )
    const zip = await JSZip.loadAsync(await manualWorkbook())
    const sheetPath = 'xl/worksheets/sheet1.xml'
    let xml = await zip.file(sheetPath)!.async('string')
    const variants =
      '<evenHeader>&amp;LEven &amp;P</evenHeader><firstFooter>&amp;CFirst only</firstFooter>'
    const headerFooter =
      '<headerFooter differentOddEven="1" differentFirst="1"><oddHeader>&amp;LOriginal left&amp;CPage &amp;P of &amp;N</oddHeader><oddFooter>&amp;RConfidential</oddFooter>' +
      variants +
      '</headerFooter>'
    xml = xml
      .replace(/<headerFooter\b[^>]*>[\s\S]*?<\/headerFooter>/, '')
      .replace('</worksheet>', headerFooter + '</worksheet>')
    zip.file(sheetPath, xml)
    await vscode.workspace.fs.writeFile(uri, await zip.generateAsync({ type: 'uint8array' }))
    const browser = await chromium.connectOverCDP(`http://127.0.0.1:${port}`)
    const getFrame = async (): Promise<Frame> => {
      const started = Date.now()
      while (Date.now() - started < 20_000) {
        for (const context of browser.contexts())
          for (const page of context.pages()) {
            if (!(await page.title()).includes('insert-header-footer-ui.xlsx')) continue
            for (const candidate of page.frames())
              if (
                await candidate
                  .locator('.workbook-desktop-header')
                  .count()
                  .catch(() => 0)
              )
                return candidate
          }
        await new Promise((resolve) => setTimeout(resolve, 100))
      }
      throw new Error('Workbook frame did not load')
    }
    try {
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      const document = api.getDocument(uri)!
      await document.whenViewReady()
      const frame = await getFrame()
      await frame.getByRole('tab', { name: 'Insert', exact: true }).click()
      const open = () =>
        frame.locator('.workbook-desktop-header [data-tip="Header & Footer"]').click()
      const dialog = frame.getByRole('dialog', { name: 'Header & Footer', exact: true })
      const before = structuredClone(document.engine.state)
      await open()
      assert.equal(
        await dialog.getByRole('textbox', { name: 'Header · Left', exact: true }).inputValue(),
        'Original left',
      )
      await dialog.getByRole('button', { name: 'Cancel', exact: true }).click()
      assert.deepEqual(document.engine.state, before)
      await open()
      await dialog.getByRole('textbox', { name: 'Header · Left', exact: true }).fill('Report $&&')
      await dialog.getByRole('textbox', { name: 'Footer · Right', exact: true }).fill('')
      const originalSubmit = document.engine.submit.bind(document.engine)
      let release: (() => void) | undefined
      let entered: (() => void) | undefined
      const gate = new Promise<void>((resolve) => {
        release = resolve
      })
      const submitting = new Promise<void>((resolve) => {
        entered = resolve
      })
      document.engine.submit = async (...args) => {
        entered!()
        await gate
        return originalSubmit(...args)
      }
      try {
        const ok = dialog.getByRole('button', { name: 'OK', exact: true })
        await ok.click()
        await submitting
        assert.equal(await ok.isDisabled(), true)
        assert.equal(
          await dialog.getByRole('button', { name: 'Cancel', exact: true }).isDisabled(),
          true,
        )
        assert.equal(
          await dialog.getByRole('textbox', { name: 'Header · Left', exact: true }).isDisabled(),
          true,
        )
        await ok.dispatchEvent('click')
        await dialog.press('Escape')
        assert.equal(await dialog.isVisible(), true)
        assert.deepEqual(document.engine.state, before)
      } finally {
        release!()
        document.engine.submit = originalSubmit
      }
      await dialog.waitFor({ state: 'hidden' })
      await document.bridge!.request('view.flush', {})
      const after = structuredClone(document.engine.state)
      assert.deepEqual(after.sheets[0]!.features!.pageSetupEdits, {
        header: { left: 'Report $&&', center: 'Page &P of &N' },
        footer: null,
      })
      assert.deepEqual(after.sheets[0]!.cells, before.sheets[0]!.cells)
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, before)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, after)
      await vscode.commands.executeCommand('workbench.action.files.save')
      const saved = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const savedXml = await saved.file(sheetPath)!.async('string')
      assert.match(savedXml, /Report \$&amp;&amp;/)
      assert.doesNotMatch(savedXml, /<oddFooter/)
      assert.ok(savedXml.includes(variants))
      assert.match(savedXml, /differentOddEven="1"/)
      assert.match(savedXml, /differentFirst="1"/)
      assert.equal(
        savedXml.match(/<sheetData[\s\S]*?<\/sheetData>/)![0],
        xml.match(/<sheetData[\s\S]*?<\/sheetData>/)![0],
      )
      assert.equal(
        await saved.file('customXml/manual-sentinel.xml')!.async('string'),
        '<preserve id="manual-fixture">unchanged</preserve>',
      )
      const previousSession = document.sessionId
      await vscode.commands.executeCommand('workbench.action.closeActiveEditor')
      const started = Date.now()
      while (api.getDocument(uri) && Date.now() - started < 10_000)
        await new Promise((resolve) => setTimeout(resolve, 50))
      assert.equal(api.getDocument(uri), undefined)
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      const reopened = api.getDocument(uri)!
      await reopened.whenViewReady()
      assert.notEqual(reopened.sessionId, previousSession)
      const reopenedFrame = await getFrame()
      await reopenedFrame.getByRole('tab', { name: 'Insert', exact: true }).click()
      await reopenedFrame.locator('.workbook-desktop-header [data-tip="Header & Footer"]').click()
      const reopenedDialog = reopenedFrame.getByRole('dialog', {
        name: 'Header & Footer',
        exact: true,
      })
      assert.equal(
        await reopenedDialog
          .getByRole('textbox', { name: 'Header · Left', exact: true })
          .inputValue(),
        'Report $&&',
      )
      assert.equal(
        await reopenedDialog
          .getByRole('textbox', { name: 'Header · Center', exact: true })
          .inputValue(),
        'Page &P of &N',
      )
      assert.equal(
        await reopenedDialog
          .getByRole('textbox', { name: 'Footer · Right', exact: true })
          .inputValue(),
        '',
      )
      await reopenedDialog.getByRole('button', { name: 'Cancel', exact: true }).click()
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
