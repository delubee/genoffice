import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import { suite, test } from 'mocha'
import { chromium, type Frame } from '@playwright/test'
import JSZip from 'jszip'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'

suite('Desktop Picture Transfers UI', () => {
  test('pastes and drops pictures through native history and saves their bytes and exact anchors', async function () {
    const port = process.env.GENOFFICE_TEST_CDP_PORT
    if (!port) return this.skip()
    this.timeout(90_000)
    const pixel =
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII='
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as { getDocument(uri: vscode.Uri): WorkbookDocument | undefined }
    const uri = vscode.Uri.joinPath(
      vscode.workspace.workspaceFolders![0]!.uri,
      'picture-transfers-ui.xlsx',
    )
    await vscode.workspace.fs.writeFile(uri, await manualWorkbook())
    const browser = await chromium.connectOverCDP(`http://127.0.0.1:${port}`)
    const getFrame = async (): Promise<Frame> => {
      const started = Date.now()
      while (Date.now() - started < 20_000) {
        for (const context of browser.contexts())
          for (const page of context.pages()) {
            if (!(await page.title()).includes('picture-transfers-ui.xlsx')) continue
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
      throw new Error('Picture transfer workbook did not load')
    }
    try {
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      let document = api.getDocument(uri)!
      await document.whenViewReady()
      const frame = await getFrame()
      const before = structuredClone(document.engine.state)
      const prevented = await frame.evaluate((base64) => {
        const grid = window.document.querySelector('.workbook-grid') as HTMLElement
        const data = new DataTransfer()
        data.items.add(
          new File([Uint8Array.from(atob(base64), (c) => c.charCodeAt(0))], 'pasted.png', {
            type: 'image/png',
          }),
        )
        grid.tabIndex = 0
        grid.focus()
        const event = new ClipboardEvent('paste', {
          bubbles: true,
          cancelable: true,
          clipboardData: data,
        })
        grid.dispatchEvent(event)
        return event.defaultPrevented
      }, pixel)
      assert.equal(prevented, true)
      await document.bridge!.request('view.flush', {})
      const afterPaste = structuredClone(document.engine.state)
      const pasted = (afterPaste.sheets[0]!.features!.visuals as any[])[0]
      assert.equal(pasted.kind, 'image')
      assert.equal(pasted.anchor.fromRow, 0)
      assert.equal(pasted.anchor.fromColumn, 0)
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, before)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, afterPaste)
      await frame.evaluate((base64) => {
        const grid = window.document.querySelector('.workbook-grid')!
        const canvas = [...grid.querySelectorAll('canvas')].sort(
          (a, b) => b.width * b.height - a.width * a.height,
        )[0]!
        const r = canvas.getBoundingClientRect(),
          data = new DataTransfer()
        data.items.add(
          new File([Uint8Array.from(atob(base64), (c) => c.charCodeAt(0))], 'dropped.png', {
            type: 'image/png',
          }),
        )
        canvas.dispatchEvent(
          new DragEvent('drop', {
            bubbles: true,
            cancelable: true,
            dataTransfer: data,
            clientX: r.x + 450,
            clientY: r.y + 200,
          }),
        )
      }, pixel)
      await document.bridge!.request('view.flush', {})
      const afterDrop = structuredClone(document.engine.state),
        visuals = afterDrop.sheets[0]!.features!.visuals as any[]
      assert.equal(visuals.length, 2)
      assert.ok(visuals[1].anchor.fromRow > 0)
      assert.ok(visuals[1].anchor.fromColumn > 0)
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, afterPaste)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, afterDrop)
      // Inject known clipboard bytes to validate the native menu transaction
      // independently of the platform's clipboard-read permission policy.
      await frame.evaluate((base64) => {
        Object.defineProperty(navigator, 'clipboard', {
          configurable: true,
          value: {
            read: async () => [
              new ClipboardItem({
                'image/png': new Blob([Uint8Array.from(atob(base64), (c) => c.charCodeAt(0))], {
                  type: 'image/png',
                }),
              }),
            ],
          },
        })
      }, pixel)
      await frame
        .locator('.workbook-desktop-header')
        .getByRole('button', { name: 'Paste', exact: true })
        .click()
      await document.bridge!.request('view.flush', {})
      const afterMenu = structuredClone(document.engine.state),
        allVisuals = afterMenu.sheets[0]!.features!.visuals as any[]
      assert.equal(allVisuals.length, 3)
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, afterDrop)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, afterMenu)
      await frame.evaluate((base64) => {
        const grid = window.document.querySelector('.workbook-grid') as HTMLElement,
          data = new DataTransfer()
        data.items.add(
          new File([Uint8Array.from(atob(base64), (c) => c.charCodeAt(0))], 'table-preview.png', {
            type: 'image/png',
          }),
        )
        data.setData(
          'text/html',
          '<table><tr><td style="font-weight:bold;background-color:#ffeecc">Table text</td></tr></table>',
        )
        data.setData('text/plain', 'Table text')
        grid.tabIndex = 0
        grid.focus()
        grid.dispatchEvent(
          new ClipboardEvent('paste', { bubbles: true, cancelable: true, clipboardData: data }),
        )
      }, pixel)
      await document.bridge!.request('view.flush', {})
      const afterTable = structuredClone(document.engine.state)
      assert.equal(afterTable.sheets[0]!.cells.A1?.rawValue, 'Table text')
      assert.equal(afterTable.sheets[0]!.cells.A1?.style?.bold, true)
      assert.equal(afterTable.sheets[0]!.cells.A1?.style?.fillColor, '#FFEECC')
      assert.equal((afterTable.sheets[0]!.features!.visuals as any[]).length, 3)
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, afterMenu)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, afterTable)
      await vscode.commands.executeCommand('workbench.action.files.save')
      assert.equal(document.engine.descriptor.dirty, false)
      const zip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      assert.match(await zip.file('xl/worksheets/sheet1.xml')!.async('string'), /Table text/)
      const media = Object.keys(zip.files).filter(
        (name) => name.startsWith('xl/media/') && !zip.files[name]!.dir,
      )
      assert.ok(media.length > 0)
      for (const name of media) assert.equal(await zip.file(name)!.async('base64'), pixel)
      assert.equal(
        await zip.file('customXml/manual-sentinel.xml')!.async('string'),
        '<preserve id="manual-fixture">unchanged</preserve>',
      )
      const sessionId = document.sessionId
      await vscode.commands.executeCommand('workbench.action.closeActiveEditor')
      const closed = Date.now()
      while (api.getDocument(uri) && Date.now() - closed < 10_000)
        await new Promise((resolve) => setTimeout(resolve, 50))
      assert.equal(api.getDocument(uri), undefined)
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      document = api.getDocument(uri)!
      await document.whenViewReady()
      assert.notEqual(document.sessionId, sessionId)
      const saved = (
        (document.engine.state.features!.originalMetadata as any).visuals as any[]
      ).filter((v) => v.kind === 'image')
      assert.equal(saved.length, 3)
      for (let i = 0; i < 3; i++)
        for (const key of Object.keys(allVisuals[i].anchor))
          assert.equal(saved[i].anchor[key], allVisuals[i].anchor[key])
      const reopened = await getFrame()
      for (const picture of saved) {
        const image = reopened.locator(`[data-print-visual="${picture.id}"] img`)
        await image.waitFor()
        const start = Date.now()
        while (
          !(await image.evaluate((el: HTMLImageElement) => el.complete && el.naturalWidth === 1)) &&
          Date.now() - start < 10_000
        )
          await new Promise((resolve) => setTimeout(resolve, 50))
        assert.equal(await image.evaluate((el: HTMLImageElement) => el.naturalWidth), 1)
      }
    } finally {
      if (api.getDocument(uri))
        await vscode.commands.executeCommand('workbench.action.files.save', uri)
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
      await browser.close()
    }
  })
})
