import { pickOutlineCommand } from '../helpers/ribbon-menu'
import { workbookFileSchema } from '@genoffice/xlsx-gateway/shared/workbook-protocol'
import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import JSZip from 'jszip'
import { suite, test } from 'mocha'
import { chromium, expect, type Frame } from '@playwright/test'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'

suite('Desktop Outline UI', () => {
  test('preserves inner folds, uses desktop gutters and clears outline atomically', async function () {
    const port = process.env.GENOFFICE_TEST_CDP_PORT
    if (!port) return this.skip()
    this.timeout(90_000)
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as {
      getDocument(uri: vscode.Uri): WorkbookDocument | undefined
    }
    const uri = vscode.Uri.joinPath(
      vscode.workspace.workspaceFolders![0]!.uri,
      'outline-detail-ui.xlsx',
    )
    await vscode.workspace.fs.writeFile(uri, await manualWorkbook())
    const browser = await chromium.connectOverCDP(`http://127.0.0.1:${port}`)
    try {
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      const document = api.getDocument(uri)!
      await document.whenViewReady()
      let frame: Frame | undefined
      const deadline = Date.now() + 20_000
      while (!frame && Date.now() < deadline) {
        for (const context of browser.contexts())
          for (const page of context.pages()) {
            if (!(await page.title()).includes('outline-detail-ui.xlsx')) continue
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
      const goTo = async (reference: string) => {
        await vscode.commands.executeCommand('genoffice.workbook.goto')
        const go = frame!.getByRole('dialog', { name: 'Go To', exact: true })
        await go.getByRole('textbox', { name: 'Reference', exact: true }).fill(reference)
        await go.getByRole('button', { name: 'Go', exact: true }).click()
        await go.waitFor({ state: 'hidden' })
        await frame!.locator('#ribbon-tab-data').click()
      }
      const click = async (name: string) => {
        await pickOutlineCommand(frame!, name)
        await document.bridge!.request('view.flush', {})
      }
      await goTo('Data!A2:A5')
      await click('Group Rows')
      await goTo('Data!A3:A4')
      await click('Group Rows')
      await goTo('Data!B1:D1')
      await click('Group Columns')
      await goTo('Data!C1')
      await click('Group Columns')
      await click('Hide Detail (Columns)')
      await goTo('Data!B1')
      await click('Hide Detail (Columns)')
      await click('Show Detail (Columns)')
      await goTo('Data!A3')
      await click('Hide Detail (Rows)')
      await goTo('Data!A2')
      await click('Hide Detail (Rows)')
      await click('Show Detail (Rows)')
      const beforeClear = structuredClone(document.engine.state)
      const hidden = (axis: 'rows' | 'cols', index: number) => {
        const edits = document.engine.state.sheets[0]!.features!.visibilityEdits as {
          axis: string
          start: number
          count: number
          hidden: boolean
        }[]
        return [...edits]
          .reverse()
          .find(
            (edit) => edit.axis === axis && index >= edit.start && index < edit.start + edit.count,
          )?.hidden
      }
      assert.deepEqual(
        [1, 2, 3, 4].map((index) => hidden('rows', index)),
        [false, true, true, false],
      )
      assert.deepEqual(
        [1, 2, 3].map((index) => hidden('cols', index)),
        [false, true, false],
      )
      await vscode.commands.executeCommand('workbench.action.files.save')
      const folded = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const foldedXml = await folded.file('xl/worksheets/sheet1.xml')!.async('string')
      const attributes = (row: number) =>
        foldedXml.match(new RegExp('<row\\b[^>]*r="' + row + '"[^>]*>'))![0]
      assert.match(attributes(3), /hidden="1"/)
      assert.match(attributes(4), /hidden="1"/)
      assert.doesNotMatch(attributes(2), /hidden="1"/)
      assert.doesNotMatch(attributes(5), /hidden="1"/)
      assert.match(attributes(5), /collapsed="1"/)
      assert.match(foldedXml, /<col\b[^>]*min="3"[^>]*max="3"[^>]*hidden="1"/)
      const depth = async (axis: 'rows' | 'cols', level: number) => {
        await expect(frame!.locator('.workbook-grid')).toHaveAttribute('aria-busy', 'false')
        const gutter = frame!.locator(`.outline-gutter-${axis}`)
        await expect(gutter).toBeVisible()
        await expect(frame!.locator('.outline-gutter-cols')).toBeVisible()
        await gutter.scrollIntoViewIfNeeded()
        const point = await gutter.evaluate(
          (canvas, { axis, level }) => {
            const bounds = canvas.getBoundingClientRect()
            const grid = [...canvas.ownerDocument.querySelectorAll('.workbook-grid canvas')]
              .map((node) => node.getBoundingClientRect())
              .sort((a, b) => b.width * b.height - a.width * a.height)[0]!
            const along = 3 + (level - 1) * 13 + 6.5
            const columnBand =
              canvas.ownerDocument.querySelector('.outline-gutter-cols')?.getBoundingClientRect()
                .height ?? 0
            const rowDepthTop = Math.max(0, (20 + columnBand - 19) / 2)
            return axis === 'rows'
              ? { x: along, y: grid.top - bounds.top + rowDepthTop + 9.5 }
              : { x: 9.5, y: along }
          },
          { axis, level },
        )
        await gutter.click({ position: point })
        await document.bridge!.request('view.flush', {})
      }
      await depth('rows', 1)
      assert.deepEqual(
        [1, 2, 3, 4].map((index) => hidden('rows', index)),
        [true, true, true, true],
      )
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, beforeClear)
      await depth('cols', 1)
      assert.deepEqual(
        [1, 2, 3].map((index) => hidden('cols', index)),
        [true, true, true],
      )
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, beforeClear)
      const revision = document.engine.descriptor.revision
      await click('Clear Outline')
      assert.equal(document.engine.descriptor.revision, revision + 1)
      assert.deepEqual(
        [1, 2, 3, 4].map((index) => hidden('rows', index)),
        [false, false, false, false],
      )
      assert.equal(hidden('cols', 2), false)
      const cleared = structuredClone(document.engine.state)
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, beforeClear)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, cleared)
      await vscode.commands.executeCommand('workbench.action.files.save')
      const zip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
      assert.doesNotMatch(xml, /outlineLevel="[1-7]"|collapsed="1"|hidden="1"/)
      assert.equal(
        await zip.file('customXml/manual-sentinel.xml')!.async('string'),
        '<preserve id="manual-fixture">unchanged</preserve>',
      )
      const previousSession = document.sessionId
      await vscode.commands.executeCommand('workbench.action.closeActiveEditor')
      const closeDeadline = Date.now() + 10_000
      while (api.getDocument(uri) && Date.now() < closeDeadline)
        await new Promise((resolve) => setTimeout(resolve, 50))
      assert.equal(api.getDocument(uri), undefined)
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      const reopened = api.getDocument(uri)!
      await reopened.whenViewReady()
      assert.notEqual(reopened.sessionId, previousSession)
      assert.equal(reopened.engine.descriptor.revision, 0)
      assert.equal(reopened.engine.descriptor.dirty, false)
      const metadata = workbookFileSchema.parse(reopened.metadata)
      assert.equal(metadata.sheets[0]!.outlineLevelRow ?? 0, 0)
      assert.equal(metadata.sheets[0]!.outlineLevelCol ?? 0, 0)
      const snapshot = reopened.snapshot()
      try {
        const cells = await snapshot.read(reopened.engine.state.sheets[0]!.id, 'A2')
        assert.equal(cells.cells[0]!.rawValue, 12.5)
      } finally {
        snapshot.release()
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
