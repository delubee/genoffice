import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import JSZip from 'jszip'
import { suite, test } from 'mocha'
import { chromium, expect, type Frame } from '@playwright/test'
import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'
import { gotoReference } from './goto-helper'

suite('Desktop Color Filter UI', () => {
  test('persists fill colors, multiple choices and imported popup editing through native history and reopen', async function () {
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
      'color-filter-ui.xlsx',
    )
    const values = [
      ['Color', 'Count'],
      ['Red', 1],
      ['Blue', 2],
      ['Plain', 3],
      ['White', 4],
    ]
    const fixture = await applyCellEditsToXlsx(await manualWorkbook(), [
      ...values.flatMap((row, r) =>
        row.map((value, c) => ({
          sheetName: 'Data',
          row: r,
          column: c,
          writeValue: true,
          cell: { value },
          ...(c === 0 && r === 1 ? { style: { fillColor: '#FF0000', fontColor: '#0000FF' } } : {}),
          ...(c === 0 && r === 2 ? { style: { fillColor: '#0000FF', fontColor: '#FF0000' } } : {}),
          ...(c === 0 && r === 4 ? { style: { fillColor: '#FFFFFF', fontColor: '#0000FF' } } : {}),
        })),
      ),
      {
        sheetName: 'Data',
        row: 79,
        column: 0,
        writeValue: true,
        cell: { value: 'Separate far table' },
      },
    ])
    await vscode.workspace.fs.writeFile(uri, fixture.buffer)
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
            if (!(await page.title()).includes('color-filter-ui.xlsx')) continue
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
      await frame.page().setViewportSize({ width: 1200, height: 900 })
      await vscode.commands.executeCommand('notifications.clearAll')
      await vscode.commands.executeCommand('workbench.action.closeAuxiliaryBar')
      await vscode.commands.executeCommand('workbench.action.closeSidebar')
      await vscode.commands.executeCommand('workbench.action.closePanel')
      await gotoReference(frame, 'Data!A1:B5')
      await frame.locator('#ribbon-tab-data').click()
      const band = frame.locator('.workbook-shared-sort-filter')
      await band.getByRole('button', { name: /^Filter\b/ }).click()
      await document.bridge!.request('view.flush', {})
      await document.engine.submit({
        schemaVersion: 1,
        sessionId: document.sessionId,
        transactionId: crypto.randomUUID(),
        expectedRevision: document.engine.descriptor.revision,
        origin: 'agent',
        summary: 'Set known header width',
        operations: [
          {
            op: 'set_col_width',
            sheetId: document.engine.state.sheets[0]!.id,
            column: 'A',
            widthPx: 140,
          },
        ],
      })
      await document.bridge!.request('view.flush', {})
      const popup = (target: Frame) =>
        target
          .locator('[data-u-comp="sheets-filter-panel"]')
          .filter({ has: target.locator('[data-u-comp="sheets-filter-panel-footer"]') })
      const openHeader = async (target: Frame) => {
        await expect
          .poll(() =>
            target
              .locator('.workbook-grid canvas')
              .evaluateAll((elements) =>
                Math.max(...elements.map((canvas) => canvas.getBoundingClientRect().height)),
              ),
          )
          .toBeGreaterThan(150)
        const index = await target.locator('.workbook-grid canvas').evaluateAll(
          (elements) =>
            elements
              .map((canvas, index) => ({
                index,
                area: canvas.getBoundingClientRect().width * canvas.getBoundingClientRect().height,
              }))
              .sort((a, b) => b.area - a.area)[0]!.index,
        )
        const box = (await target.locator('.workbook-grid canvas').nth(index).boundingBox())!
        await target.page().mouse.click(box.x + 46 + 140 - 9, box.y + 20 + 20 - 9)
        await expect(popup(target)).toBeVisible()
      }
      const findFrame = async (): Promise<Frame> => {
        const until = Date.now() + 20000
        while (Date.now() < until) {
          for (const context of browser.contexts())
            for (const page of context.pages()) {
              if (!(await page.title()).includes('color-filter-ui.xlsx')) continue
              for (const target of page.frames())
                if (
                  await target
                    .locator('#ribbon-tab-data')
                    .count()
                    .catch(() => 0)
                )
                  return target
            }
          await new Promise((resolve) => setTimeout(resolve, 100))
        }
        throw new Error('Workbook frame did not reopen')
      }
      const reopen = async () => {
        const prior = api.getDocument(uri)!.sessionId
        await vscode.commands.executeCommand('workbench.action.files.save')
        await vscode.commands.executeCommand('workbench.action.closeActiveEditor')
        const until = Date.now() + 10000
        while (api.getDocument(uri) && Date.now() < until)
          await new Promise((resolve) => setTimeout(resolve, 50))
        await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
        const document = api.getDocument(uri)!
        await document.whenViewReady()
        assert.notEqual(document.sessionId, prior)
        assert.equal(document.engine.descriptor.revision, 0)
        assert.equal(document.engine.descriptor.dirty, false)
        return { document, frame: await findFrame() }
      }
      const before = structuredClone(document.engine.state),
        revision = document.engine.descriptor.revision
      await openHeader(frame)
      let panel = popup(frame)
      await panel.getByText('By Colors', { exact: true }).click()
      await panel
        .locator('[data-u-comp="sheets-filter-panel-colors-container"] button[style*="255, 0, 0"]')
        .click()
      await panel.getByRole('button', { name: 'Confirm', exact: true }).click()
      await document.bridge!.request('view.flush', {})
      await expect(frame.getByRole('alert')).toHaveCount(0)
      await expect(panel).toBeHidden()
      const after = structuredClone(document.engine.state)
      assert.equal(document.engine.descriptor.revision, revision + 1)
      assert.deepEqual((after.sheets[0]!.features!.filterEdits as any).filter.columns, [
        { colId: 0, colors: { fill: ['#FF0000'] } },
      ])
      assert.deepEqual((after.sheets[0]!.features!.filterEdits as any).hiddenRows, [2, 3, 4])
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, before)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, after)
      await vscode.commands.executeCommand('workbench.action.files.save')
      let zip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      assert.match(
        await zip.file('xl/worksheets/sheet1.xml')!.async('string'),
        /<colorFilter dxfId="[0-9]+" cellColor="1"/,
      )
      assert.match(
        await zip.file('xl/styles.xml')!.async('string'),
        /<dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/,
      )
      const imported = await reopen()
      await openHeader(imported.frame)
      panel = popup(imported.frame)
      const colors = panel.locator('[data-u-comp="sheets-filter-panel-colors-container"]')
      await expect(colors).toBeVisible()
      await expect(colors.locator('.univerjs-icon-success-icon')).toHaveCount(1)
      const importedBefore = structuredClone(imported.document.engine.state)
      await colors.locator('button[style*="0, 0, 255"]').click()
      await panel.getByRole('button', { name: 'Confirm', exact: true }).click()
      await imported.document.bridge!.request('view.flush', {})
      await expect(imported.frame.getByRole('alert')).toHaveCount(0)
      await expect(panel).toBeHidden()
      const multi = structuredClone(imported.document.engine.state)
      assert.deepEqual((multi.sheets[0]!.features!.filterEdits as any).filter.columns, [
        { colId: 0, colors: { fill: ['#FF0000', '#0000FF'] } },
      ])
      assert.deepEqual((multi.sheets[0]!.features!.filterEdits as any).hiddenRows, [3, 4])
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(imported.document.engine.state, importedBefore)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(imported.document.engine.state, multi)
      const twice = await reopen()
      await openHeader(twice.frame)
      panel = popup(twice.frame)
      await expect(
        panel.locator(
          '[data-u-comp="sheets-filter-panel-colors-container"] .univerjs-icon-success-icon',
        ),
      ).toHaveCount(2)
      const snapshot = twice.document.snapshot()
      try {
        const cells = new Map(
          (await snapshot.read(twice.document.engine.state.sheets[0]!.id, 'A1:B5')).cells.map(
            (cell) => [cell.address, cell],
          ),
        )
        assert.equal(cells.get('A2')!.rawValue, 'Red')
        assert.equal(cells.get('A2')!.style!.fillColor, '#FF0000')
        assert.equal(cells.get('A3')!.style!.fontColor, '#FF0000')
        assert.equal(cells.get('B5')!.rawValue, 4)
      } finally {
        snapshot.release()
      }
      await panel.getByRole('button', { name: 'Clear Filter', exact: true }).click()
      await twice.document.bridge!.request('view.flush', {})
      await expect(panel).toBeHidden()
      assert.deepEqual(
        (twice.document.engine.state.sheets[0]!.features!.filterEdits as any).filter.columns,
        [],
      )
      await vscode.commands.executeCommand('workbench.action.files.save')
      zip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
      assert.match(xml, /<autoFilter ref="A1:B5"\s*\/>/)
      assert.ok(!xml.includes('go:colors'))
      assert.equal(
        await zip.file('customXml/manual-sentinel.xml')!.async('string'),
        '<preserve id="manual-fixture">unchanged</preserve>',
      )
    } finally {
      if (api.getDocument(uri))
        await vscode.commands.executeCommand('workbench.action.files.save', uri)
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
      await browser.close()
    }
  })
})
