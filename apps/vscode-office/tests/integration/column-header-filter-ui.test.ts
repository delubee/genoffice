import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import JSZip from 'jszip'
import { suite, test } from 'mocha'
import { chromium, expect, type Frame } from '@playwright/test'
import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'
import { gotoReference } from './goto-helper'

suite('Desktop Column Header Filter UI', () => {
  test('applies native popup values through canonical history and preserves saved imported criteria', async function () {
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
      'column-header-filter-ui.xlsx',
    )
    const values = [
      ['Name', 'Amount'],
      ['North', 12.5],
      ['South', -30],
      ['East', 18],
    ]
    const fixture = await applyCellEditsToXlsx(await manualWorkbook(), [
      ...values.flatMap((row, r) =>
        row.map((value, c) => ({
          sheetName: 'Data',
          row: r,
          column: c,
          writeValue: true,
          cell: { value },
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
            if (!(await page.title()).includes('column-header-filter-ui.xlsx')) continue
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
      await gotoReference(frame, 'Data!A1:B4')
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
      const before = structuredClone(document.engine.state)
      const revision = document.engine.descriptor.revision
      const panel = frame.locator('[data-u-comp="sheets-filter-panel"]').filter({
        has: frame.locator('[data-u-comp="sheets-filter-panel-footer"]'),
      })
      await expect
        .poll(() =>
          frame!
            .locator('.workbook-grid canvas')
            .evaluateAll((elements) =>
              Math.max(...elements.map((canvas) => canvas.getBoundingClientRect().height)),
            ),
        )
        .toBeGreaterThan(150)
      const canvasIndex = await frame.locator('.workbook-grid canvas').evaluateAll(
        (elements) =>
          elements
            .map((canvas, index) => ({
              index,
              area: canvas.getBoundingClientRect().width * canvas.getBoundingClientRect().height,
            }))
            .sort((a, b) => b.area - a.area)[0]!.index,
      )
      const canvas = frame.locator('.workbook-grid canvas').nth(canvasIndex)
      const box = (await canvas.boundingBox())!
      await frame.page().mouse.click(box.x + 46 + 140 - 9, box.y + 20 + 20 - 9)
      await expect(panel).toBeVisible()
      await expect(panel.getByText('North', { exact: true })).toBeVisible()
      for (const value of ['East', 'South'])
        await panel
          .getByText(value, { exact: true })
          .locator('xpath=ancestor::div[label[@data-u-comp="checkbox"]][1]')
          .locator('label[data-u-comp="checkbox"]')
          .click()
      await panel.getByRole('button', { name: 'Confirm', exact: true }).click()
      await document.bridge!.request('view.flush', {})
      await expect(frame.getByRole('alert')).toHaveCount(0)
      await expect(panel).toHaveCount(0)
      const after = structuredClone(document.engine.state)
      assert.equal(document.engine.descriptor.revision, revision + 1)
      assert.deepEqual((after.sheets[0]!.features!.filterEdits as any).filter.columns, [
        { colId: 0, values: ['North'], blank: false },
      ])
      assert.deepEqual((after.sheets[0]!.features!.filterEdits as any).hiddenRows, [2, 3])
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, before)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, after)
      await vscode.commands.executeCommand('workbench.action.files.save')
      const zip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
      assert.match(xml, /<autoFilter ref="A1:B4"/)
      assert.match(xml, /<filter val="North"/)
      assert.equal(
        await zip.file('customXml/manual-sentinel.xml')!.async('string'),
        '<preserve id="manual-fixture">unchanged</preserve>',
      )
      const session = document.sessionId
      await vscode.commands.executeCommand('workbench.action.closeActiveEditor')
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      const reopened = api.getDocument(uri)!
      await reopened.whenViewReady()
      assert.notEqual(reopened.sessionId, session)
      assert.equal(reopened.engine.descriptor.dirty, false)
      const snapshot = reopened.snapshot()
      try {
        const result = await snapshot.read(reopened.engine.state.sheets[0]!.id, 'A1:B4')
        assert.equal(
          result.cells.find((cell) => cell.address === 'A2')!.style!.fillColor,
          '#CCEEFF',
        )
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
