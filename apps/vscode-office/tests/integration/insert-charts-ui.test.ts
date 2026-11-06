import * as assert from 'node:assert/strict'
import { randomUUID } from 'node:crypto'
import * as vscode from 'vscode'
import JSZip from 'jszip'
import { suite, test } from 'mocha'
import { chromium, type Frame } from '@playwright/test'
import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'
import { gotoReference } from './goto-helper'

suite('Desktop Insert Charts UI', () => {
  test('inserts every shared desktop chart family with saved references, exact history and reopened rendering', async function () {
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
      'insert-charts-ui.xlsx',
    )
    const values = [
      ['Year', 'Revenue', 'Cost'],
      [2020, 12, 6],
      [2021, 8, 10],
      [2022, 18, 14],
    ]
    const original = await applyCellEditsToXlsx(
      await manualWorkbook(),
      values.flatMap((row, r) =>
        row.map((value, c) => ({
          sheetName: 'Data',
          row: 77 + r,
          column: c,
          writeValue: true,
          cell: { value },
        })),
      ),
    )
    await vscode.workspace.fs.writeFile(uri, original.buffer)
    const browser = await chromium.connectOverCDP(`http://127.0.0.1:${port}`)
    const getFrame = async (): Promise<Frame> => {
      const started = Date.now()
      while (Date.now() - started < 20_000) {
        for (const context of browser.contexts())
          for (const page of context.pages()) {
            if (!(await page.title()).includes('insert-charts-ui.xlsx')) continue
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
      throw new Error('Workbook frame did not load')
    }
    try {
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      const document = api.getDocument(uri)!
      await document.whenViewReady()
      const frame = await getFrame()
      await gotoReference(frame, 'Data!A78:C81')
      await frame.getByRole('tab', { name: 'Insert', exact: true }).click()
      const initial = structuredClone(document.engine.state)
      const labels = [
        'Column',
        'Bar',
        'Line',
        'Area',
        'Pie',
        'Scatter',
        'Radar',
        'Doughnut',
        'Combo (column + line)',
      ]
      for (const [index, label] of labels.entries()) {
        const before = structuredClone(document.engine.state)
        await frame
          .locator('.workbook-desktop-header')
          .getByRole('button', { name: `${label} chart — from selection`, exact: true })
          .click()
        await frame.locator('.xlsx-print-visual').nth(index).waitFor()
        await document.bridge!.request('view.flush', {})
        const after = structuredClone(document.engine.state)
        assert.deepEqual(after.sheets[0]!.cells, initial.sheets[0]!.cells)
        assert.equal((after.sheets[0]!.features!.charts as unknown[]).length, index + 1)
        await vscode.commands.executeCommand('undo')
        assert.deepEqual(document.engine.state, before)
        await vscode.commands.executeCommand('redo')
        assert.deepEqual(document.engine.state, after)
      }
      const charts = document.engine.state.sheets[0]!.features!.charts as any[]
      assert.deepEqual(
        charts.map((visual) => visual.chart.chartTypes),
        [
          ['barChart'],
          ['barChart'],
          ['lineChart'],
          ['areaChart'],
          ['pieChart'],
          ['scatterChart'],
          ['radarChart'],
          ['doughnutChart'],
          ['barChart', 'lineChart'],
        ],
      )
      for (const visual of charts) {
        assert.deepEqual(
          visual.chart.series.map((series: any) => series.values),
          [
            [12, 8, 18],
            [6, 10, 14],
          ],
        )
        assert.equal(visual.chart.series[0].valuesRef, "'Data'!$B$79:$B$81")
        assert.equal(visual.chart.series[0].categoriesRef, "'Data'!$A$79:$A$81")
      }
      // The last chart is on top; use actual pointer clicks on its shared
      // inline editor, including the portalled desktop color picker.
      // Move beyond the chart, then reveal its middle: Go To deliberately
      // leaves already-visible cells in place, and the right-edge pencil can
      // otherwise remain clipped outside the actual Webview viewport.
      for (const reference of ['Data!N78', 'Data!H78']) {
        await gotoReference(frame, reference)
      }
      const beforeProperties = structuredClone(document.engine.state)
      const editButton = frame.locator('.xlsx-print-visual').nth(8).locator('.chart-edit-button')
      const hit = await editButton.evaluate((button) => {
        const rect = button.getBoundingClientRect()
        return {
          rect: rect.toJSON(),
          width: innerWidth,
          height: innerHeight,
          hit: window.document
            .elementFromPoint(rect.x + rect.width / 2, rect.y + rect.height / 2)
            ?.outerHTML.slice(0, 300),
        }
      })
      assert.ok(hit.hit?.includes('chart-edit-button'), JSON.stringify(hit))
      await editButton.click()
      const properties = frame.locator('.chart-editor')
      await properties.locator('.chart-editor-title').fill('Native combo properties')
      await properties.getByRole('button', { name: 'Revenue', exact: true }).click()
      await frame
        .getByRole('grid', { name: 'Standard Colors', exact: true })
        .getByRole('gridcell', { name: 'Red', exact: true })
        .click()
      await properties.getByRole('button', { name: 'Apply', exact: true }).click()
      await properties.waitFor({ state: 'hidden' })
      await document.bridge!.request('view.flush', {})
      const afterProperties = structuredClone(document.engine.state)
      const editedCombo = (afterProperties.sheets[0]!.features!.charts as any[])[8]
      assert.equal(editedCombo.chart.title, 'Native combo properties')
      assert.deepEqual(editedCombo.chart.chartTypes, ['barChart', 'lineChart'])
      assert.equal(editedCombo.chart.series[0].color, '#ff0000')
      assert.deepEqual(afterProperties.sheets[0]!.cells, initial.sheets[0]!.cells)
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, beforeProperties)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, afterProperties)
      await gotoReference(frame, 'Data!A78:C81')
      await frame.locator('.workbook-desktop-header [data-tip="Recommended Charts"]').click()
      const dialog = frame.getByRole('dialog', { name: 'Recommended Charts', exact: true })
      await dialog.waitFor()
      await document.engine.submit({
        schemaVersion: 1,
        sessionId: document.sessionId,
        transactionId: randomUUID(),
        expectedRevision: document.engine.descriptor.revision,
        origin: 'manual',
        summary: 'Concurrent edit',
        operations: [
          { op: 'set_note', sheetId: initial.sheets[0]!.id, address: 'A1', text: 'Concurrent' },
        ],
      })
      const revision = document.engine.descriptor.revision
      await dialog.locator('.recommended-charts-grid button').first().click()
      await dialog.getByRole('alert').filter({ hasText: 'document changed' }).waitFor()
      assert.equal(document.engine.descriptor.revision, revision)
      await dialog.getByRole('button', { name: 'Close', exact: true }).click()
      await vscode.commands.executeCommand('undo')
      await vscode.commands.executeCommand('workbench.action.files.save')
      const saved = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const paths = Object.keys(saved.files).filter((path) =>
        /^xl\/charts\/chart[^/]+\.xml$/.test(path),
      )
      assert.equal(paths.length, 9)
      const xmls = await Promise.all(paths.map((path) => saved.file(path)!.async('string')))
      for (const tag of [
        'barChart',
        'lineChart',
        'areaChart',
        'pieChart',
        'scatterChart',
        'radarChart',
        'doughnutChart',
      ])
        assert.ok(
          xmls.some((xml) => xml.includes(`<c:${tag}>`)),
          tag,
        )
      assert.ok(xmls.some((xml) => xml.includes('<c:barChart>') && xml.includes('<c:lineChart>')))
      for (const xml of xmls) {
        assert.ok(xml.includes("'Data'!$B$79:$B$81"))
        assert.ok(xml.includes('<c:v>12</c:v>'))
      }
      assert.ok(xmls.some((xml) => xml.includes('Native combo properties') && /FF0000/i.test(xml)))
      const originalZip = await JSZip.loadAsync(original.buffer)
      const sourceXml = await originalZip.file('xl/worksheets/sheet1.xml')!.async('string')
      const savedXml = await saved.file('xl/worksheets/sheet1.xml')!.async('string')
      assert.equal(
        savedXml.match(/<sheetData[\s\S]*?<\/sheetData>/)![0],
        sourceXml.match(/<sheetData[\s\S]*?<\/sheetData>/)![0],
      )
      assert.equal(
        await saved.file('customXml/manual-sentinel.xml')!.async('string'),
        '<preserve id="manual-fixture">unchanged</preserve>',
      )
      const session = document.sessionId
      await vscode.commands.executeCommand('workbench.action.closeActiveEditor')
      const started = Date.now()
      while (api.getDocument(uri) && Date.now() - started < 10_000)
        await new Promise((resolve) => setTimeout(resolve, 50))
      assert.equal(api.getDocument(uri), undefined)
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      const reopened = api.getDocument(uri)!
      await reopened.whenViewReady()
      assert.notEqual(reopened.sessionId, session)
      const reopenedFrame = await getFrame()
      await vscode.commands.executeCommand('genoffice.workbook.goto')
      const reopenedGoTo = reopenedFrame.getByRole('dialog', { name: 'Go To', exact: true })
      await reopenedGoTo
        .getByRole('textbox', { name: 'Reference', exact: true })
        .fill('Data!A78:C81')
      await reopenedGoTo.getByRole('button', { name: 'Go', exact: true }).click()
      await reopenedGoTo.waitFor({ state: 'hidden' })
      await reopenedFrame.locator('.xlsx-print-visual').nth(8).waitFor()
      assert.equal(await reopenedFrame.locator('.xlsx-print-visual .chart-svg').count(), 9)
      await reopenedFrame
        .getByRole('img', { name: 'Native combo properties', exact: true })
        .waitFor()
      assert.equal(await reopenedFrame.locator('.chart-edit-button').count(), 9)
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
