import { backupWorkbookDocument } from '../../src/extension/backup'
import type { WorkbookSource } from '@genoffice/workbook-session/ports'
import { formatAddress } from '@genoffice/xlsx-gateway/domain/cell-address'
import { randomUUID } from 'node:crypto'
import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import JSZip from 'jszip'
import { suite, test } from 'mocha'
import { chromium, type Frame } from '@playwright/test'
import {
  planCellEditsToXlsx,
  createBufferEntrySource,
  assembleWithJsZip,
} from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'

suite('Desktop Floating Charts UI', () => {
  test('moves source drawings through native history and preserves chart XML before cascading deletion', async function () {
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
      'floating-charts-ui.xlsx',
    )
    const anchor = {
      fromRow: 0,
      fromColumn: 3,
      fromRowOffset: 0,
      fromColumnOffset: 0,
      toRow: 12,
      toColumn: 8,
      toRowOffset: 0,
      toColumnOffset: 0,
    }
    const originalBase = await manualWorkbook()
    const originalPlan = await planCellEditsToXlsx(
      await createBufferEntrySource(originalBase),
      [],
      [],
      [],
      undefined,
      [],
      [],
      [],
      [],
      [],
      null,
      [
        {
          sheetName: 'Data',
          anchor,
          chart: {
            chartType: 'column',
            title: 'Source chart',
            series: [{ name: 'Sales', categories: ['North', 'South'], values: [12.5, 9] }],
          },
        },
      ],
    )
    const original = await assembleWithJsZip(originalBase, originalPlan)
    await vscode.workspace.fs.writeFile(uri, original.buffer)
    const browser = await chromium.connectOverCDP(`http://127.0.0.1:${port}`)
    const getFrame = async (): Promise<Frame> => {
      const started = Date.now()
      while (Date.now() - started < 20_000) {
        for (const context of browser.contexts())
          for (const page of context.pages()) {
            if (!(await page.title()).includes('floating-charts-ui.xlsx')) continue
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
    let releaseRange: (() => void) | undefined
    let restoreRange: (() => void) | undefined
    try {
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      const document = api.getDocument(uri)!
      await document.whenViewReady()
      const frame = await getFrame()
      const visual = frame.locator('.xlsx-print-visual .shape-editable').first()
      await visual.waitFor()
      const initial = structuredClone(document.engine.state)
      const source = (initial.features!.originalMetadata as any).visuals[0]
      const beforeRect = (await visual.boundingBox())!
      const title = (await visual.locator('figcaption').boundingBox())!
      await frame.page().mouse.move(title.x + 50, title.y + 8)
      await frame.page().mouse.down()
      await frame.page().mouse.move(title.x + 86, title.y + 32, { steps: 8 })
      await frame.page().mouse.up()
      await document.bridge!.request('view.flush', {})
      assert.equal(document.engine.descriptor.revision, 1)
      const moved = structuredClone(document.engine.state)
      assert.deepEqual(moved.sheets[0]!.cells, initial.sheets[0]!.cells)
      assert.deepEqual(moved.features, initial.features)
      const movedRect = (await visual.boundingBox())!
      assert.ok(Math.abs(movedRect.x - beforeRect.x - 36) < 2)
      assert.ok(Math.abs(movedRect.y - beforeRect.y - 24) < 2)
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, initial)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, moved)
      const navigate = async (reference: string) => {
        await vscode.commands.executeCommand('genoffice.workbook.goto')
        const dialog = frame.getByRole('dialog', { name: 'Go To', exact: true })
        await dialog.getByRole('textbox', { name: 'Reference', exact: true }).fill(reference)
        await dialog.getByRole('button', { name: 'Go', exact: true }).click()
        await dialog.waitFor({ state: 'hidden' })
      }
      const movedAnchor = (moved.sheets[0]!.features!.visualEdits as any)[source.id].anchor
      await navigate(`Data!${formatAddress(movedAnchor.toRow, movedAnchor.toColumn)}`)
      await visual.locator('[data-corner="nw"]').waitFor({ state: 'visible' })
      const hit = await visual.locator('[data-corner="nw"]').evaluate((element) => {
        const rect = element.getBoundingClientRect()
        return {
          rect: rect.toJSON(),
          viewport: { width: window.innerWidth, height: window.innerHeight },
          target: window.document
            .elementFromPoint(rect.x + rect.width / 2, rect.y + rect.height / 2)
            ?.outerHTML.slice(0, 400),
        }
      })
      assert.ok(hit.target?.includes('data-corner="nw"'), JSON.stringify(hit))
      const box = (await visual.locator('[data-corner="nw"]').boundingBox())!
      const oldSize = (await visual.boundingBox())!
      await frame.page().mouse.move(box.x + box.width / 2, box.y + box.height / 2)
      await frame.page().mouse.down()
      await frame
        .page()
        .mouse.move(box.x + box.width / 2 - 30, box.y + box.height / 2 - 20, { steps: 8 })
      await frame.page().mouse.up()
      await document.bridge!.request('view.flush', {})
      assert.equal(document.engine.descriptor.revision, 4)
      const resized = structuredClone(document.engine.state)
      const resizedRect = (await visual.boundingBox())!
      assert.ok(Math.abs(resizedRect.width - oldSize.width - 30) < 2)
      assert.ok(Math.abs(resizedRect.height - oldSize.height - 20) < 2)
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, moved)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, resized)
      const acceptedAnchor = (resized.sheets[0]!.features!.visualEdits as any)[source.id].anchor
      await navigate('Data!A1')
      const savedRect = (await visual.boundingBox())!
      await vscode.commands.executeCommand('workbench.action.files.save')
      const saved = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const originalZip = await JSZip.loadAsync(original.buffer)
      const chartPath = Object.keys(originalZip.files).find((path) =>
        /^xl\/charts\/chart[^/]+\.xml$/.test(path),
      )!
      assert.equal(
        await saved.file(chartPath)!.async('string'),
        await originalZip.file(chartPath)!.async('string'),
      )
      assert.equal(
        await saved.file('xl/worksheets/sheet1.xml')!.async('string'),
        await originalZip.file('xl/worksheets/sheet1.xml')!.async('string'),
      )
      assert.equal(
        await saved.file('customXml/manual-sentinel.xml')!.async('string'),
        '<preserve id="manual-fixture">unchanged</preserve>',
      )
      const drawingXml = await saved.file(source.drawingPath)!.async('string')
      assert.ok(drawingXml.includes(`<xdr:colOff>${acceptedAnchor.fromColumnOffset}</xdr:colOff>`))
      const session = document.sessionId
      await vscode.commands.executeCommand('workbench.action.closeActiveEditor')
      const started = Date.now()
      while (api.getDocument(uri) && Date.now() - started < 10_000)
        await new Promise((resolve) => setTimeout(resolve, 50))
      assert.equal(api.getDocument(uri), undefined)
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      let reopened = api.getDocument(uri)!
      await reopened.whenViewReady()
      assert.notEqual(reopened.sessionId, session)
      let sourceAfter = (reopened.engine.state.features!.originalMetadata as any).visuals[0]
      for (const key of Object.keys(acceptedAnchor))
        assert.equal(sourceAfter.anchor[key], acceptedAnchor[key])
      let reopenedFrame = await getFrame()
      let reopenedVisual = reopenedFrame.locator('.xlsx-print-visual .shape-editable').first()
      await reopenedVisual.waitFor()
      const reopenedRect = (await reopenedVisual.boundingBox())!
      assert.ok(Math.abs(reopenedRect.x - savedRect.x) < 2)
      assert.ok(Math.abs(reopenedRect.y - savedRect.y) < 2)
      assert.ok(Math.abs(reopenedRect.width - savedRect.width) < 2)
      assert.ok(Math.abs(reopenedRect.height - savedRect.height) < 2)
      assert.equal(reopened.engine.descriptor.revision, 0)
      assert.equal(reopened.engine.descriptor.dirty, false)
      const beforeProperties = structuredClone(reopened.engine.state)
      await reopenedFrame.locator('.chart-edit-button').click()
      const properties = reopenedFrame.locator('.chart-editor')
      await properties.locator('.chart-editor-title').fill('Edited original sales')
      await properties.locator('.chart-editor-type').click()
      await reopenedFrame.locator('button[role=option][data-value=line]').click()
      await properties.getByRole('listbox').waitFor({ state: 'hidden', timeout: 5000 })
      await properties.getByRole('button', { name: 'Sales', exact: true }).click()
      await reopenedFrame
        .getByRole('grid', { name: 'Standard Colors', exact: true })
        .getByRole('gridcell', { name: 'Red', exact: true })
        .click()
      await properties.locator('.chart-editor-range').nth(0).fill('Other!A1')
      await properties.locator('.chart-editor-range').nth(1).fill('Data!A1')
      const sourcePort = (reopened as unknown as { source: WorkbookSource }).source
      const originalRead = sourcePort.readBaseRange.bind(sourcePort)
      const heldRead = new Promise<void>((resolve) => {
        releaseRange = resolve
      })
      let heldOnce = false
      const otherId = reopened.engine.state.sheets[1]!.originalSheetId
      sourcePort.readBaseRange = async (sessionId, sheetId, range) => {
        if (!heldOnce && sheetId === otherId) {
          heldOnce = true
          await heldRead
        }
        return originalRead(sessionId, sheetId, range)
      }
      restoreRange = () => {
        sourcePort.readBaseRange = originalRead
      }
      await properties.getByRole('button', { name: 'Apply', exact: true }).click()
      const waitRead = Date.now()
      while (!heldOnce && Date.now() - waitRead < 5000)
        await new Promise((resolve) => setTimeout(resolve, 25))
      assert.ok(heldOnce, 'host read is pending before the chart transaction exists')
      assert.equal(await properties.locator('.chart-editor-cancel').isDisabled(), true)
      let saveFinished = false
      let saveError: unknown
      const pendingSave = vscode.commands.executeCommand('workbench.action.files.save').then(
        () => {
          saveFinished = true
        },
        (error) => {
          saveFinished = true
          saveError = error
        },
      )
      await properties.locator('.chart-editor-apply').dispatchEvent('click')
      await reopenedFrame.page().keyboard.press('Escape')
      assert.equal(await properties.isVisible(), true)
      assert.equal(reopened.engine.descriptor.revision, 0)
      assert.equal(saveFinished, false, 'native Save waits for the held range read')
      releaseRange!()
      await pendingSave
      assert.equal(saveError, undefined)
      restoreRange()
      restoreRange = undefined

      await properties.waitFor({ state: 'hidden' })
      await reopened.bridge!.request('view.flush', {})
      assert.equal(reopened.engine.descriptor.dirty, false)
      const withProperties = structuredClone(reopened.engine.state)
      assert.deepEqual((withProperties.sheets[0]!.features!.chartEdits as any)[chartPath], {
        chartPath,
        title: 'Edited original sales',
        chartType: 'line',
        seriesColors: { '0': '#ff0000' },
        series: [
          {
            index: 0,
            values: [7],
            valuesRef: "'Other'!$A$1:$A$1",
            categories: ['00123'],
            categoriesRef: "'Data'!$A$1:$A$1",
          },
        ],
      })
      assert.deepEqual(withProperties.features, beforeProperties.features)
      assert.deepEqual(withProperties.sheets[0]!.cells, beforeProperties.sheets[0]!.cells)
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(reopened.engine.state, beforeProperties)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(reopened.engine.state, withProperties)
      const chartFigure = reopenedFrame.locator('.xlsx-print-visual .xlsx-chart').first()
      await chartFigure.click({ button: 'right', position: { x: 50, y: 15 } })
      await reopenedFrame
        .locator('.chart-menu')
        .first()
        .getByRole('button', { name: 'Add Data Labels', exact: true })
        .click()
      await reopened.bridge!.request('view.flush', {})
      await chartFigure.click({ button: 'right', position: { x: 50, y: 15 } })
      await reopenedFrame
        .locator('.chart-menu')
        .first()
        .getByRole('button', { name: 'Format Selection…', exact: true })
        .click()
      const formatPane = reopenedFrame.locator('.chart-format-pane')
      await formatPane.waitFor()
      const beforeLegend = structuredClone(reopened.engine.state)
      await formatPane.getByRole('button', { name: 'Legend Position', exact: true }).click()
      await formatPane.getByRole('option', { name: 'Bottom', exact: true }).click()
      await reopened.bridge!.request('view.flush', {})
      const afterLegend = structuredClone(reopened.engine.state)
      assert.equal((afterLegend.sheets[0]!.features!.chartEdits as any)[chartPath].legend, 'bottom')
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(reopened.engine.state, beforeLegend)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(reopened.engine.state, afterLegend)
      await formatPane.getByLabel('Gridlines', { exact: true }).uncheck()
      await reopened.bridge!.request('view.flush', {})
      await formatPane.getByLabel('Minimum', { exact: true }).fill('0')
      await formatPane.getByLabel('Minimum', { exact: true }).press('Enter')
      await reopened.bridge!.request('view.flush', {})
      await formatPane.getByLabel('Maximum', { exact: true }).fill('20')
      await formatPane.getByLabel('Maximum', { exact: true }).press('Enter')
      await reopened.bridge!.request('view.flush', {})
      await formatPane.getByLabel('Label Number Format', { exact: true }).fill('0.00')
      await formatPane.getByLabel('Label Number Format', { exact: true }).press('Enter')
      await reopened.bridge!.request('view.flush', {})
      const paneTitle = formatPane.getByRole('textbox', { name: 'Title', exact: true })
      await paneTitle.fill('Format pane original')
      const backupDirectory = vscode.Uri.joinPath(
        vscode.workspace.workspaceFolders![0]!.uri,
        'chart-title-backup',
      )
      const cancellation = new vscode.CancellationTokenSource()
      const beforeBackup = structuredClone(reopened.engine.state)
      const previousBackupId = reopened.backupId
      // Use an independent test slot; the existing native slot belongs to host storage.
      reopened.backupId = undefined
      try {
        const backup = await backupWorkbookDocument(
          reopened,
          vscode.Uri.joinPath(backupDirectory, 'backup'),
          cancellation.token,
        )
        await backup.delete()
        assert.deepEqual(reopened.engine.state, beforeBackup)
        assert.equal(
          await paneTitle.evaluate((input) => window.document.activeElement === input),
          true,
        )
        assert.equal(await paneTitle.inputValue(), 'Format pane original')
      } finally {
        cancellation.dispose()
        reopened.backupId = previousBackupId
        await vscode.workspace.fs.delete(backupDirectory, { recursive: true }).then(
          () => {},
          () => {},
        )
      }
      await paneTitle.fill('')
      await paneTitle.pressSequentially('Format pane original title', { delay: 10 })
      assert.equal(
        await paneTitle.evaluate((input) => window.document.activeElement === input),
        true,
      )
      // Do not blur the input: native Save must publish its current draft.
      await vscode.commands.executeCommand('workbench.action.files.save')
      const formatState = reopened.engine.state
      assert.equal(
        (formatState.sheets[0]!.features!.chartEdits as any)[chartPath].title,
        'Format pane original title',
      )
      assert.deepEqual(formatState.sheets[0]!.cells, beforeProperties.sheets[0]!.cells)
      assert.equal(reopened.engine.descriptor.dirty, false)
      await formatPane.getByRole('button', { name: 'Close', exact: true }).click()
      await formatPane.waitFor({ state: 'hidden' })
      const propertySaved = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const propertyXml = await propertySaved.file(chartPath)!.async('string')
      assert.ok(propertyXml.includes('Format pane original title'))
      assert.ok(propertyXml.includes('<c:legendPos val="b"/>'))
      assert.ok(propertyXml.includes('<c:min val="0"/>'))
      assert.ok(propertyXml.includes('<c:max val="20"/>'))
      assert.ok(propertyXml.includes('formatCode="0.00"'))
      assert.ok(!propertyXml.includes('<c:majorGridlines'))
      assert.ok(propertyXml.includes('<c:lineChart>'))
      assert.ok(/FF0000/i.test(propertyXml))
      assert.ok(propertyXml.includes("'Other'!$A$1:$A$1"))
      assert.ok(propertyXml.includes("'Data'!$A$1:$A$1"))
      assert.ok(propertyXml.includes('<c:v>00123</c:v>'))
      assert.equal(
        await propertySaved.file(source.drawingPath)!.async('string'),
        await saved.file(source.drawingPath)!.async('string'),
      )
      assert.equal(
        await propertySaved.file('xl/worksheets/sheet1.xml')!.async('string'),
        await saved.file('xl/worksheets/sheet1.xml')!.async('string'),
      )
      const propertySession = reopened.sessionId
      await vscode.commands.executeCommand('workbench.action.closeActiveEditor')
      const propertyClosed = Date.now()
      while (api.getDocument(uri) && Date.now() - propertyClosed < 10_000)
        await new Promise((resolve) => setTimeout(resolve, 50))
      assert.equal(api.getDocument(uri), undefined)
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      reopened = api.getDocument(uri)!
      await reopened.whenViewReady()
      assert.notEqual(reopened.sessionId, propertySession)
      sourceAfter = (reopened.engine.state.features!.originalMetadata as any).visuals[0]
      assert.equal(sourceAfter.chart.title, 'Format pane original title')
      assert.equal(sourceAfter.chart.legend, 'bottom')
      assert.equal(sourceAfter.chart.dataLabels, 'value')
      assert.equal(sourceAfter.chart.dataLabelFormat, '0.00')
      assert.equal(sourceAfter.chart.gridlines, false)
      assert.deepEqual(sourceAfter.chart.valueAxis, { min: 0, max: 20 })
      assert.deepEqual(sourceAfter.chart.chartTypes, ['lineChart'])
      assert.ok(/^#ff0000$/i.test(sourceAfter.chart.series[0].color))
      assert.deepEqual(sourceAfter.chart.series[0].values, [7])
      assert.deepEqual(sourceAfter.chart.series[0].categories, ['00123'])
      assert.equal(sourceAfter.chart.series[0].valuesRef, "'Other'!$A$1:$A$1")
      assert.equal(sourceAfter.chart.series[0].categoriesRef, "'Data'!$A$1:$A$1")
      reopenedFrame = await getFrame()
      await reopenedFrame
        .getByRole('img', { name: 'Format pane original title', exact: true })
        .waitFor()
      reopenedVisual = reopenedFrame.locator('.xlsx-print-visual .shape-editable').first()
      assert.equal(reopened.engine.descriptor.revision, 0)
      assert.equal(reopened.engine.descriptor.dirty, false)
      // An AI transaction uses the same original-chart overlay and history.
      const beforeAi = structuredClone(reopened.engine.state)
      await reopened.engine.submit({
        schemaVersion: 1,
        sessionId: reopened.sessionId,
        transactionId: randomUUID(),
        expectedRevision: 0,
        origin: 'agent',
        summary: 'AI original chart title',
        operations: [{ op: 'edit_chart', chartPath, title: 'AI original title' }],
      })
      await reopened.bridge!.request('view.flush', {})
      await reopenedFrame.getByRole('img', { name: 'AI original title', exact: true }).waitFor()
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(reopened.engine.state, beforeAi)
      await reopenedVisual.click({ button: 'right', position: { x: 50, y: 15 } })
      await reopenedFrame
        .locator('.chart-menu')
        .first()
        .getByRole('button', { name: 'Select Data…', exact: true })
        .click()
      const selectionDialog = reopenedFrame.locator('.select-data-dialog')
      const seriesBefore = structuredClone(reopened.engine.state)
      await selectionDialog.locator('.select-data-rows input').nth(0).fill('Native retained')
      await selectionDialog.getByRole('button', { name: '＋ Add Series', exact: true }).click()
      const addedRow = selectionDialog.locator('.select-data-rows .select-data-row').nth(1)
      await addedRow.locator('input').nth(0).fill('Native added')
      await addedRow.locator('input').nth(1).fill("'Other'!A1")
      await addedRow.locator('input').nth(2).fill("'Data'!A1")
      await addedRow.getByRole('button', { name: 'Move Up', exact: true }).click()
      await selectionDialog.getByRole('button', { name: 'OK', exact: true }).click()
      await selectionDialog.waitFor({ state: 'hidden' })
      await reopened.bridge!.request('view.flush', {})
      const seriesAfter = structuredClone(reopened.engine.state)
      assert.deepEqual(
        (seriesAfter.sheets[0]!.features!.chartEdits as any)[chartPath].seriesSet.map(
          (s: any) => s.name,
        ),
        ['Native added', 'Native retained'],
      )
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(reopened.engine.state, seriesBefore)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(reopened.engine.state, seriesAfter)
      await vscode.commands.executeCommand('workbench.action.files.save')
      const seriesSaved = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const seriesXml = await seriesSaved.file(chartPath)!.async('string')
      assert.equal((seriesXml.match(/<c:ser>/g) ?? []).length, 2)
      assert.ok(seriesXml.includes('Native added'))
      assert.ok(seriesXml.includes('Native retained'))
      const selectionSession = reopened.sessionId
      await vscode.commands.executeCommand('workbench.action.closeActiveEditor')
      const selectionClosed = Date.now()
      while (api.getDocument(uri) && Date.now() - selectionClosed < 10_000)
        await new Promise((resolve) => setTimeout(resolve, 50))
      assert.equal(api.getDocument(uri), undefined)
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      reopened = api.getDocument(uri)!
      await reopened.whenViewReady()
      assert.notEqual(reopened.sessionId, selectionSession)
      sourceAfter = (reopened.engine.state.features!.originalMetadata as any).visuals[0]
      assert.deepEqual(
        sourceAfter.chart.series.map((s: any) => s.name),
        ['Native added', 'Native retained'],
      )
      assert.deepEqual(
        sourceAfter.chart.series.map((s: any) => s.values),
        [[7], [7]],
      )
      assert.ok(/^#ff0000$/i.test(sourceAfter.chart.series[1].color))
      reopenedFrame = await getFrame()
      reopenedVisual = reopenedFrame.locator('.xlsx-print-visual .shape-editable').first()
      await reopenedVisual.waitFor()
      await reopenedVisual.click({ position: { x: 50, y: 15 } })
      const chartDesign = reopenedFrame.getByRole('tab', { name: 'Chart Design', exact: true })
      await chartDesign.click()
      const designRibbon = reopenedFrame.locator('#ribbon-panel-chartDesign')
      const beforeLayout = structuredClone(reopened.engine.state)
      const layoutRevision = reopened.engine.descriptor.revision
      await designRibbon.getByRole('button', { name: 'Quick Layout', exact: true }).click()
      await designRibbon
        .getByRole('option', { name: 'Layout 3 — Legend bottom, no labels', exact: true })
        .click()
      await reopened.bridge!.request('view.flush', {})
      assert.equal(reopened.engine.descriptor.revision, layoutRevision + 1)
      const afterLayout = structuredClone(reopened.engine.state)
      assert.equal((afterLayout.sheets[0]!.features!.chartEdits as any)[chartPath].legend, 'bottom')
      assert.equal(
        (afterLayout.sheets[0]!.features!.chartEdits as any)[chartPath].dataLabels,
        'none',
      )
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(reopened.engine.state, beforeLayout)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(reopened.engine.state, afterLayout)
      await designRibbon.getByRole('button', { name: 'Add Chart Element', exact: true }).click()
      await designRibbon.getByRole('option', { name: 'Axis Title: Category…', exact: true }).click()
      const axisDialog = reopenedFrame.getByRole('dialog', {
        name: 'Category Axis Title',
        exact: true,
      })
      await axisDialog.getByRole('textbox').fill('Native chart ribbon category')
      await axisDialog.getByRole('button', { name: 'OK', exact: true }).click()
      await axisDialog.waitFor({ state: 'hidden' })
      await vscode.commands.executeCommand('workbench.action.files.save')
      const designSaved = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      assert.ok(
        (await designSaved.file(chartPath)!.async('string')).includes(
          'Native chart ribbon category',
        ),
      )
      const designSession = reopened.sessionId
      await vscode.commands.executeCommand('workbench.action.closeActiveEditor')
      const designClosed = Date.now()
      while (api.getDocument(uri) && Date.now() - designClosed < 10_000)
        await new Promise((resolve) => setTimeout(resolve, 50))
      assert.equal(api.getDocument(uri), undefined)
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      reopened = api.getDocument(uri)!
      await reopened.whenViewReady()
      assert.notEqual(reopened.sessionId, designSession)
      sourceAfter = (reopened.engine.state.features!.originalMetadata as any).visuals[0]
      assert.equal(sourceAfter.chart.axisTitles.category, 'Native chart ribbon category')
      assert.equal(sourceAfter.chart.legend, 'bottom')
      assert.equal(sourceAfter.chart.dataLabels ?? 'none', 'none')
      reopenedFrame = await getFrame()
      reopenedVisual = reopenedFrame.locator('.xlsx-print-visual .shape-editable').first()
      await reopenedVisual.waitFor()
      await reopenedVisual.click({ position: { x: 50, y: 15 } })
      const beforePropertiesRoot = structuredClone(reopened.engine.state)
      const chartBeforeProperties = await (
        await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      )
        .file(chartPath)!
        .async('string')
      const anchorBeforeProperties = structuredClone(sourceAfter.anchor)
      await reopenedVisual.click({ button: 'right', position: { x: 50, y: 15 } })
      await reopenedFrame
        .locator('.chart-menu')
        .first()
        .getByRole('button', { name: 'Edit Alt Text…', exact: true })
        .click()
      const altDialog = reopenedFrame.getByRole('dialog', { name: 'Alt Text', exact: true })
      await altDialog.getByRole('textbox').fill('Native alt & sales')
      await altDialog.getByRole('button', { name: 'OK', exact: true }).click()
      await altDialog.waitFor({ state: 'hidden' })
      await reopened.bridge!.request('view.flush', {})
      const afterAlt = structuredClone(reopened.engine.state)
      assert.equal(
        (afterAlt.sheets[0]!.features!.visualEdits as any)[sourceAfter.id].altText,
        'Native alt & sales',
      )
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(reopened.engine.state, beforePropertiesRoot)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(reopened.engine.state, afterAlt)
      await reopenedVisual.click({ button: 'right', position: { x: 50, y: 15 } })
      await reopenedFrame
        .locator('.chart-menu')
        .first()
        .getByRole('button', { name: 'Size and Properties…', exact: true })
        .click()
      const sizeDialog = reopenedFrame.getByRole('dialog', {
        name: 'Size and Properties',
        exact: true,
      })
      await sizeDialog
        .getByRole('radio', { name: "Don't move or size with cells", exact: true })
        .check()
      await sizeDialog.getByRole('button', { name: 'OK', exact: true }).click()
      await sizeDialog.waitFor({ state: 'hidden' })
      await vscode.commands.executeCommand('workbench.action.files.save')
      const propertiesArchive = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      assert.equal(await propertiesArchive.file(chartPath)!.async('string'), chartBeforeProperties)
      const drawingProperties = await propertiesArchive
        .file(sourceAfter.drawingPath)!
        .async('string')
      assert.ok(drawingProperties.includes('descr="Native alt &amp; sales"'))
      assert.ok(drawingProperties.includes('editAs="absolute"'))
      const visualSession = reopened.sessionId
      await vscode.commands.executeCommand('workbench.action.closeActiveEditor')
      const visualClosed = Date.now()
      while (api.getDocument(uri) && Date.now() - visualClosed < 10_000)
        await new Promise((resolve) => setTimeout(resolve, 50))
      assert.equal(api.getDocument(uri), undefined)
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      reopened = api.getDocument(uri)!
      await reopened.whenViewReady()
      assert.notEqual(reopened.sessionId, visualSession)
      sourceAfter = (reopened.engine.state.features!.originalMetadata as any).visuals[0]
      assert.equal(sourceAfter.altText, 'Native alt & sales')
      assert.equal(sourceAfter.editAs, 'absolute')
      assert.deepEqual(sourceAfter.anchor, anchorBeforeProperties)
      reopenedFrame = await getFrame()
      reopenedVisual = reopenedFrame.locator('.xlsx-print-visual .shape-editable').first()
      await reopenedVisual.waitFor()
      await reopenedVisual.click({ position: { x: 50, y: 15 } })
      const beforeAgentVisual = structuredClone(reopened.engine.state)
      await reopened.engine.submit({
        schemaVersion: 1,
        sessionId: reopened.sessionId,
        transactionId: randomUUID(),
        expectedRevision: reopened.engine.descriptor.revision,
        origin: 'agent',
        summary: 'AI visual description',
        operations: [
          { op: 'edit_visual', visualId: sourceAfter.id, altText: 'AI temporary description' },
        ],
      })
      await reopened.bridge!.request('view.flush', {})
      assert.equal(
        (reopened.engine.state.sheets[0]!.features!.visualEdits as any)[sourceAfter.id].altText,
        'AI temporary description',
      )
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(reopened.engine.state, beforeAgentVisual)
      const beforeDelete = structuredClone(reopened.engine.state)
      await reopenedVisual.press('Delete')
      await reopenedVisual.waitFor({ state: 'detached' })
      await reopened.bridge!.request('view.flush', {})
      const deleted = structuredClone(reopened.engine.state)
      assert.deepEqual((deleted.sheets[0]!.features!.visualEdits as any)[sourceAfter.id], {
        remove: true,
      })
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(reopened.engine.state, beforeDelete)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(reopened.engine.state, deleted)
      await vscode.commands.executeCommand('workbench.action.files.save')
      const removed = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      assert.equal(removed.file(chartPath), null)
      const remainingDrawing = removed.file(sourceAfter.drawingPath)
      if (remainingDrawing) {
        assert.ok(!(await remainingDrawing.async('string')).includes('<c:chart'))
      } else {
        const worksheet = await removed.file('xl/worksheets/sheet1.xml')!.async('string')
        assert.doesNotMatch(worksheet, /<drawing\b/)
        const relationships = await removed
          .file('xl/worksheets/_rels/sheet1.xml.rels')
          ?.async('string')
        assert.ok(!(relationships ?? '').includes(sourceAfter.drawingPath.split('/').pop()))
      }
      const contentTypes = await removed.file('[Content_Types].xml')!.async('string')
      assert.ok(!contentTypes.includes(`/${chartPath}`))
      if (!remainingDrawing) assert.ok(!contentTypes.includes(`/${sourceAfter.drawingPath}`))
      assert.equal(
        await removed.file('customXml/manual-sentinel.xml')!.async('string'),
        '<preserve id="manual-fixture">unchanged</preserve>',
      )
    } finally {
      releaseRange?.()
      restoreRange?.()
      if (api.getDocument(uri))
        await vscode.commands.executeCommand('workbench.action.files.save', uri)
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
      await browser.close()
    }
  })
})
