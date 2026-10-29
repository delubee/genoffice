import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import JSZip from 'jszip'
import { suite, test } from 'mocha'
import { chromium, type Frame } from '@playwright/test'
import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'

suite('Desktop Advanced Filter UI', () => {
  test('applies raw-number AND criteria and preserves imported conditions through history and independent reopen', async function () {
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
      'advanced-filter-ui.xlsx',
    )
    const values = [
      ['Sales', 'Region'],
      [12.5, 'North'],
      [8, 'South'],
      [18, 'East'],
      ['12', 'West'],
      [null, 'Other'],
    ]
    const fixture = await applyCellEditsToXlsx(
      await manualWorkbook(),
      values.flatMap((row, r) =>
        row.map((value, c) => ({
          sheetName: 'Data',
          row: r,
          column: c,
          writeValue: true,
          cell: { value },
        })),
      ),
    )
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
            if (!(await page.title()).includes('advanced-filter-ui.xlsx')) continue
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
      await vscode.commands.executeCommand('genoffice.workbook.goto')
      const go = frame.getByRole('dialog', { name: 'Go To', exact: true })
      await go.getByRole('textbox', { name: 'Reference', exact: true }).fill('Data!A1:B6')
      await go.getByRole('button', { name: 'Go', exact: true }).click()
      await go.waitFor({ state: 'hidden' })
      await frame.locator('#ribbon-tab-data').click()
      await frame
        .locator('.workbook-shared-sort-filter')
        .getByRole('button', { name: /^Filter\b/ })
        .click()
      await document.bridge!.request('view.flush', {})
      const before = structuredClone(document.engine.state)
      const revision = document.engine.descriptor.revision
      await frame.getByRole('button', { name: 'Advanced', exact: true }).click()
      const dialog = frame.getByRole('dialog', { name: 'Advanced Filter', exact: true })
      await dialog.getByRole('button', { name: 'Condition 1', exact: true }).click()
      await frame.locator('[role=option][data-value="greaterThanOrEqual"]').click()
      await dialog.getByRole('textbox', { name: 'Value', exact: true }).nth(0).fill('10')
      await dialog.getByRole('button', { name: 'Condition 2 (optional)', exact: true }).click()
      await frame.locator('[role=option][data-value="lessThan"]').click()
      await dialog.getByRole('textbox', { name: 'Value', exact: true }).nth(1).fill('20')
      await dialog.getByRole('button', { name: 'OK', exact: true }).click()
      await document.bridge!.request('view.flush', {})
      if (await dialog.getByRole('alert').count())
        assert.fail(await dialog.getByRole('alert').innerText())
      await dialog.waitFor({ state: 'hidden' })
      assert.equal(document.engine.descriptor.revision, revision + 1)
      const after = structuredClone(document.engine.state)
      assert.deepEqual(
        (after.sheets[0]!.features!.filterEdits as { hiddenRows: number[] }).hiddenRows,
        [2, 4, 5],
      )
      assert.deepEqual(after.sheets[0]!.cells, before.sheets[0]!.cells)
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, before)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, after)
      await vscode.commands.executeCommand('workbench.action.files.save')
      const zip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
      assert.match(xml, /<customFilters and="1">/)
      assert.match(xml, /operator="greaterThanOrEqual" val="10"/)
      assert.match(xml, /operator="lessThan" val="20"/)
      for (const row of [3, 5, 6])
        assert.match(xml, new RegExp('<row\\b[^>]*r="' + row + '"[^>]*hidden="1"'))
      assert.equal(
        await zip.file('customXml/manual-sentinel.xml')!.async('string'),
        '<preserve id="manual-fixture">unchanged</preserve>',
      )
      const previousSession = document.sessionId
      await vscode.commands.executeCommand('workbench.action.closeActiveEditor')
      const closeDeadline = Date.now() + 10000
      while (api.getDocument(uri) && Date.now() < closeDeadline)
        await new Promise((resolve) => setTimeout(resolve, 50))
      assert.equal(api.getDocument(uri), undefined)
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      const reopened = api.getDocument(uri)!
      await reopened.whenViewReady()
      assert.notEqual(reopened.sessionId, previousSession)
      assert.equal(reopened.engine.descriptor.revision, 0)
      assert.equal(reopened.engine.descriptor.dirty, false)
      const snapshot = reopened.snapshot()
      try {
        const source = await snapshot.read(reopened.engine.state.sheets[0]!.id, 'A1:B6')
        const cells = new Map(source.cells.map((cell) => [cell.address, cell]))
        assert.equal(cells.get('A2')!.rawValue, 12.5)
        assert.equal(cells.get('A2')!.numberFormat, '0.00')
        assert.equal(cells.get('A2')!.style!.fillColor, '#CCEEFF')
        assert.equal(cells.get('A4')!.rawValue, 18)
        assert.equal(cells.get('A5')!.rawValue, '12')
        assert.equal(cells.get('A5')!.kind, 'string')
        const other = await snapshot.read(reopened.engine.state.sheets[1]!.id, 'A1')
        assert.equal(other.cells[0]!.rawValue, 7)
      } finally {
        snapshot.release()
      }
      let reopenedFrame: Frame | undefined
      const reopenDeadline = Date.now() + 20000
      while (!reopenedFrame && Date.now() < reopenDeadline) {
        for (const context of browser.contexts())
          for (const page of context.pages()) {
            if (!(await page.title()).includes('advanced-filter-ui.xlsx')) continue
            for (const candidate of page.frames())
              if (
                await candidate
                  .locator('#ribbon-tab-data')
                  .count()
                  .catch(() => 0)
              )
                reopenedFrame = candidate
          }
        if (!reopenedFrame) await new Promise((resolve) => setTimeout(resolve, 100))
      }
      assert.ok(reopenedFrame)
      const imported = structuredClone(reopened.engine.state)
      await reopenedFrame.locator('#ribbon-tab-data').click()
      await reopenedFrame.getByRole('button', { name: 'Advanced', exact: true }).click()
      const nextDialog = reopenedFrame.getByRole('dialog', { name: 'Advanced Filter', exact: true })
      await nextDialog.getByRole('button', { name: 'Column', exact: true }).click()
      await reopenedFrame.locator('[role=option][data-value="1"]').click()
      await nextDialog.getByRole('textbox', { name: 'Value', exact: true }).nth(0).fill('North')
      await nextDialog.getByRole('button', { name: 'OK', exact: true }).click()
      await reopened.bridge!.request('view.flush', {})
      if (await nextDialog.getByRole('alert').count())
        assert.fail(await nextDialog.getByRole('alert').innerText())
      await nextDialog.waitFor({ state: 'hidden' })
      const edited = structuredClone(reopened.engine.state)
      const filter = edited.sheets[0]!.features!.filterEdits as {
        hiddenRows: number[]
        filter: { columns: { colId: number; customs: { filters: unknown[] } }[] }
      }
      assert.deepEqual(filter.hiddenRows, [2, 3, 4, 5])
      assert.deepEqual(
        filter.filter.columns.map((column) => column.colId),
        [0, 1],
      )
      assert.equal(filter.filter.columns[0]!.customs.filters.length, 2)
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(reopened.engine.state, imported)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(reopened.engine.state, edited)
      await vscode.commands.executeCommand('workbench.action.files.save')
      const finalZip = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const finalXml = await finalZip.file('xl/worksheets/sheet1.xml')!.async('string')
      assert.match(finalXml, /<filterColumn colId="0">/)
      assert.match(finalXml, /<filterColumn colId="1">/)
      assert.match(finalXml, /operator="greaterThanOrEqual" val="10"/)
      assert.match(finalXml, /<customFilter val="North"/)
    } finally {
      if (api.getDocument(uri))
        await vscode.commands.executeCommand('workbench.action.files.save', uri)
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
      await browser.close()
    }
  })
})
