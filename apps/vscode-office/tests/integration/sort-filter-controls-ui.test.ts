import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import JSZip from 'jszip'
import { suite, test } from 'mocha'
import { chromium, expect, type Frame } from '@playwright/test'
import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'

suite('Desktop Sort Filter Controls UI', () => {
  test('uses the desktop group for smart toggle, reapply after an AI edit, criteria clearing and imported-filter removal', async function () {
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
      'sort-filter-controls-ui.xlsx',
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
            if (!(await page.title()).includes('sort-filter-controls-ui.xlsx')) continue
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
      await go.getByRole('textbox', { name: 'Reference', exact: true }).fill('Data!B2')
      await go.getByRole('button', { name: 'Go', exact: true }).click()
      await go.waitFor({ state: 'hidden' })
      await frame.locator('#ribbon-tab-data').click()
      const band = frame.locator('.workbook-shared-sort-filter')
      const click = async (name: string | RegExp) => {
        await band.getByRole('button', { name, exact: typeof name === 'string' }).click()
        await document.bridge!.request('view.flush', {})
      }
      await frame
        .page()
        .keyboard.press(`${process.platform === 'darwin' ? 'Meta' : 'Control'}+Shift+l`)
      await document.bridge!.request('view.flush', {})
      const range = { startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 }
      assert.deepEqual(
        (document.engine.state.sheets[0]!.features!.filterEdits as any).filter.range,
        range,
      )
      await click('Advanced')
      const dialog = frame.getByRole('dialog', { name: 'Advanced Filter', exact: true })
      await dialog.getByRole('button', { name: 'Column', exact: true }).click()
      await frame.locator('[role=option][data-value="1"]').click()
      await expect(dialog.getByRole('button', { name: 'Column', exact: true })).toHaveAttribute(
        'aria-expanded',
        'false',
      )
      await dialog.getByRole('button', { name: 'Condition 1', exact: true }).click()
      await frame.locator('[role=option][data-value="greaterThan"]').click()
      await dialog.getByRole('textbox', { name: 'Value', exact: true }).nth(0).fill('0')
      await dialog.getByRole('button', { name: 'OK', exact: true }).click()
      await document.bridge!.request('view.flush', {})
      if (await dialog.getByRole('alert').count())
        assert.fail(await dialog.getByRole('alert').innerText())
      await dialog.waitFor({ state: 'hidden' })
      assert.deepEqual(
        (document.engine.state.sheets[0]!.features!.filterEdits as any).hiddenRows,
        [2],
      )
      await document.engine.submit({
        schemaVersion: 1,
        sessionId: document.sessionId,
        transactionId: crypto.randomUUID(),
        expectedRevision: document.engine.descriptor.revision,
        origin: 'agent',
        summary: 'AI changes a previously filtered amount',
        operations: [
          {
            op: 'set_cell',
            sheetId: document.engine.state.sheets[0]!.id,
            address: 'B3',
            value: 30,
          },
        ],
      })
      await document.bridge!.request('view.flush', {})
      const beforeReapply = structuredClone(document.engine.state)
      const revision = document.engine.descriptor.revision
      await click('Reapply')
      assert.equal(document.engine.descriptor.revision, revision + 1)
      const applied = structuredClone(document.engine.state)
      assert.deepEqual((applied.sheets[0]!.features!.filterEdits as any).hiddenRows, [])
      assert.deepEqual(
        (applied.sheets[0]!.features!.filterEdits as any).filter,
        (beforeReapply.sheets[0]!.features!.filterEdits as any).filter,
      )
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, beforeReapply)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, applied)
      await vscode.commands.executeCommand('workbench.action.files.save')
      let saved = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      let xml = await saved.file('xl/worksheets/sheet1.xml')!.async('string')
      assert.match(xml, /operator="greaterThan" val="0"/)
      assert.doesNotMatch(xml, /<row\b[^>]*r="3"[^>]*hidden="1"/)
      const beforeClear = structuredClone(document.engine.state)
      await click('Clear')
      const cleared = structuredClone(document.engine.state)
      assert.deepEqual((cleared.sheets[0]!.features!.filterEdits as any).filter, {
        range,
        columns: [],
      })
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, beforeClear)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, cleared)
      await click(/^Filter\b/)
      const removed = structuredClone(document.engine.state)
      assert.equal((removed.sheets[0]!.features!.filterEdits as any).filter, null)
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, cleared)
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, removed)
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, cleared)
      await vscode.commands.executeCommand('workbench.action.files.save')
      saved = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      xml = await saved.file('xl/worksheets/sheet1.xml')!.async('string')
      assert.match(xml, /<autoFilter ref="A1:B4"/)
      assert.doesNotMatch(xml, /<customFilter/)
      assert.equal(
        await saved.file('customXml/manual-sentinel.xml')!.async('string'),
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
        const source = await snapshot.read(reopened.engine.state.sheets[0]!.id, 'A1:B4')
        const cells = new Map(source.cells.map((cell) => [cell.address, cell]))
        assert.equal(cells.get('A2')!.rawValue, 'North')
        assert.equal(cells.get('A2')!.style!.fillColor, '#CCEEFF')
        assert.equal(cells.get('B3')!.rawValue, 30)
        assert.equal(cells.get('B3')!.kind, 'number')
        const far = await snapshot.read(reopened.engine.state.sheets[0]!.id, 'A80')
        assert.equal(far.cells[0]!.rawValue, 'Separate far table')
      } finally {
        snapshot.release()
      }
      let importedFrame: Frame | undefined
      const reopenDeadline = Date.now() + 20000
      while (!importedFrame && Date.now() < reopenDeadline) {
        for (const context of browser.contexts())
          for (const page of context.pages()) {
            if (!(await page.title()).includes('sort-filter-controls-ui.xlsx')) continue
            for (const candidate of page.frames())
              if (
                await candidate
                  .locator('#ribbon-tab-data')
                  .count()
                  .catch(() => 0)
              )
                importedFrame = candidate
          }
        if (!importedFrame) await new Promise((resolve) => setTimeout(resolve, 100))
      }
      assert.ok(importedFrame)
      await importedFrame.locator('#ribbon-tab-data').click()
      await importedFrame
        .locator('.workbook-shared-sort-filter')
        .getByRole('button', { name: /^Filter\b/ })
        .click()
      await reopened.bridge!.request('view.flush', {})
      assert.equal((reopened.engine.state.sheets[0]!.features!.filterEdits as any).filter, null)
      await vscode.commands.executeCommand('workbench.action.files.save')
      saved = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      assert.doesNotMatch(
        await saved.file('xl/worksheets/sheet1.xml')!.async('string'),
        /<autoFilter/,
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
