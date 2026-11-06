import * as assert from 'node:assert/strict'
import { randomUUID } from 'node:crypto'
import * as vscode from 'vscode'
import JSZip from 'jszip'
import { suite, test } from 'mocha'
import { chromium, type Browser, type Frame } from '@playwright/test'
import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'
import { gotoReference } from './goto-helper'

async function workbookFrame(browser: Browser): Promise<Frame> {
  const started = Date.now()
  while (Date.now() - started < 20_000) {
    for (const context of browser.contexts())
      for (const page of context.pages()) {
        if (!(await page.title()).includes('insert-symbol-ui.xlsx')) continue
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
  throw new Error('Workbook frame did not become ready')
}
async function until(predicate: () => boolean | Promise<boolean>, label: string) {
  const started = Date.now()
  while (!(await predicate()) && Date.now() - started < 10_000)
    await new Promise((resolve) => setTimeout(resolve, 50))
  assert.ok(await predicate(), label)
}

suite('Desktop Insert Symbol UI', () => {
  test('uses current source values, facade coercion, exact receipts, native history and saved/reopened content', async function () {
    if (!process.env.GENOFFICE_TEST_CDP_PORT) return this.skip()
    this.timeout(120_000)
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as { getDocument(uri: vscode.Uri): WorkbookDocument | undefined }
    const uri = vscode.Uri.joinPath(
      vscode.workspace.workspaceFolders![0]!.uri,
      'insert-symbol-ui.xlsx',
    )
    const original = await applyCellEditsToXlsx(
      await manualWorkbook(),
      [
        { sheetName: 'Data', row: 0, column: 2, writeValue: true, cell: { value: true } },
        { sheetName: 'Data', row: 0, column: 3, writeValue: true, cell: { value: '=1+' } },
        {
          sheetName: 'Data',
          row: 0,
          column: 4,
          writeValue: true,
          cell: { value: 60 },
          style: { numberFormat: 'yyyy-mm-dd' },
        },
        {
          sheetName: 'Data',
          row: 4,
          column: 0,
          writeValue: true,
          cell: { value: 'Merged original' },
        },
        {
          sheetName: 'Other',
          row: 0,
          column: 2,
          writeValue: true,
          cell: { value: null, formula: '=Data!A2*3' },
        },
        {
          sheetName: 'Data',
          row: 79,
          column: 0,
          writeValue: true,
          cell: { value: 'Far original' },
          style: { bold: true, fillColor: '#DDEEFF' },
        },
      ],
      [],
      [],
      undefined,
      [],
      [],
      [],
      [],
      [],
      null,
      [],
      [],
      [
        // Inputs are unchanged for B1: its known result is 12.5 * 2 = 25.
        // C1 deliberately has no cache and exercises the pending-result guard.
        { sheetName: 'Other', cells: [{ row: 0, column: 1, value: 25 }] },
      ],
    )
    const seeded = await JSZip.loadAsync(original.buffer)
    const worksheet = await seeded.file('xl/worksheets/sheet1.xml')!.async('string')
    seeded.file(
      'xl/worksheets/sheet1.xml',
      worksheet.replace(
        '</sheetData>',
        '</sheetData><mergeCells count="1"><mergeCell ref="A5:B6"/></mergeCells>',
      ),
    )
    await vscode.workspace.fs.writeFile(uri, await seeded.generateAsync({ type: 'nodebuffer' }))
    const browser = await chromium.connectOverCDP(
      `http://127.0.0.1:${process.env.GENOFFICE_TEST_CDP_PORT}`,
    )
    let release = () => {},
      restoreSubmit = () => {}
    try {
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      const document = api.getDocument(uri)!
      await document.whenViewReady()
      let frame = await workbookFrame(browser)
      const initial = structuredClone(document.engine.state)
      const readySource = await document.snapshot().read(initial.sheets[1]!.id, 'B1:C1')
      assert.equal(readySource.cells.find((cell) => cell.address === 'B1')!.cachedValue, 25)
      assert.equal(
        readySource.cells.find((cell) => cell.address === 'B1')!.calculationState,
        'ready',
      )
      // Tool reads calculate the cache-less C1 on demand; the editor view still sees it pending,
      // which the dialog's pending-result guard below depends on.
      assert.equal(
        readySource.cells.find((cell) => cell.address === 'C1')!.calculationState,
        'ready',
      )
      const goto = async (reference: string) => {
        await gotoReference(frame, reference)
      }
      const open = async (category = 'Greek') => {
        await frame.getByRole('tab', { name: 'Insert', exact: true }).click()
        await frame.locator('.workbook-desktop-header [data-tip="Symbol"]').click()
        const dialog = frame.getByRole('dialog', { name: 'Symbol', exact: true })
        await dialog.getByRole('button', { name: category, exact: true }).click()
        return dialog
      }
      for (const [reference, category, symbol, address, expected, sheetIndex] of [
        ['Other!B1', 'Greek', 'Ω', 'B1', '25Ω', 1],
        ['Data!E1', 'Greek', 'Ω', 'E1', '60Ω', 0],
        ['Data!C1', 'Greek', 'Ω', 'C1', 'trueΩ', 0],
        ['Data!A1', 'Greek', 'Ω', 'A1', '00123Ω', 0],
        ['Data!B2:C3', 'Greek', 'Ω', 'B2', 'Ω', 0],
        ['Data!B6', 'Greek', 'Ω', 'A5', 'Merged originalΩ', 0],
        ['Data!D1', 'Math', 'π', 'D1', '=1+π', 0],
      ] as const) {
        await goto(reference)
        const dialog = await open(category)
        const revision = document.engine.descriptor.revision
        await dialog.getByRole('button', { name: symbol, exact: true }).click()
        await until(
          () => document.engine.descriptor.revision === revision + 1,
          `insert ${reference}`,
        )
        await until(
          () => dialog.getByRole('button', { name: 'Close', exact: true }).isEnabled(),
          'symbol commit acknowledgment',
        )
        const cell = document.engine.state.sheets[sheetIndex]!.cells[address]!
        if (reference === 'Data!D1') assert.equal(cell.formula, expected)
        else {
          assert.equal(cell.rawValue, expected)
          assert.equal(cell.formula, undefined)
        }
        if (reference === 'Data!B2:C3')
          assert.equal(document.engine.state.sheets[0]!.cells.C3, undefined)
        if (reference === 'Data!B6')
          assert.equal(document.engine.state.sheets[0]!.cells.B6, undefined)
        await dialog.getByRole('button', { name: 'Close', exact: true }).click()
        await vscode.commands.executeCommand('undo')
        assert.deepEqual(document.engine.state, initial)
      }
      await goto('Other!C1')
      const pendingDialog = await open()
      const pendingRevision = document.engine.descriptor.revision
      await pendingDialog.getByRole('button', { name: 'Ω', exact: true }).click()
      await frame.getByRole('alert').filter({ hasText: 'formula result is pending' }).waitFor()
      assert.equal(document.engine.descriptor.revision, pendingRevision)
      assert.deepEqual(document.engine.state, initial)
      await pendingDialog.getByRole('button', { name: 'Close', exact: true }).click()
      await goto('Data!A2')
      const numberDialog = await open('Currency & Common')
      const originalSubmit = document.engine.submit.bind(document.engine)
      let calls = 0
      const held = new Promise<void>((resolve) => {
        release = resolve
      })
      document.engine.submit = async (raw, signal) => {
        calls++
        if (calls === 1) await held
        return originalSubmit(raw, signal)
      }
      restoreSubmit = () => {
        document.engine.submit = originalSubmit
      }
      const beforeNumber = structuredClone(document.engine.state),
        numberRevision = document.engine.descriptor.revision
      const euro = numberDialog.getByRole('button', { name: '€', exact: true })
      await euro.click()
      await until(() => calls === 1, 'host receives one insertion')
      assert.equal(await euro.isDisabled(), true)
      assert.equal(
        await numberDialog.getByRole('button', { name: 'Close', exact: true }).isDisabled(),
        true,
      )
      let saveFinished = false,
        saveError: unknown
      const concurrentSave = vscode.commands.executeCommand('workbench.action.files.save').then(
        () => {
          saveFinished = true
        },
        (error) => {
          saveFinished = true
          saveError = error
        },
      )
      await euro.dispatchEvent('click')
      await frame.page().keyboard.press('Escape')
      assert.equal(await numberDialog.isVisible(), true)
      assert.equal(document.engine.descriptor.revision, numberRevision)
      assert.deepEqual(document.engine.state, beforeNumber)
      assert.equal(saveFinished, false, 'native Save waits for the pending symbol edit')
      release()
      await until(() => euro.isEnabled(), 'unlocks after accepted receipt')
      await concurrentSave
      assert.equal(saveError, undefined)
      assert.equal(document.engine.descriptor.dirty, false)
      const currencySave = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      assert.match(await currencySave.file('xl/styles.xml')!.async('string'), /#,##0\.00€/)
      assert.equal(calls, 1)
      assert.equal(document.engine.state.sheets[0]!.cells.A2!.rawValue, 12.5)
      assert.equal(document.engine.state.sheets[0]!.cells.A2!.numberFormat, '#,##0.00€')
      await numberDialog.getByRole('button', { name: 'Greek', exact: true }).click()
      await numberDialog.getByRole('button', { name: 'Ω', exact: true }).click()
      await until(
        () => document.engine.descriptor.revision === numberRevision + 2,
        'repeated insertion advances from accepted receipt',
      )
      await until(
        () => numberDialog.getByRole('button', { name: 'Close', exact: true }).isEnabled(),
        'second insertion complete',
      )
      restoreSubmit()
      assert.equal(document.engine.state.sheets[0]!.cells.A2!.rawValue, '12.5Ω')
      assert.equal(document.engine.state.sheets[0]!.cells.A2!.style!.bold, true)
      assert.equal(document.engine.state.sheets[0]!.cells.A2!.style!.fillColor, '#CCEEFF')
      await numberDialog.getByRole('button', { name: 'Close', exact: true }).click()
      const afterNumber = structuredClone(document.engine.state)
      await vscode.commands.executeCommand('undo')
      await vscode.commands.executeCommand('undo')
      assert.deepEqual(document.engine.state, beforeNumber)
      await vscode.commands.executeCommand('redo')
      await vscode.commands.executeCommand('redo')
      assert.deepEqual(document.engine.state, afterNumber)

      await goto('Data!A80')
      const farDialog = await open('Currency & Common')
      const acknowledge = new Promise<void>((resolve) => {
        release = resolve
      })
      let receiptHeld = false
      document.engine.submit = async (raw, signal) => {
        const receipt = await originalSubmit(raw, signal)
        receiptHeld = true
        await acknowledge
        return receipt
      }
      const farRevision = document.engine.descriptor.revision
      await farDialog.getByRole('button', { name: '™', exact: true }).click()
      await until(() => receiptHeld, 'authority commits before delayed acknowledgment')
      assert.equal(document.engine.state.sheets[0]!.cells.A80!.rawValue, 'Far original™')
      await originalSubmit({
        schemaVersion: 1,
        sessionId: document.sessionId,
        transactionId: randomUUID(),
        expectedRevision: document.engine.descriptor.revision,
        origin: 'manual',
        summary: 'Concurrent edit before symbol acknowledgment',
        operations: [
          { op: 'set_note', sheetId: initial.sheets[0]!.id, address: 'A1', text: 'concurrent' },
        ],
      })
      assert.equal(document.engine.descriptor.revision, farRevision + 2)
      release()
      await until(
        () => farDialog.getByRole('button', { name: 'Close', exact: true }).isEnabled(),
        'delayed receipt delivered',
      )
      restoreSubmit()
      await farDialog.getByRole('button', { name: '€', exact: true }).click()
      await frame.getByRole('alert').filter({ hasText: 'document changed' }).waitFor()
      assert.equal(document.engine.descriptor.revision, farRevision + 2)
      assert.equal(document.engine.state.sheets[0]!.cells.A80!.rawValue, 'Far original™')
      await farDialog.getByRole('button', { name: 'Close', exact: true }).click()
      await vscode.commands.executeCommand('workbench.action.files.save')
      const saved = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      const xml = await saved.file('xl/worksheets/sheet1.xml')!.async('string')
      assert.match(xml, /12\.5Ω/)
      assert.match(xml, /Far original™/)
      assert.equal(
        await saved.file('customXml/manual-sentinel.xml')!.async('string'),
        '<preserve id="manual-fixture">unchanged</preserve>',
      )
      const previousSession = document.sessionId
      await vscode.commands.executeCommand('workbench.action.closeActiveEditor')
      await until(() => !api.getDocument(uri), 'old document disposed')
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      const reopened = api.getDocument(uri)!
      await reopened.whenViewReady()
      assert.notEqual(reopened.sessionId, previousSession)
      frame = await workbookFrame(browser)
      await goto('Data!A2')
      let presentation = (await reopened.bridge!.request('view.presentation', {
        stateId: reopened.engine.state.stateId,
        range: 'A2',
      })) as { values: unknown[][]; styles: Record<string, { bl?: number; bg?: { rgb?: string } }> }
      assert.deepEqual(presentation.values, [['12.5Ω']])
      assert.equal(presentation.styles.A2!.bl, 1)
      assert.equal(presentation.styles.A2!.bg!.rgb, '#CCEEFF')
      await goto('Data!A80')
      presentation = (await reopened.bridge!.request('view.presentation', {
        stateId: reopened.engine.state.stateId,
        range: 'A80',
      })) as typeof presentation
      assert.deepEqual(presentation.values, [['Far original™']])
      assert.equal(reopened.engine.descriptor.revision, 0)
      assert.equal(reopened.engine.descriptor.dirty, false)
    } finally {
      release()
      restoreSubmit()
      if (api.getDocument(uri))
        await vscode.commands.executeCommand('workbench.action.files.save', uri)
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri)
      await browser.close()
    }
  })
})
