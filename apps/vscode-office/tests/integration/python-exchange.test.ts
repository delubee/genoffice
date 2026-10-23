import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import { suite, test } from 'mocha'
import { execFile } from 'node:child_process'
import { promisify } from 'node:util'
import { copyFile, mkdir, readFile, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { createHash } from 'node:crypto'
import JSZip from 'jszip'
import { salesWorkbook, SALES_EXPECTED } from '../fixtures/sales-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'

suite('Product Python workbook exchange', () => {
  test('exports 100k current cells, runs visible standard-library analysis and imports an independently verified summary/chart', async function () {
    this.timeout(180_000)
    const extension = vscode.extensions.getExtension('genoffice.genoffice-workbook')!,
      api = (await extension.activate()) as {
        getDocument(uri: vscode.Uri): WorkbookDocument | undefined
        revertForTests(uri: vscode.Uri): Promise<void>
      }
    const uri = vscode.Uri.joinPath(vscode.workspace.workspaceFolders![0]!.uri, 'python-sales.xlsx')
    const invoke = async (name: string, input: object) => {
      const result = await vscode.lm.invokeTool(`genoffice_workbook_${name}`, {
        input,
        toolInvocationToken: undefined,
      })
      const content = result.content.find((part) => part instanceof vscode.LanguageModelTextPart)
      assert.ok(content instanceof vscode.LanguageModelTextPart)
      const parsed = JSON.parse(content.value)
      assert.equal(parsed.status, 'ok', JSON.stringify(parsed))
      return parsed
    }
    let document: WorkbookDocument | undefined
    const startedAt = Date.now()
    try {
      const original = await salesWorkbook(),
        fixtureHash = createHash('sha256').update(original).digest('hex')
      await vscode.workspace.fs.writeFile(uri, original)
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      document = api.getDocument(uri)!
      await document.whenViewReady()
      await document.bridge!.request('view.flush', {})
      const dataId = document.engine.state.sheets.find((sheet) => sheet.name === 'Data')!.id
      await document.engine.submit({
        schemaVersion: 1,
        sessionId: document.sessionId,
        transactionId: 'python-manual-d2',
        expectedRevision: 0,
        origin: 'manual',
        summary: 'Unsaved Data D2 100 to 120',
        operations: [{ op: 'set_cell', sheetId: dataId, address: 'D2', value: 120 }],
      })
      const context = await invoke('context', {
        schemaVersion: 1,
        documentUri: uri.toString(),
        sheetId: dataId,
        range: 'A1:J10000',
      })
      const bindingId = context.data.binding.bindingId
      const exported = await invoke('export', {
        schemaVersion: 1,
        bindingId,
        expectedRevision: 1,
        sheetId: dataId,
        range: 'A1:J10000',
      })
      assert.equal(exported.data.shape.cellCount, 100000)
      assert.deepEqual(
        await vscode.workspace.fs.readFile(uri),
        original,
        'export does not save unsaved input',
      )
      console.log('Python exchange: 100,000 current cells exported without saving the XLSX.')
      const script = join(exported.data.taskDirectory, 'analysis.py')
      await copyFile(
        join(extension.extensionPath, 'skills/genoffice-workbook/scripts/analyze.py'),
        script,
      )
      await promisify(execFile)('python3', [script, exported.data.manifestPath], {
        cwd: exported.data.taskDirectory,
        timeout: 60_000,
      })
      const prepared = await invoke('prepareImport', {
        schemaVersion: 1,
        bindingId,
        expectedRevision: 1,
        snapshotId: exported.data.snapshotId,
        resultManifestUri: vscode.Uri.file(
          join(exported.data.taskDirectory, 'result.manifest.json'),
        ).toString(),
      })
      assert.equal(document.engine.descriptor.revision, 1, 'prepare creates no worksheet or cells')
      const imported = await invoke('applyImport', {
        schemaVersion: 1,
        bindingId,
        expectedRevision: 1,
        planId: prepared.data.planId,
        cursor: 0,
        transactionId: 'python-summary-batch',
      })
      assert.equal(imported.data.status, 'completed')
      const summaryId = imported.data.target.sheetId as string
      const summaryContext = await invoke('context', {
        schemaVersion: 1,
        documentUri: uri.toString(),
        sheetId: summaryId,
        range: 'A1:B5',
      })
      const summaryBinding = summaryContext.data.binding.bindingId
      const summary = await invoke('read', {
        schemaVersion: 1,
        bindingId: summaryBinding,
        expectedRevision: 2,
        sheetId: summaryId,
        range: 'A1:B5',
      })
      const actualCents: Record<string, number> = {}
      for (let row = 1; row < 5; row++)
        actualCents[summary.data.cells[row * 2].rawValue] = Math.round(
          summary.data.cells[row * 2 + 1].rawValue * 100,
        )
      assert.deepEqual({ ...actualCents }, SALES_EXPECTED.editedCents)
      assert.equal(actualCents.North! - SALES_EXPECTED.originalCents.North, 2000)
      for (const region of ['South', 'East', 'West'] as const)
        assert.equal(actualCents[region], SALES_EXPECTED.originalCents[region])
      assert.deepEqual(
        Object.keys(document.engine.state.sheets[0]!.cells),
        ['D2'],
        'source materialization must not become extra edits',
      )
      const chart = await invoke('apply', {
        schemaVersion: 1,
        bindingId: summaryBinding,
        expectedRevision: 2,
        transactionId: 'python-summary-chart',
        summary: 'Create regional sales column chart',
        operations: [
          {
            op: 'add_chart',
            sheetId: summaryId,
            chartType: 'column',
            dataRange: 'A1:B5',
            title: 'Regional sales',
          },
        ],
      })
      assert.equal(chart.data.revision, 3)
      assert.equal(
        (
          document.engine.state.sheets.find((sheet) => sheet.id === summaryId)!.features!
            .charts as unknown[]
        ).length,
        1,
      )
      await document.engine.submit({
        schemaVersion: 1,
        sessionId: document.sessionId,
        transactionId: 'manual-after-python',
        expectedRevision: 3,
        origin: 'manual',
        summary: 'Continue manual editing',
        operations: [
          {
            op: 'set_cell',
            sheetId: summaryId,
            address: 'B2',
            value: actualCents.North! / 100 + 1,
          },
        ],
      })
      await document.requestHistory('undo')
      await document.requestHistory('redo')
      await vscode.commands.executeCommand('workbench.action.files.save')
      await document.requestHistory('undo')
      await vscode.commands.executeCommand('workbench.action.files.save')
      assert.equal(document.engine.history.length, 4)
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await document.whenDisposed()
      const saved = await readFile(uri.fsPath),
        zip = await JSZip.loadAsync(saved)
      assert.equal(
        await zip.file('customXml/genoffice-sentinel.xml')!.async('string'),
        '<preserve fixture="genoffice-v1">unknown part must survive</preserve>',
      )
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      document = api.getDocument(uri)!
      await document.whenViewReady()
      await document.bridge!.request('view.flush', {})
      const reopenedSummary = document.engine.state.sheets.find(
          (sheet) => sheet.name === 'Summary',
        )!,
        reopenedData = document.engine.state.sheets.find((sheet) => sheet.name === 'Data')!
      const reopenedContext = await invoke('context', {
        schemaVersion: 1,
        documentUri: uri.toString(),
        sheetId: reopenedSummary.id,
      })
      const verified = await invoke('read', {
        schemaVersion: 1,
        bindingId: reopenedContext.data.binding.bindingId,
        expectedRevision: 0,
        sheetId: reopenedSummary.id,
        range: 'A1:B5',
      })
      for (let row = 1; row < 5; row++)
        assert.equal(
          Math.round(verified.data.cells[row * 2 + 1].rawValue * 100),
          actualCents[verified.data.cells[row * 2].rawValue],
        )
      const source = await invoke('read', {
        schemaVersion: 1,
        bindingId: reopenedContext.data.binding.bindingId,
        sheetId: reopenedData.id,
        range: 'D2',
      })
      assert.equal(source.data.cells[0].rawValue, 120)
      const sourceAfterReopen = await invoke('export', {
        schemaVersion: 1,
        bindingId: reopenedContext.data.binding.bindingId,
        expectedRevision: 0,
        sheetId: reopenedData.id,
        range: 'A1:J10000',
      })
      assert.equal(
        sourceAfterReopen.data.dataSha256,
        exported.data.dataSha256,
        'all 100,000 source cells and their declared types/formats must survive analysis and save unchanged',
      )
      assert.ok(
        (document.metadata as { visuals: unknown[] }).visuals.length >= 1,
        'the saved chart must reopen as an original workbook chart',
      )
      await mkdir(join(extension.extensionPath, 'test-results'), { recursive: true })
      await writeFile(
        join(extension.extensionPath, 'test-results/python-native-proof.json'),
        JSON.stringify(
          {
            passed: true,
            version: vscode.version,
            platform: `${process.platform}-${process.arch}`,
            fixtureHash,
            inputCells: 100000,
            durationMs: Date.now() - startedAt,
            sourceDataDigestPreserved: true,
            sourceDataSha256: exported.data.dataSha256,
            sourceRevision: 1,
            actualCents,
            independentExpectedCents: SALES_EXPECTED.editedCents,
            chartReopened: true,
            originalUnknownPartPreserved: true,
            scriptPath: script,
            manifestPath: exported.data.manifestPath,
            checkedAt: new Date().toISOString(),
            execution:
              'Existing Python interpreter invoked by test harness; native agent terminal approval not exercised.',
          },
          null,
          2,
        ),
      )
    } finally {
      if (document && !document.isClosed && document.engine.descriptor.dirty) {
        await api.revertForTests(uri)
        await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
        await document.whenViewReady()
        await vscode.commands.executeCommand('workbench.action.files.save')
      }
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri, { useTrash: false }).then(undefined, () => {})
    }
  })
})
