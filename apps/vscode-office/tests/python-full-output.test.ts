import { expect, it } from 'vitest'
import { execFile } from 'node:child_process'
import { promisify } from 'node:util'
import { copyFile, mkdir, writeFile } from 'node:fs/promises'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { join } from 'node:path'
import { salesWorkbook } from './fixtures/sales-workbook'
import { nativeSession } from './fixtures/native-session'
import { withTestWorkspace } from './helpers'
import { initialState } from '@genoffice/workbook-session/state'
import { readCurrentRange } from '@genoffice/workbook-session/read'
import { retainedHistoryStorage } from '@genoffice/workbook-session/history-records'
import { FileResourceStore } from '@genoffice/workbook-session/store'
import { WorkbookSnapshots } from '@genoffice/workbook-session/exchange/snapshots'

it('commits a full 100k typed result in 100 undoable batches, separately proves cancellation after batch three and saves/reopens all values', async () => {
  await withTestWorkspace(async (directory) => {
    const progressPath = fileURLToPath(
      new URL('../test-results/full-output-progress.json', import.meta.url),
    )
    await mkdir(join(progressPath, '..'), { recursive: true })
    const progress = (phase: string, data: object = {}) =>
      writeFile(
        progressPath,
        JSON.stringify({ phase, ...data, checkedAt: new Date().toISOString() }),
      )
    await progress('preparing-initial-snapshot')
    const path = join(directory, 'full-output.xlsx'),
      privateStorage = join(directory, 'private')
    await writeFile(path, await salesWorkbook())
    const runtime = await nativeSession(path, privateStorage),
      { engine, imports, snapshots } = runtime
    const startedAt = Date.now(),
      receipts: {
        cursor: number
        revision: number
        completedCount: number
        remainingCount: number
        durationMs: number
      }[] = []
    try {
      await runtime.waitForIndex()
      const sheetId = engine.state.sheets.find((sheet) => sheet.name === 'Data')!.id,
        sessionId = engine.descriptor.sessionId
      await engine.submit({
        schemaVersion: 1,
        sessionId,
        transactionId: 'manual-source',
        expectedRevision: 0,
        origin: 'manual',
        summary: 'Unsaved D2 100 to 120',
        operations: [{ op: 'set_cell', sheetId, address: 'D2', value: 120 }],
      })
      const manualRoot = engine.state.stateId
      const prepare = async () => {
        const exported = await snapshots.export({
          view: runtime.view(),
          workspaceRoot: directory,
          sheetId,
          range: 'A1:J10000',
        })
        const script = join(exported.directory, 'analysis.py')
        await copyFile(
          fileURLToPath(
            new URL('../skills/genoffice-workbook/scripts/analyze.py', import.meta.url),
          ),
          script,
        )
        await promisify(execFile)(
          'python3',
          [script, exported.manifestPath, '--mode', 'copy', '--new-sheet-name', 'FullCopy'],
          { cwd: exported.directory, timeout: 60_000 },
        )
        const plan = await imports.prepare({
          snapshotId: exported.manifest.snapshotId,
          bindingId: 'full-result-binding',
          resultManifestPath: join(exported.directory, 'result.manifest.json'),
          privateStorage,
          view: runtime.view(),
          expectedRevision: engine.descriptor.revision,
          currentDescriptor: () => engine.descriptor,
        })
        return { exported, plan }
      }
      const partial = await prepare()
      await progress('verifying-three-batch-cancellation')
      expect(partial.plan.batchCount).toBe(100)
      let revision = 1
      for (let cursor = 0; cursor < 3; cursor++) {
        const result = await imports.apply({
          planId: partial.plan.planId,
          bindingId: 'full-result-binding',
          cursor,
          expectedRevision: revision,
          transactionId: `partial-${cursor}`,
          engine,
        })
        revision = result.receipt.revision
      }
      const cancelled = new AbortController()
      cancelled.abort()
      await expect(
        imports.apply({
          planId: partial.plan.planId,
          bindingId: 'full-result-binding',
          cursor: 3,
          expectedRevision: revision,
          transactionId: 'cancel-after-third',
          engine,
          signal: cancelled.signal,
        }),
      ).rejects.toMatchObject({
        code: 'operation_cancelled',
        details: {
          progress: {
            status: 'cancelled',
            completedBatchCount: 3,
            completedCount: 3000,
            remainingCount: 97000,
          },
        },
      })
      expect(
        Object.keys(engine.state.sheets.find((sheet) => sheet.name === 'FullCopy')!.cells),
      ).toHaveLength(3000)
      for (let count = 0; count < 3; count++) await engine.undo()
      expect(engine.state.stateId).toBe(manualRoot)
      expect(engine.state.sheets.some((sheet) => sheet.name === 'FullCopy')).toBe(false)
      const full = await prepare()
      await progress('applying-full-result', { completedCount: 0 })
      await writeFile(
        join(full.exported.directory, 'result.ndjson'),
        'public result changed after immutable preparation',
      )
      for (let cursor = 0; cursor < full.plan.batchCount; cursor++) {
        const before = Date.now()
        const result = await imports.apply({
          planId: full.plan.planId,
          bindingId: 'full-result-binding',
          cursor,
          expectedRevision: engine.descriptor.revision,
          transactionId: `full-${cursor}`,
          engine,
        })
        receipts.push({
          cursor,
          revision: result.receipt.revision,
          completedCount: result.completedCount,
          remainingCount: result.remainingCount,
          durationMs: Date.now() - before,
        })
        expect(result.status).toBe(cursor === 99 ? 'completed' : 'applying')
        if ((cursor + 1) % 10 === 0) {
          await progress('applying-full-result', {
            completedCount: result.completedCount,
            remainingCount: result.remainingCount,
          })
          console.log(
            `Full output: ${cursor + 1}/100 batches, ${result.completedCount} committed cells.`,
          )
        }
      }
      const target = engine.state.sheets.find((sheet) => sheet.name === 'FullCopy')!
      expect(Object.keys(target.cells)).toHaveLength(100000)
      const current = await snapshots.export({
        view: runtime.view(),
        workspaceRoot: directory,
        sheetId: target.id,
        range: 'A1:J10000',
      })
      expect(current.manifest.data.sha256).toBe(full.exported.manifest.data.sha256)
      expect(engine.history).toHaveLength(101)
      const retained = await retainedHistoryStorage(
        await FileResourceStore.create(join(privateStorage, 'session')),
        engine.history,
      )
      expect(retained.bytes).toBeLessThanOrEqual(256 * 1024 * 1024)
      await progress('saving-and-reopening', { completedCount: 100000 })
      await runtime.save()
      expect(engine.descriptor.dirty).toBe(false)
      const reopened = await runtime.source.openSnapshot(path)
      try {
        const state = initialState(reopened.metadata),
          sheet = state.sheets.find((item) => item.name === 'FullCopy')!
        await runtime.waitForIndex(
          reopened.sessionId,
          state.sheets.flatMap((item) => (item.originalSheetId ? [item.originalSheetId] : [])),
        )
        const verified = await new WorkbookSnapshots().export({
          workspaceRoot: directory,
          sheetId: sheet.id,
          range: 'A1:J10000',
          view: {
            state,
            descriptor: {
              ...engine.descriptor,
              documentUri: pathToFileURL(path).href,
              sessionId: reopened.sessionId,
              revision: 0,
              stateId: state.stateId,
              savedStateId: state.stateId,
              dirty: false,
            },
            read: (logicalSheetId, range, signal) =>
              readCurrentRange({
                state,
                logicalSheetId,
                range,
                signal,
                source: runtime.source,
                sessionId: reopened.sessionId,
              }),
          },
        })
        expect(verified.manifest.data.sha256).toBe(full.exported.manifest.data.sha256)
      } finally {
        await runtime.source.close(reopened.sessionId)
      }
      const report = fileURLToPath(
        new URL('../test-results/python-full-output-proof.json', import.meta.url),
      )
      await mkdir(join(report, '..'), { recursive: true })
      await writeFile(
        report,
        JSON.stringify(
          {
            passed: true,
            platform: `${process.platform}-${process.arch}`,
            checkedAt: new Date().toISOString(),
            durationMs: Date.now() - startedAt,
            cells: 100000,
            batchCount: 100,
            retainedHistoryBytes: retained.bytes,
            sourceAndReopenedResultSha256: full.exported.manifest.data.sha256,
            cancelledAfterThree: {
              committedCount: 3000,
              remainingCount: 97000,
              undoRestoredManualRoot: true,
            },
            privateCopySurvivedPublicMutation: true,
            receipts,
            scope:
              'Actual Node session/import/save and Rust reopening; this test does not measure Webview latency.',
          },
          null,
          2,
        ),
      )
      await progress('passed', { completedCount: 100000 })
    } finally {
      await runtime.dispose()
    }
  })
}, 1_200_000)
