import { expect, it } from 'vitest'
import { writeFile, readFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import { join } from 'node:path'
import { manualWorkbook } from './fixtures/manual-workbook'
import { withTestWorkspace } from './helpers'
import { XlsxSidecarClient } from '@genoffice/workbook-session/sidecar/client'
import { SidecarWorkbookSource } from '@genoffice/workbook-session/sidecar/source'
import { openWorkbookSnapshot } from '@genoffice/workbook-session/open'
import { initialState } from '@genoffice/workbook-session/state'
import { FileResourceStore } from '@genoffice/workbook-session/store'
import { WorkbookTransactions } from '@genoffice/workbook-session/transactions'
import { prepareCandidateMaterial } from '@genoffice/workbook-session/candidate-material'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { SessionPackageWriter } from '@genoffice/workbook-session/package-writer'
import { hashFile } from '@genoffice/workbook-session/resources'

async function historyFixture(
  run: (context: {
    engine: WorkbookTransactions
    sheetId: string
    otherId: string
    edit(operations: WorkbookOperation[], origin?: 'manual' | 'agent'): Promise<unknown>
    save(): Promise<void>
    target: string
    original: Buffer
    counts(): { edits: number; moves: number }
  }) => Promise<void>,
) {
  await withTestWorkspace(async (directory) => {
    const client = new XlsxSidecarClient(
        fileURLToPath(
          new URL('../../sheets/native/xlsx-engine/target/release/xlsx-sidecar', import.meta.url),
        ),
      ),
      source = new SidecarWorkbookSource(client),
      original = await manualWorkbook(),
      path = join(directory, 'history.xlsx'),
      target = join(directory, 'saved.xlsx')
    await writeFile(path, original)
    const opened = await openWorkbookSnapshot({
      uri: new URL(`file://${path}`).href,
      storageDirectory: directory,
      source,
    })
    let engine: WorkbookTransactions | undefined
    try {
      const state = initialState(opened.metadata),
        sheetId = state.sheets[0]!.id,
        otherId = state.sheets[1]!.id
      let edits = 0,
        moves = 0,
        sequence = 0
      engine = new WorkbookTransactions({
        sessionId: 'history-session',
        documentUri: new URL(`file://${path}`).href,
        state,
        store: await FileResourceStore.create(join(directory, 'store')),
        interpreter: {
          async prepare(before, intent, signal) {
            return computeCandidateState(
              before,
              intent,
              await prepareCandidateMaterial({
                state: before,
                operations: intent.operations,
                source,
                sourceSessionId: opened.original.sidecarSessionId,
                signal,
              }),
              { signal },
            )
          },
          dispose() {},
        },
        host: {
          committed() {
            edits++
          },
          historyChanged() {
            moves++
          },
          projectionChanged() {},
          conflict() {},
        },
      })
      const session = engine
      await run({
        engine,
        sheetId,
        otherId,
        target,
        original,
        counts: () => ({ edits, moves }),
        edit(operations, origin = 'manual') {
          return session.submit({
            schemaVersion: 1,
            sessionId: 'history-session',
            transactionId: `edit-${++sequence}`,
            expectedRevision: session.descriptor.revision,
            origin,
            summary: `Edit ${sequence}`,
            operations,
          })
        },
        async save() {
          await session.saveWith(async (current) => {
            const writer = new SessionPackageWriter(client, opened.metadata),
              expected = await hashFile(target).catch((error: NodeJS.ErrnoException) => {
                if (error.code === 'ENOENT') return null
                throw error
              }),
              staged = await writer.stage(opened.original.path, current, target)
            await writer.publish(staged.stagedPath, target, expected)
          })
        },
      })
    } finally {
      await engine?.dispose()
      await opened.dispose()
      client.stop()
    }
  })
}

it('has one entry per mixed-origin batch, retains history across save and truncates only the redo branch on a new edit', async () => {
  await historyFixture(async ({ engine, sheetId, otherId, edit, save, counts }) => {
    await edit([{ op: 'set_cell', sheetId, address: 'A2', value: 20 }])
    await edit(
      [
        { op: 'set_cell', sheetId, address: 'A2', value: 30 },
        { op: 'set_cell', sheetId: otherId, address: 'A1', value: 9 },
      ],
      'agent',
    )
    await edit([
      { op: 'format_range', sheetId, range: 'A2', format: { bold: false, fontColor: '#123456' } },
    ])
    const entries = engine.history,
      savedId = engine.state.stateId
    await save()
    expect(engine.history).toEqual(entries)
    expect(engine.descriptor).toMatchObject({ revision: 3, dirty: false, savedStateId: savedId })
    await engine.undo(entries[2]!.entryId)
    expect(engine.state.sheets[0]!.cells.A2!.style?.bold).toBe(true)
    await engine.undo(entries[1]!.entryId)
    expect(engine.state.sheets[0]!.cells.A2!.rawValue).toBe(20)
    expect(engine.state.sheets[1]!.cells.A1).toBeUndefined()
    await engine.redo(entries[1]!.entryId)
    await engine.redo(entries[2]!.entryId)
    expect(engine.descriptor).toMatchObject({ stateId: savedId, revision: 7, dirty: false })
    await engine.undo()
    await edit([{ op: 'set_cell', sheetId, address: 'A1', value: 'new branch' }])
    expect(engine.history).toHaveLength(3)
    const current = engine.state
    await expect(engine.redo()).rejects.toMatchObject({ code: 'revision_conflict' })
    expect(engine.state).toBe(current)
    expect(counts()).toEqual({ edits: 4, moves: 5 })
  })
})

it('undoes structural and chart state across saves, then serializes the original identity exactly', async () => {
  await historyFixture(async ({ engine, sheetId, edit, save, target, original }) => {
    await edit([
      { op: 'add_chart', sheetId, dataRange: 'A1:A2', chartType: 'column', title: 'History chart' },
    ])
    const chartStateId = engine.state.stateId
    await edit([{ op: 'insert_rows', sheetId, row: 2, count: 1 }])
    await save()
    const retained = engine.history
    await engine.undo()
    expect(engine.state.stateId).toBe(chartStateId)
    expect(engine.state.structuralLog).toHaveLength(0)
    await save()
    await engine.undo()
    await save()
    expect(await readFile(target)).toEqual(original)
    expect(engine.history).toEqual(retained)
    await engine.redo()
    await engine.redo()
    expect(engine.state.structuralLog).toHaveLength(1)
    expect(engine.descriptor.dirty).toBe(true)
    await save()
  })
})
