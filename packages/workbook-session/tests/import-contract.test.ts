import { expect, it } from 'vitest'
import { resultManifestSchema, shapeSchema } from '../src/exchange/schemas'
import { writeFile, readFile, symlink, stat } from 'node:fs/promises'
import { join } from 'node:path'
import { WorkbookSnapshots } from '../src/exchange/snapshots'
import { prepareResultCopy } from '../src/exchange/prepare-import'
import { writeTypedNdjson, readTypedNdjson, readPrivateBatch } from '../src/exchange/ndjson'
import { initialState } from '../src/state'
import { parseRange, formatAddress } from '@genoffice/xlsx-gateway/domain/cell-address'
import { useSessionTestResources } from './helpers'
import { importFixture } from './import-fixture'
import { workspaceTaskDirectory, openWorkspaceFile } from '../src/exchange/paths'
const resources = useSessionTestResources()

it('rejects redirected workspace/task directories before creating or reading files outside the registered root', async () => {
  const workspace = await resources.directory(),
    outside = await resources.directory()
  await symlink(
    outside,
    join(workspace, '.genoffice'),
    process.platform === 'win32' ? 'junction' : 'dir',
  )
  await expect(workspaceTaskDirectory(workspace, 'must-not-create')).rejects.toMatchObject({
    code: 'path_not_allowed',
  })
  await expect(stat(join(outside, 'work'))).rejects.toMatchObject({ code: 'ENOENT' })
  await expect(
    openWorkspaceFile(
      join(workspace, '.genoffice', 'secret.json'),
      join(workspace, '.genoffice'),
      1024,
    ),
  ).rejects.toMatchObject({ code: 'path_not_allowed' })
})

it('commits a new sheet and its first batch atomically, uses exact cursors and preserves one undo entry per batch', async () => {
  const fixture = await importFixture(await resources.directory())
  const { engine, imports } = fixture
  const originalId = engine.state.stateId
  try {
    const plan = await fixture.prepare()
    expect(engine.state.stateId).toBe(originalId)
    expect(plan).toMatchObject({ status: 'prepared', batchCount: 3, completedCount: 0 })
    const input = {
      planId: plan.planId,
      bindingId: 'task-binding',
      cursor: 0,
      expectedRevision: 0,
      transactionId: 'batch-0',
      engine,
    }
    await expect(imports.apply({ ...input, cursor: 1 })).rejects.toMatchObject({
      code: 'cursor_invalid',
    })
    const first = await imports.apply(input)
    expect(first).toMatchObject({
      nextCursor: 1,
      completedCount: 1000,
      remainingCount: 1502,
      receipt: { revision: 1 },
    })
    expect(engine.state.sheets).toHaveLength(2)
    expect(engine.state.sheets[1]!.cells.B2!.rawValue).toBe(0)
    expect(engine.history).toHaveLength(1)
    expect((await imports.apply(input)).receipt).toEqual(first.receipt)
    expect(engine.descriptor.revision).toBe(1)
    await expect(imports.apply({ ...input, transactionId: 'replay-token' })).rejects.toMatchObject({
      code: 'cursor_invalid',
    })
    const second = await imports.apply({
      ...input,
      cursor: 1,
      expectedRevision: 1,
      transactionId: 'batch-1',
    })
    const last = await imports.apply({
      ...input,
      cursor: 2,
      expectedRevision: second.receipt.revision,
      transactionId: 'batch-2',
    })
    expect(last).toMatchObject({ status: 'completed', completedCount: 2502, remainingCount: 0 })
    expect(engine.state.sheets[1]!.cells.G418!.rawValue).toBe(2501)
    expect(engine.history).toHaveLength(3)
    const finalId = engine.state.stateId
    for (let index = 0; index < 3; index++) await engine.undo()
    expect(engine.state.stateId).toBe(originalId)
    expect(engine.state.sheets).toHaveLength(1)
    for (let index = 0; index < 3; index++) await engine.redo()
    expect(engine.state.stateId).toBe(finalId)
    expect(engine.descriptor.revision).toBe(9)
  } finally {
    await fixture.dispose()
  }
})

it('does not let a refreshed revision bypass manual edits between import batches and expires prepared plans after 30 minutes', async () => {
  const fixture = await importFixture(await resources.directory())
  const { engine, imports } = fixture
  try {
    const plan = await fixture.prepare(),
      input = {
        planId: plan.planId,
        bindingId: 'task-binding',
        cursor: 0,
        expectedRevision: 0,
        transactionId: 'first',
        engine,
      }
    await imports.apply(input)
    await engine.submit({
      schemaVersion: 1,
      sessionId: 'import-session',
      transactionId: 'manual-between-batches',
      expectedRevision: 1,
      origin: 'manual',
      summary: 'Manual interruption',
      operations: [{ op: 'set_cell', sheetId: fixture.sheetId, address: 'A1', value: 121 }],
    })
    const current = engine.state
    await expect(
      imports.apply({
        ...input,
        cursor: 1,
        expectedRevision: 2,
        transactionId: 'old-result-new-revision',
      }),
    ).rejects.toMatchObject({
      code: 'snapshot_stale',
      details: { progress: { completedCount: 1000, status: 'stale' } },
    })
    expect(engine.state).toBe(current)
    expect(engine.state.sheets[1]!.cells.B2!.rawValue).toBe(0)
  } finally {
    await fixture.dispose()
  }
  resources.freezeTime()
  const expiring = await importFixture(await resources.directory())
  try {
    const plan = await expiring.prepare(),
      original = expiring.engine.state
    resources.freezeTime('2026-10-21T00:30:01.000Z')
    await expect(
      expiring.imports.apply({
        planId: plan.planId,
        bindingId: 'task-binding',
        cursor: 0,
        expectedRevision: 0,
        transactionId: 'expired',
        engine: expiring.engine,
      }),
    ).rejects.toMatchObject({ code: 'plan_expired' })
    expect(expiring.engine.state).toBe(original)
  } finally {
    await expiring.dispose()
  }
})

it('leaves no empty sheet after a failed first commit and reports cancellation after a commit as partial committed progress', async () => {
  const failed = await importFixture(await resources.directory(), {
    beforeCommit() {
      throw new Error('Injected disk failure before root publication')
    },
  })
  try {
    const plan = await failed.prepare(),
      original = failed.engine.state
    await expect(
      failed.imports.apply({
        planId: plan.planId,
        bindingId: 'task-binding',
        cursor: 0,
        expectedRevision: 0,
        transactionId: 'failed-root',
        engine: failed.engine,
      }),
    ).rejects.toMatchObject({ details: { progress: { completedCount: 0, status: 'failed' } } })
    expect(failed.engine.state).toBe(original)
    expect(failed.engine.state.sheets).toHaveLength(1)
    expect(failed.engine.history).toHaveLength(0)
  } finally {
    await failed.dispose()
  }
  const controller = new AbortController(),
    cancelled = await importFixture(await resources.directory(), {
      committed() {
        controller.abort()
      },
    })
  try {
    const plan = await cancelled.prepare()
    const result = await cancelled.imports.apply({
      planId: plan.planId,
      bindingId: 'task-binding',
      cursor: 0,
      expectedRevision: 0,
      transactionId: 'cancel-after-commit',
      engine: cancelled.engine,
      signal: controller.signal,
    })
    expect(result).toMatchObject({
      status: 'cancelled',
      completedCount: 1000,
      remainingCount: 1502,
      receipt: { status: 'committed', revision: 1 },
    })
    expect(cancelled.engine.history).toHaveLength(1)
    expect(cancelled.engine.state.sheets[1]!.cells.B2!.rawValue).toBe(0)
  } finally {
    await cancelled.dispose()
  }
})

it('requires exact nonempty rectangles, an explicit source, a single target and a complete finite typed result', () => {
  expect(shapeSchema.safeParse({ rows: 0, columns: 2, cellCount: 0 }).success).toBe(false)
  expect(shapeSchema.safeParse({ rows: 2, columns: 2, cellCount: 3 }).success).toBe(false)
  const result = {
    schemaVersion: 1,
    kind: 'workbook-result',
    source: {
      snapshotId: 'snapshot',
      sourceManifestSha256: 'a'.repeat(64),
      sourceDataSha256: 'b'.repeat(64),
      sourceRevision: 1,
    },
    output: {
      fileName: 'result.ndjson',
      sha256: 'c'.repeat(64),
      bytes: 128,
      recordCount: 1,
      encoding: 'utf-8',
      format: 'typed-ndjson-v1',
    },
    shape: { rows: 1, columns: 2, cellCount: 2 },
    target: { newSheetName: 'Summary', startCell: 'A1' },
    complete: true,
  }
  expect(resultManifestSchema.parse(result).mode).toBe('values')
  expect(resultManifestSchema.safeParse({ ...result, complete: false }).success).toBe(false)
  expect(
    resultManifestSchema.safeParse({ ...result, target: { ...result.target, sheetId: 'other' } })
      .success,
  ).toBe(false)
  expect(
    resultManifestSchema.safeParse({
      ...result,
      output: { ...result.output, fileName: '../private.ndjson' },
    }).success,
  ).toBe(false)
  expect(resultManifestSchema.safeParse({ ...result, schemaVersion: 2 }).success).toBe(false)
})

it('exports current unsaved state, keeps source authority private and stages immutable validated output before any document changes', async () => {
  const workspace = await resources.directory(),
    privateStorage = await resources.directory()
  const state = structuredClone(
    initialState({ sheets: [{ id: 'base', name: 'Data', rowCount: 2, columnCount: 2 }] }),
  )
  state.sheets[0]!.cells.A1 = { kind: 'number', rawValue: 120 }
  const sheetId = state.sheets[0]!.id
  const descriptor = {
    documentUri: 'file:///workspace/source.xlsx',
    sessionId: 'source-session',
    revision: 7,
    stateId: state.stateId,
    savedStateId: 'before-manual-edit',
    dirty: true,
    lifecycle: 'ready' as const,
  }
  const snapshots = new WorkbookSnapshots()
  const source = await snapshots.export({
    workspaceRoot: workspace,
    sheetId,
    range: 'A1:B2',
    view: {
      state,
      descriptor,
      async read(id, range) {
        const bounds = parseRange(range),
          cells = []
        for (let row = bounds.startRow; row <= bounds.endRow; row++)
          for (let column = bounds.startColumn; column <= bounds.endColumn; column++) {
            const address = formatAddress(row, column),
              cell = state.sheets[0]!.cells[address] ?? { kind: 'blank' as const, rawValue: null }
            cells.push({ address, ...cell, displayValue: String(cell.rawValue ?? '') })
          }
        return { stateId: state.stateId, logicalSheetId: id, range, cells }
      },
    },
  })
  const input = []
  for await (const record of readTypedNdjson({
    path: source.dataPath,
    shape: source.manifest.shape,
    expected: source.manifest.data,
  }))
    input.push(...record.cells)
  expect(input[0]).toEqual({ kind: 'number', value: 120 })
  const shape = { rows: 1, columns: 2, cellCount: 2 },
    output = await writeTypedNdjson({
      path: join(source.directory, 'result.ndjson'),
      shape,
      records: [
        {
          rowOffset: 0,
          columnOffset: 0,
          cells: [
            { kind: 'number', value: 42 },
            { kind: 'string', value: '=literal' },
          ],
        },
      ],
    })
  const result = {
    schemaVersion: 1,
    kind: 'workbook-result',
    source: {
      snapshotId: source.manifest.snapshotId,
      sourceRevision: 7,
      sourceManifestSha256: source.manifestSha256,
      sourceDataSha256: source.manifest.data.sha256,
    },
    shape,
    output,
    target: { sheetId, startCell: 'A1' },
    mode: 'values',
    complete: true,
  }
  const resultManifestPath = join(source.directory, 'result.manifest.json')
  await writeFile(resultManifestPath, JSON.stringify(result))
  const prepared = await prepareResultCopy({
    source: snapshots.get(source.manifest.snapshotId),
    resultManifestPath,
    privateStorage,
    state,
    sessionId: descriptor.sessionId,
    revision: 7,
  })
  expect(prepared.batches).toHaveLength(1)
  expect(
    await readPrivateBatch({ path: prepared.dataPath, ...prepared.batches[0]!, count: 2 }),
  ).toEqual([
    { kind: 'number', value: 42 },
    { kind: 'string', value: '=literal' },
  ])
  expect(state.sheets[0]!.cells.A1!.rawValue).toBe(120)
  await writeFile(join(source.directory, 'result.ndjson'), 'changed after preparation')
  expect(await readFile(prepared.dataPath, 'utf8')).toContain('=literal')
  await expect(
    prepareResultCopy({
      source,
      resultManifestPath,
      privateStorage,
      state,
      sessionId: descriptor.sessionId,
      revision: 7,
    }),
  ).rejects.toMatchObject({ code: 'artifact_hash_mismatch' })
  await writeFile(source.manifestPath, JSON.stringify({ ...source.manifest, sourceRevision: 8 }))
  await writeFile(
    resultManifestPath,
    JSON.stringify({ ...result, source: { ...result.source, sourceRevision: 8 } }),
  )
  await expect(
    prepareResultCopy({
      source: snapshots.get(source.manifest.snapshotId),
      resultManifestPath,
      privateStorage,
      state,
      sessionId: descriptor.sessionId,
      revision: 8,
    }),
  ).rejects.toMatchObject({ code: 'source_mismatch' })
  await writeFile(join(workspace, 'outside.json'), JSON.stringify(result))
  await expect(
    prepareResultCopy({
      source,
      resultManifestPath: join(workspace, 'outside.json'),
      privateStorage,
      state,
      sessionId: descriptor.sessionId,
      revision: 7,
    }),
  ).rejects.toMatchObject({ code: 'path_not_allowed' })
  await prepared.dispose()
})
