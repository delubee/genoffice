import { expect, it } from 'vitest'
import { WorkbookTransactions } from '../src/transactions'
import { FileResourceStore } from '../src/store'
import { initialState } from '../src/state'
import type { CandidateInterpreter, HostLifecycle } from '../src/ports'
import { useSessionTestResources } from './helpers'
const resources = useSessionTestResources()

async function setup(
  options: {
    failPrepare?: boolean
    failRoot?: boolean
    failProjection?: boolean
    onPrepared?(): void
    onCommitted?(): void
    transactionByteLimit?: number
    historyByteLimit?: number
  } = {},
) {
  const state = initialState({
    sheets: [{ id: 'base', name: 'Data', rowCount: 10, columnCount: 3 }],
  })
  const directory = await resources.directory()
  let prepares = 0,
    commits = 0
  const interpreter: CandidateInterpreter = {
    async prepare(before) {
      prepares++
      if (options.failPrepare) throw new Error('Injected candidate failure')
      const after = { ...structuredClone(before), stateId: `state-${prepares}` }
      after.sheets[0]!.cells.A1 = { kind: 'number', rawValue: prepares }
      options.onPrepared?.()
      return {
        beforeStateId: before.stateId,
        nextState: after,
        forward: { schemaVersion: 1, fromStateId: before.stateId, state: after },
        inverse: { schemaVersion: 1, fromStateId: after.stateId, state: before },
        affectedTargets: [after.sheets[0]!.id],
        calculationState: 'ready',
      }
    },
    dispose() {},
  }
  let manifestName = ''
  const store = await FileResourceStore.create(directory, (point) => {
    if (options.failRoot && point === 'beforeManifestRename' && manifestName === 'current')
      throw new Error('Injected storage failure')
  })
  const publish = store.publishManifest.bind(store)
  store.publishManifest = async (name, data) => {
    manifestName = name
    await publish(name, data)
  }
  const host: HostLifecycle = {
    committed() {
      commits++
      options.onCommitted?.()
    },
    projectionChanged() {
      if (options.failProjection) throw new Error('Injected projection failure')
    },
    conflict() {},
  }
  const engine = new WorkbookTransactions({
    transactionByteLimit: options.transactionByteLimit,
    historyByteLimit: options.historyByteLimit,
    sessionId: 'session',
    documentUri: 'file:///test.xlsx',
    state,
    store,
    interpreter,
    host,
  })
  const intent = (id: string, revision = 0) => ({
    schemaVersion: 1 as const,
    sessionId: 'session',
    transactionId: id,
    expectedRevision: revision,
    origin: 'manual' as const,
    summary: 'Edit cell',
    operations: [
      { op: 'set_cell' as const, sheetId: state.sheets[0]!.id, address: 'A1', value: 1 },
    ],
  })
  return { engine, store, state, intent, counts: () => ({ prepares, commits }) }
}
it('serializes commits, deduplicates before revision checks and rejects reused payloads', async () => {
  const { engine, intent, counts } = await setup()
  const first = await engine.submit(intent('first'))
  expect(first.revision).toBe(1)
  expect(await engine.submit(intent('first'))).toEqual(first)
  expect(counts()).toEqual({ prepares: 1, commits: 1 })
  await expect(
    engine.submit({ ...intent('first'), summary: 'Different payload' }),
  ).rejects.toMatchObject({ code: 'idempotency_conflict' })
  await expect(engine.submit(intent('stale'))).rejects.toMatchObject({ code: 'revision_conflict' })
  const outcomes = await Promise.allSettled([
    engine.submit(intent('second', 1)),
    engine.submit(intent('third', 1)),
  ])
  expect(outcomes.map((outcome) => outcome.status)).toEqual(['fulfilled', 'rejected'])
  expect(engine.descriptor.revision).toBe(2)
})
it('rejects candidate/storage failures and precommit cancellation without changing the canonical root', async () => {
  for (const options of [{ failPrepare: true }, { failRoot: true }]) {
    const { engine, state, intent, counts } = await setup(options)
    await expect(engine.submit(intent('failed'))).rejects.toThrow('Injected')
    expect(engine.state).toEqual(state)
    expect(engine.descriptor.revision).toBe(0)
    expect(counts().commits).toBe(0)
  }
  const { engine, intent } = await setup(),
    controller = new AbortController()
  controller.abort()
  await expect(engine.submit(intent('cancelled'), controller.signal)).rejects.toMatchObject({
    code: 'operation_cancelled',
  })
  expect(engine.descriptor.revision).toBe(0)
})
it('returns durable committed receipts after projection failure and lost-response retries', async () => {
  const { engine, store, intent, counts } = await setup({ failProjection: true })
  const receipt = await engine.submit(intent('committed'))
  expect(engine.descriptor.dirty).toBe(true)
  expect(engine.projectionError).toBeDefined()
  expect(await engine.operationStatus('committed')).toMatchObject({ status: 'committed', receipt })
  expect((await store.readManifest('current')).payload).toMatchObject({
    revision: 1,
    stateId: receipt.stateId,
  })
  expect(await engine.submit(intent('committed'))).toEqual(receipt)
  expect(counts().commits).toBe(1)
})

it('distinguishes cancellation before root publication from cancellation after commit', async () => {
  const before = new AbortController()
  const pending = await setup({
    onPrepared() {
      before.abort()
    },
  })
  await expect(
    pending.engine.submit(pending.intent('before-root'), before.signal),
  ).rejects.toMatchObject({ code: 'operation_cancelled' })
  expect(pending.engine.descriptor.revision).toBe(0)
  expect(pending.counts().commits).toBe(0)
  const after = new AbortController()
  const committed = await setup({
    onCommitted() {
      after.abort()
    },
  })
  const receipt = await committed.engine.submit(committed.intent('after-root'), after.signal)
  expect(after.signal.aborted).toBe(true)
  expect(receipt.status).toBe('committed')
  expect(committed.engine.descriptor.revision).toBe(1)
  expect(await committed.engine.operationStatus('after-root')).toMatchObject({
    status: 'committed',
    receipt,
  })
})

it('rejects oversized undo records for every write origin before any content/history publication', async () => {
  for (const origin of ['manual', 'agent', 'import', 'history'] as const) {
    const { engine, state, intent, counts } = await setup({ transactionByteLimit: 1 })
    const before = engine.state
    await expect(engine.submit({ ...intent(`budget-${origin}`), origin })).rejects.toMatchObject({
      code: 'undo_budget_exceeded',
    })
    expect(engine.state).toBe(before)
    expect(engine.state).toEqual(state)
    expect(engine.historyPosition).toEqual({ cursor: 0, length: 0 })
    expect(counts().commits).toBe(0)
  }
})

it('retains dirty content and history when the file was written but its saved marker could not be published', async () => {
  const faults = { failRoot: false },
    { engine, intent, store } = await setup(faults)
  await engine.submit(intent('before-save'))
  const current = engine.state,
    history = engine.history,
    savedStateId = engine.descriptor.savedStateId
  let fileWritten = false
  faults.failRoot = true
  await expect(
    engine.saveWith(async () => {
      fileWritten = true
    }),
  ).rejects.toMatchObject({ code: 'save_failed', details: { fileWritten: true } })
  expect(fileWritten).toBe(true)
  expect(engine.state).toBe(current)
  expect(engine.history).toEqual(history)
  expect(engine.descriptor).toMatchObject({ dirty: true, savedStateId })
  const beforeUndo = engine.historyPosition
  await expect(engine.undo()).rejects.toThrow('storage failure')
  expect(engine.historyPosition).toEqual(beforeUndo)
  expect(engine.state).toBe(current)
  faults.failRoot = false
  await engine.undo()
  await engine.redo()
  await engine.saveWith(async () => {})
  const manifest = await store.readManifest('current'),
    chunks = (manifest.payload as { chunks: string[] }).chunks
  const root = JSON.parse(
    Buffer.concat(await Promise.all(chunks.map((chunk) => store.readChunk(chunk)))).toString(
      'utf8',
    ),
  )
  expect(root.savedStateId).toBe(current.stateId)
  expect(root.cursor).toBe(1)
  expect(engine.descriptor.dirty).toBe(false)
})
