import { randomUUID } from 'node:crypto'
import { expect, it } from 'vitest'
import { WorkbookBindings } from '../src/bindings'
import { captureReadView, assertFreshRevision } from '../src/freshness'
import { WorkbookTransactions } from '../src/transactions'
import { FileResourceStore } from '../src/store'
import { initialState } from '../src/state'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import { useSessionTestResources } from './helpers'
const resources = useSessionTestResources()

async function session(bindings: WorkbookBindings, uri = 'file:///one/sales.xlsx') {
  const state = initialState({
      sheets: [{ id: 'base', name: 'Data', rowCount: 3, columnCount: 2 }],
    }),
    sessionId = randomUUID(),
    sheetId = state.sheets[0]!.id
  const engine = new WorkbookTransactions({
    sessionId,
    documentUri: uri,
    state,
    store: await FileResourceStore.create(await resources.directory()),
    interpreter: {
      async prepare(before, intent) {
        return computeCandidateState(before, intent, {
          state: before,
          completeSheetIds: before.sheets.map((sheet) => sheet.id),
          writableSheetIds: before.sheets.map((sheet) => sheet.id),
          allocation: { stateId: randomUUID(), sheetIds: {} },
        })
      },
      dispose() {},
    },
    host: {
      committed() {
        bindings.observe(engine.descriptor, engine.state)
      },
      projectionChanged() {},
      conflict() {},
    },
  })
  bindings.observe(engine.descriptor, engine.state)
  const binding = bindings.bind(sessionId, sheetId, 'A1:A3')
  const edit = (value: number) =>
    engine.submit({
      schemaVersion: 1,
      sessionId,
      transactionId: randomUUID(),
      expectedRevision: engine.descriptor.revision,
      origin: 'manual',
      summary: 'Edit',
      operations: [{ op: 'set_cell', sheetId, address: 'A1', value }],
    })
  return { engine, binding, sheetId, edit }
}

it('rejects ABA revisions while keeping an immutable concurrent read view', async () => {
  const bindings = new WorkbookBindings(),
    { engine, binding, edit } = await session(bindings)
  await edit(1)
  const view = captureReadView(engine.descriptor, engine.state, 1)
  await edit(2)
  await edit(1)
  expect(engine.state.sheets[0]!.cells.A1!.rawValue).toBe(1)
  expect(view.state.sheets[0]!.cells.A1!.rawValue).toBe(1)
  expect(view.descriptor.revision).toBe(1)
  expect(engine.descriptor.revision).toBe(3)
  expect(() => assertFreshRevision(engine.descriptor, engine.descriptor.sessionId, 1)).toThrow(
    'changed',
  )
  expect(bindings.get(binding.bindingId).status).toBe('stale')
  expect(() => bindings.authorizeWrite(binding.bindingId, 3)).toThrow(/read/i)
  bindings.acknowledgeRead(binding.bindingId, 3)
  expect(bindings.authorizeWrite(binding.bindingId, 3).documentUri).toBe('file:///one/sales.xlsx')
})

it('keeps structural selection staleness until explicit rebinding even after a fresh read', async () => {
  const bindings = new WorkbookBindings(),
    { engine, binding, sheetId } = await session(bindings)
  await engine.submit({
    schemaVersion: 1,
    sessionId: engine.descriptor.sessionId,
    transactionId: 'insert',
    expectedRevision: 0,
    origin: 'manual',
    summary: 'Insert',
    operations: [{ op: 'insert_rows', sheetId, row: 1, count: 1 }],
  })
  bindings.acknowledgeRead(binding.bindingId, 1)
  expect(bindings.get(binding.bindingId).selectionStatus).toBe('selection_stale_structure')
  expect(() => bindings.authorizeWrite(binding.bindingId, 1, true)).toThrow('bind')
  expect(bindings.authorizeWrite(binding.bindingId, 1, false).sessionId).toBe(
    engine.descriptor.sessionId,
  )
  const rebound = bindings.bind(engine.descriptor.sessionId, sheetId, 'A1:A3')
  expect(bindings.authorizeWrite(rebound.bindingId, 1, true).selectionStatus).toBe('current')
})

it('separates same-name documents, windows and closed/reopened or reverted session epochs', async () => {
  const bindings = new WorkbookBindings(),
    a = await session(bindings),
    b = await session(bindings, 'file:///two/sales.xlsx')
  await b.edit(7)
  expect(bindings.get(a.binding.bindingId).documentUri).toBe('file:///one/sales.xlsx')
  expect(bindings.get(a.binding.bindingId).status).toBe('active')
  expect(() => new WorkbookBindings().get(a.binding.bindingId)).toThrow('binding')
  bindings.invalidateSession(a.engine.descriptor.sessionId)
  await a.engine.dispose()
  const reopened = await session(bindings)
  expect(reopened.engine.descriptor.sessionId).not.toBe(a.engine.descriptor.sessionId)
  expect(bindings.get(a.binding.bindingId).status).toBe('invalidated')
  expect(() => bindings.authorizeWrite(a.binding.bindingId, 0)).toThrow('closed')
  expect(() =>
    assertFreshRevision(reopened.engine.descriptor, a.engine.descriptor.sessionId, 0),
  ).toThrow('session')
})
