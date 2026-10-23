import { expect, it } from 'vitest'
import { prepareHistoryRecord, loadHistoryState } from '../src/history-records'
import { FileResourceStore } from '../src/store'
import { initialState } from '../src/state'
import { LIMITS } from '../src/domain'
import { useSessionTestResources } from './helpers'
const resources = useSessionTestResources()
it('preflights actual bidirectional UTF-8 bytes and retains serialized chart/structure states', async () => {
  const store = await FileResourceStore.create(await resources.directory())
  const before = initialState({
    sheets: [{ id: 'base', name: 'Data', rowCount: 10, columnCount: 3 }],
  })
  const after = structuredClone(before)
  after.stateId = 'after'
  after.sheets[0]!.features = { charts: [{ id: 'chart', title: '图表', values: [1, 2] }] }
  after.structuralLog.push({ op: 'insert_rows', sheetId: after.sheets[0]!.id, row: 1, count: 1 })
  const prepared = {
    beforeStateId: before.stateId,
    nextState: after,
    forward: { schemaVersion: 1, fromStateId: before.stateId, state: after },
    inverse: { schemaVersion: 1, fromStateId: after.stateId, state: before },
    affectedTargets: [after.sheets[0]!.id],
    calculationState: 'ready' as const,
  }
  const record = await prepareHistoryRecord({
    store,
    before,
    prepared,
    origin: 'manual',
    label: 'Insert and chart',
    retainedBytes: 0,
  })
  expect(record.serializedBytes).toBe(
    Buffer.byteLength(JSON.stringify(prepared.forward)) +
      Buffer.byteLength(JSON.stringify(prepared.inverse)),
  )
  expect(await loadHistoryState(store, record, 'forward')).toEqual(after)
  expect(await loadHistoryState(store, record, 'inverse')).toEqual(before)
  await expect(
    prepareHistoryRecord({
      store,
      before,
      prepared,
      origin: 'manual',
      label: 'Too big',
      retainedBytes: LIMITS.historyBytes,
    }),
  ).rejects.toMatchObject({ code: 'undo_budget_exceeded', details: { scope: 'retained' } })
  await expect(
    prepareHistoryRecord({
      store,
      before,
      prepared,
      origin: 'manual',
      label: 'Small budget injection',
      retainedBytes: 0,
      transactionByteLimit: 1,
    }),
  ).rejects.toMatchObject({ code: 'undo_budget_exceeded', details: { scope: 'transaction' } })
  const broken = { ...prepared, inverse: () => before }
  await expect(
    prepareHistoryRecord({
      store,
      before,
      prepared: broken,
      origin: 'manual',
      label: 'Closure',
      retainedBytes: 0,
    }),
  ).rejects.toThrow()
})
