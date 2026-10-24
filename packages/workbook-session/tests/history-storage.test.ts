import { expect, it } from 'vitest'
import { randomUUID } from 'node:crypto'
import { join } from 'node:path'
import { initialState } from '../src/state'
import { FileResourceStore } from '../src/store'
import { WorkbookTransactions } from '../src/transactions'
import { retainedHistoryStorage } from '../src/history-records'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import { useSessionTestResources } from './helpers'
const resources = useSessionTestResources()

it('charges shared serialized history chunks once while retaining every undo root and the per-transaction cap', async () => {
  const directory = await resources.directory(),
    store = await FileResourceStore.create(join(directory, 'store'))
  const state = structuredClone(
      initialState({ sheets: [{ id: 'base', name: 'Data', rowCount: 25, columnCount: 1 }] }),
    ),
    sheetId = state.sheets[0]!.id
  state.sheets[0]!.cells.A1 = { kind: 'number', rawValue: 0 }
  for (let row = 2; row <= 21; row++)
    state.sheets[0]!.cells[`A${row}`] = { kind: 'string', rawValue: 'x'.repeat(32767) }
  const engine = new WorkbookTransactions({
    sessionId: 'storage-budget',
    documentUri: 'file:///history.xlsx',
    state,
    store,
    historyByteLimit: 8 * 1024 * 1024,
    interpreter: {
      prepare: (before, intent) =>
        computeCandidateState(before, intent, {
          state: structuredClone(before),
          completeSheetIds: [sheetId],
          writableSheetIds: [sheetId],
          allocation: { stateId: randomUUID(), sheetIds: {} },
        }),
      dispose() {},
    },
    host: { committed() {}, projectionChanged() {}, conflict() {} },
  })
  try {
    for (let revision = 0; revision < 10; revision++)
      await engine.submit({
        schemaVersion: 1,
        sessionId: 'storage-budget',
        transactionId: String(revision),
        expectedRevision: revision,
        origin: 'manual',
        summary: 'Small edit over a shared large root',
        operations: [{ op: 'set_cell', sheetId, address: 'A1', value: (revision + 1) % 10 }],
      })
    expect(engine.history).toHaveLength(10)
    const retained = await retainedHistoryStorage(store, engine.history)
    expect(retained.bytes).toBeLessThan(
      engine.history.reduce((total, entry) => total + entry.serializedBytes, 0) * 0.75,
    )
    for (let index = 0; index < 10; index++) await engine.undo()
    expect(engine.state.sheets[0]!.cells.A1!.rawValue).toBe(0)
    expect(engine.descriptor.revision).toBe(20)
    expect(engine.history).toHaveLength(10)
  } finally {
    await engine.dispose()
  }
})
