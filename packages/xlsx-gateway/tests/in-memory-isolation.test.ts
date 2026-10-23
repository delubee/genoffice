import { expect, it } from 'vitest'
import { InMemoryWorkbookAdapter } from '../src/domain/in-memory-workbook'

it('keeps caller, captured and undo snapshots isolated while updating a private batch in place', () => {
  const original = {
    id: 'book',
    revision: 0,
    sheets: [{ id: 'sheet', name: 'Data', cells: { A1: { value: 1 }, A2: { value: 2 } } }],
  }
  const adapter = new InMemoryWorkbookAdapter(original)
  const captured = adapter.getSnapshot()
  const plan = adapter.plan({
    dslVersion: 1,
    transactionId: 'one-batch',
    baseRevision: 0,
    summary: 'A private batch',
    operations: [
      { op: 'set_cell', sheetId: 'sheet', address: 'A1', value: 3 },
      { op: 'set_cell', sheetId: 'sheet', address: 'A1', value: 4 },
      { op: 'set_cell', sheetId: 'sheet', address: 'A2', value: 5 },
    ],
  })
  expect(original.sheets[0]!.cells).toEqual({ A1: { value: 1 }, A2: { value: 2 } })
  expect(adapter.getSnapshot()).toEqual(captured)
  expect(plan.cellChanges[1]!.before.value).toBe(3)
  adapter.apply(plan)
  expect(adapter.getSnapshot().sheets[0]!.cells).toMatchObject({
    A1: { value: 4 },
    A2: { value: 5 },
  })
  expect(captured.sheets[0]!.cells.A1!.value).toBe(1)
  expect(original.sheets[0]!.cells.A1.value).toBe(1)
  adapter.undo()
  expect(adapter.getSnapshot().sheets[0]!.cells).toEqual({ A1: { value: 1 }, A2: { value: 2 } })
  expect(captured.sheets[0]!.cells.A2!.value).toBe(2)
})
