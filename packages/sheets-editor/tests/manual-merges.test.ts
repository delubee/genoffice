import { expect, it } from 'vitest'
import { manualValueIntents } from '../src/manual-intents'
const range = { startRow: 0, startColumn: 0, endRow: 1, endColumn: 2 }
it('coalesces repeated scalar values across a merged selection into one anchor', () => {
  expect(manualValueIntents('sheet', { range, value: { v: 'Heading' } }, [range])).toEqual([
    { op: 'set_cell', sheetId: 'sheet', address: 'A1', value: 'Heading' },
  ])
})
it('rejects conflicting pasted values and covered-only edits instead of discarding data', () => {
  expect(() =>
    manualValueIntents('sheet', { range, value: [[{ v: 'One' }, { v: 'Two' }]] }, [range]),
  ).toThrow('merged_cell_target')
  expect(() =>
    manualValueIntents('sheet', { value: { 1: { 1: { v: 'Hidden' } } } }, [range]),
  ).toThrow('merged_cell_target')
})
