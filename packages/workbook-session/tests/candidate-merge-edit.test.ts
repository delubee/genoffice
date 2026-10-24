import { expect, it } from 'vitest'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { initialState } from '../src/state'

const initial = () =>
  structuredClone(
    initialState({ sheets: [{ id: 'base', name: 'Data', rowCount: 6, columnCount: 4 }] }),
  )
async function apply(state: ReturnType<typeof initial>, operations: WorkbookOperation[]) {
  return computeCandidateState(
    state,
    {
      schemaVersion: 1,
      sessionId: 'merges',
      transactionId: crypto.randomUUID(),
      expectedRevision: 0,
      origin: 'agent',
      summary: 'Edit merge geometry',
      operations,
    },
    {
      state: structuredClone(state),
      completeSheetIds: [state.sheets[0]!.id],
      writableSheetIds: [state.sheets[0]!.id],
      allocation: { stateId: crypto.randomUUID(), sheetIds: {} },
    },
  )
}
it('merges without losing styles, keeps inverse state, and guards covered writes in the same batch', async () => {
  const state = initial(),
    sheetId = state.sheets[0]!.id
  state.sheets[0]!.cells = {
    A1: { kind: 'string', rawValue: 'Title' },
    B1: { kind: 'blank', rawValue: null, style: { bold: true } },
  }
  const result = await apply(state, [{ op: 'merge_cells', sheetId, range: 'A1:C2' }])
  expect(result.inverse.state).toEqual(state)
  expect(result.nextState.sheets[0]!.cells).toEqual(state.sheets[0]!.cells)
  expect(result.nextState.sheets[0]!.features?.mergeEdits).toEqual([
    { startRow: 0, endRow: 1, startColumn: 0, endColumn: 2 },
  ])
  await expect(
    apply(state, [
      { op: 'merge_cells', sheetId, range: 'A1:C2' },
      { op: 'set_cell', sheetId, address: 'B1', value: 'Hidden' },
    ]),
  ).rejects.toThrow('merged_cell_target')
  await expect(
    apply(result.nextState, [{ op: 'insert_rows', sheetId, row: 1, count: 1 }]),
  ).rejects.toThrow('unsupported_structural_feature')
})
it('rejects data loss, including zero, false, empty strings and formulas with blank results', async () => {
  for (const cell of [
    { kind: 'number' as const, rawValue: 0 },
    { kind: 'boolean' as const, rawValue: false },
    { kind: 'string' as const, rawValue: '' },
    {
      kind: 'blank' as const,
      rawValue: null,
      formula: '=""',
      calculationState: 'pending' as const,
    },
  ]) {
    const state = initial(),
      sheetId = state.sheets[0]!.id
    state.sheets[0]!.cells.B1 = cell
    const before = structuredClone(state)
    await expect(apply(state, [{ op: 'merge_cells', sheetId, range: 'A1:B1' }])).rejects.toThrow(
      'merge_nonempty_cells',
    )
    expect(state).toEqual(before)
  }
})
it('unmerges an original range by selecting a covered cell and permits a following write', async () => {
  const state = initial(),
    sheetId = state.sheets[0]!.id
  state.sheets[0]!.features = { merges: [{ startRow: 0, endRow: 1, startColumn: 0, endColumn: 2 }] }
  state.sheets[0]!.cells.A1 = { kind: 'string', rawValue: 'Title' }
  const result = await apply(state, [
    { op: 'unmerge_cells', sheetId, range: 'B2' },
    { op: 'set_cell', sheetId, address: 'B2', value: 'Visible' },
  ])
  expect(result.nextState.sheets[0]!.features?.mergeEdits).toEqual([])
  expect(result.nextState.sheets[0]!.cells.B2?.rawValue).toBe('Visible')
  expect(result.nextState.sheets[0]!.cells.A1?.rawValue).toBe('Title')
  await expect(apply(state, [{ op: 'merge_cells', sheetId, range: 'B1:D3' }])).rejects.toThrow(
    'merge_partial_overlap',
  )
  await expect(apply(state, [{ op: 'merge_cells', sheetId, range: 'D4' }])).rejects.toThrow(
    'merge_select_range',
  )
})

it('rejects merges over tables, pivot ranges and active filter ranges before saving', async () => {
  const range = { startRow: 0, endRow: 1, startColumn: 0, endColumn: 2 }
  const cases: NonNullable<ReturnType<typeof initial>['sheets'][number]['features']>[] = [
    { originalMetadata: { tables: [{ range }] } },
    { originalMetadata: { pivotRanges: [range] } },
    { autoFilter: range },
    { filterEdits: { filter: { range } } },
  ]
  for (const features of cases) {
    const state = initial(),
      sheetId = state.sheets[0]!.id
    state.sheets[0]!.features = features
    await expect(apply(state, [{ op: 'merge_cells', sheetId, range: 'A1:C2' }])).rejects.toThrow(
      'merged_range_operation',
    )
  }
})
