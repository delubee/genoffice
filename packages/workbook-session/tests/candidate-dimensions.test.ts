import { expect, it } from 'vitest'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import { initialState } from '../src/state'
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'

it('persists row heights and column widths without materializing cell edits and shifts them with structure', async () => {
  let state = initialState({ sheets: [{ id: 'base', name: 'Data', rowCount: 4, columnCount: 4 }] })
  const id = state.sheets[0]!.id
  const apply = async (operations: WorkbookOperation[]) => {
    const before = structuredClone(state)
    const result = await computeCandidateState(
      state,
      {
        schemaVersion: 1,
        sessionId: 'sizes',
        transactionId: crypto.randomUUID(),
        expectedRevision: 0,
        origin: 'manual',
        summary: 'Resize',
        operations,
      },
      {
        state: structuredClone(state),
        completeSheetIds: [id],
        writableSheetIds: [id],
        allocation: { stateId: crypto.randomUUID(), sheetIds: {} },
      },
    )
    expect(result.inverse.state).toEqual(before)
    state = result.nextState
  }
  await apply([
    { op: 'set_row_height', sheetId: id, row: 2, count: 1, heightPoints: 30 },
    { op: 'set_col_width', sheetId: id, column: 'B', count: 1, widthPx: 140 },
  ])
  expect(state.sheets[0]!.features).toMatchObject({
    rowHeights: { '2': 30 },
    colWidths: { B: 140 },
  })
  expect(state.sheets[0]!.cells).toEqual({})
  await apply([
    { op: 'insert_rows', sheetId: id, row: 1, count: 1 },
    { op: 'insert_cols', sheetId: id, column: 'A', count: 1 },
  ])
  expect(state.sheets[0]!.features).toMatchObject({
    rowHeights: { '3': 30 },
    colWidths: { C: 140 },
  })
})
