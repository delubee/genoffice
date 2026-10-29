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
    {
      op: 'set_col_width',
      sheetId: id,
      column: 'B',
      count: 1,
      widthPx: 140,
      widthCharacters: 17.5,
    },
  ])
  expect(state.sheets[0]!.features).toMatchObject({
    rowHeights: { '2': 30 },
    colWidths: { B: 140 },
    columnWidthsCharacters: { B: 17.5 },
  })
  expect(state.sheets[0]!.cells).toEqual({})
  await apply([
    { op: 'insert_rows', sheetId: id, row: 1, count: 1 },
    { op: 'insert_cols', sheetId: id, column: 'A', count: 1 },
  ])
  expect(state.sheets[0]!.features).toMatchObject({
    rowHeights: { '3': 30 },
    colWidths: { C: 140 },
    columnWidthsCharacters: { C: 17.5 },
  })
})

it('matches desktop dialog units, chunks a large axis selection and accepts zero and maximum sizes', async () => {
  const { ribbonIntents } = await import('@genoffice/sheets-editor/ribbon-intents')
  const { toSelectionFormat } = await import('../../../apps/sheets/src/renderer/selection-format')
  const target = {
    sheetId: 'sheet',
    address: 'B2:C3',
    maximumDigitWidth: 8,
    format: toSelectionFormat({}, 'General'),
    doubleUnderline: false,
  }
  expect(ribbonIntents('row-height:25.3', target)).toMatchObject([
    { op: 'set_row_height', row: 2, count: 2, heightPoints: 25.5 },
  ])
  expect(ribbonIntents('col-width:10', target)).toMatchObject([
    { op: 'set_col_width', column: 'B', count: 2, widthPx: 80, widthCharacters: 10 },
  ])
  expect(ribbonIntents('row-height:409.5', target)[0]).toMatchObject({ heightPoints: 409.5 })
  expect(ribbonIntents('row-height:0', target)[0]).toMatchObject({ heightPoints: 0 })
  expect(ribbonIntents('col-width:0', target)[0]).toMatchObject({ widthPx: 0, widthCharacters: 0 })
  expect(ribbonIntents('row-height:30', { ...target, address: 'A1:A1001' })).toHaveLength(3)
  expect(() => ribbonIntents('row-height:410', target)).toThrow()
  expect(() => ribbonIntents('col-width:256', target)).toThrow()
})
