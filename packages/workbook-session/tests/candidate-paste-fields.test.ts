import { expect, it } from 'vitest'
import { initialState } from '../src/state'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
function fixture() {
  const state = structuredClone(
    initialState({ sheets: [{ id: 'base', name: 'Data', rowCount: 20, columnCount: 10 }] }),
  )
  const sheet = state.sheets[0]!
  sheet.cells.D5 = {
    kind: 'number',
    rawValue: 12,
    formula: '=A1',
    calculationState: 'ready',
    cachedValue: 12,
    numberFormat: '0.00',
    style: {
      italic: true,
      fillColor: '#FF0000',
      numberFormat: '0.00',
      borderLeft: { style: 'double', color: '#123456' },
    },
  }
  const apply = (options: Partial<Extract<WorkbookOperation, { op: 'paste_cells' }>>) =>
    computeCandidateState(
      state,
      {
        schemaVersion: 1,
        sessionId: 'selective',
        transactionId: crypto.randomUUID(),
        expectedRevision: 0,
        origin: 'manual',
        summary: 'Paste Special',
        operations: [
          {
            op: 'paste_cells',
            sheetId: sheet.id,
            start: 'D5',
            cells: [[{ kind: 'string', rawValue: '=literal', style: { bold: true } }]],
            ...options,
          },
        ],
      },
      {
        state: structuredClone(state),
        completeSheetIds: [sheet.id],
        writableSheetIds: [sheet.id],
        allocation: { stateId: crypto.randomUUID(), sheetIds: {} },
      },
    )
  return { state, sheet, apply }
}
it('values-only preserves target formatting and formats-only preserves target formulas and values', async () => {
  const f = fixture()
  const values = await f.apply({ format: 'none' })
  expect(values.nextState.sheets[0]!.cells.D5).toMatchObject({
    kind: 'string',
    rawValue: '=literal',
    numberFormat: '0.00',
    style: { italic: true, fillColor: '#FF0000', borderLeft: { style: 'double' } },
  })
  expect(values.nextState.sheets[0]!.cells.D5!.formula).toBeUndefined()
  const formats = await f.apply({ writeValues: false })
  expect(formats.nextState.sheets[0]!.cells.D5).toMatchObject({
    rawValue: 12,
    formula: '=A1',
    styleReset: true,
    style: { bold: true },
  })
  expect(formats.nextState.sheets[0]!.cells.D5!.style!.italic).toBeUndefined()
  expect(formats.inverse.state).toEqual(f.state)
})
it('number formats and except-border pastes replace only their selected fields', async () => {
  const f = fixture()
  const number = await f.apply({ format: 'number' })
  expect(number.nextState.sheets[0]!.cells.D5).toMatchObject({
    numberFormat: 'General',
    style: { italic: true, fillColor: '#FF0000', borderLeft: { style: 'double' } },
  })
  const borders = await f.apply({ format: 'except-borders' })
  expect(borders.nextState.sheets[0]!.cells.D5).toMatchObject({
    styleReset: true,
    style: { bold: true, borderLeft: { style: 'double', color: '#123456' }, borderTop: null },
  })
  expect(borders.nextState.sheets[0]!.cells.D5!.style!.fillColor).toBeUndefined()
})
it('skips selected cells atomically and keeps a 101-cell styled paste within one operation', async () => {
  const f = fixture()
  const skipped = await f.apply({ skipCells: ['D5'] })
  expect(skipped.nextState.sheets[0]!.cells.D5!.rawValue).toBe(12)
  await expect(f.apply({ skipCells: ['A1'] })).rejects.toThrow()
  const bulk = await f.apply({
    cells: [
      Array.from({ length: 101 }, (_, i) => ({
        kind: 'number' as const,
        rawValue: i,
        style: { bold: i % 2 === 0 },
      })),
    ],
    format: 'all',
  })
  expect(bulk.nextState.sheets[0]!.cells.CZ5).toMatchObject({
    rawValue: 100,
    style: { bold: true },
  })
})
