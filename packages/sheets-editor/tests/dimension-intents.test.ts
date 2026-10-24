import { expect, it } from 'vitest'
import { dimensionIntents } from '../src/dimension-intents'

it('converts row pixels to Excel points and splits columns within DSL limits', () => {
  const range = { startRow: 1, endRow: 2, startColumn: 0, endColumn: 100 }
  expect(dimensionIntents('sheet', 'row', [range], 40)).toEqual([
    { op: 'set_row_height', sheetId: 'sheet', row: 2, count: 2, heightPoints: 30 },
  ])
  expect(dimensionIntents('sheet', 'column', [range], 140)).toEqual([
    { op: 'set_col_width', sheetId: 'sheet', column: 'A', count: 100, widthPx: 140 },
    { op: 'set_col_width', sheetId: 'sheet', column: 'CW', count: 1, widthPx: 140 },
  ])
  expect(() => dimensionIntents('sheet', 'row', [range], NaN)).toThrow('invalid_dimension')
})
