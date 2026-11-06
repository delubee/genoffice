import { expect, it } from 'vitest'
import { functionPreviewRanges } from '../src/function-preview'
const sheets = [
  { id: 'data', name: 'Data', rows: 100, columns: 5 },
  { id: 'quoted', name: "Bob's Data", rows: 10, columns: 2 },
]
it('resolves static precedents and whole axes with the authoritative used dimensions', () => {
  expect(functionPreviewRanges('SUM(B2:B3)', 'data', sheets)).toEqual([
    { sheetId: 'data', range: { startRow: 1, endRow: 2, startColumn: 1, endColumn: 1 } },
  ])
  expect(functionPreviewRanges("SUM('Bob''s Data'!$A$1:$A$3)", 'data', sheets)?.[0]?.sheetId).toBe(
    'quoted',
  )
  expect(functionPreviewRanges('SUM(B:B)', 'data', sheets)?.[0]?.range.endRow).toBe(99)
  expect(functionPreviewRanges('SUM(2:3)', 'data', sheets)?.[0]?.range.endColumn).toBe(4)
  expect(functionPreviewRanges('SUM(1,2,"B9")', 'data', sheets)).toEqual([])
})
it('declines unverified dynamic, named, external, missing-sheet and oversized previews', () => {
  for (const expression of [
    'SUM(Sales)',
    'INDIRECT("B2")',
    'SUM(OFFSET(A1,2,3))',
    "'Missing'!A1",
    "'[Other.xlsx]Sheet1'!A1",
    'SUM(A1:XFD1048576)',
  ]) {
    expect(functionPreviewRanges(expression, 'data', sheets)).toBeNull()
  }
})

import { autoFunctionCells } from '../../../apps/sheets/src/renderer/auto-function-command'
it('shares multi-column AutoSum geometry, minimum selection height and the grid boundary', () => {
  expect(autoFunctionCells('SUM', 'B2:C3')).toEqual([
    { address: 'B4', formula: '=SUM(B2:B3)' },
    { address: 'C4', formula: '=SUM(C2:C3)' },
  ])
  expect(autoFunctionCells('COUNT', 'B2:C2')).toEqual([])
  expect(() => autoFunctionCells('SUM', 'B1048575:B1048576')).toThrow('invalid_summary_target')
})
