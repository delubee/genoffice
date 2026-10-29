import { expect, it } from 'vitest'
import { createCopiedBlock } from '../src/clipboard-copy'
import {
  defaultPasteSpecial as defaults,
  specialPasteOperations,
  type SpecialPasteSnapshot,
} from '../src/special-paste'
import type { ViewCell } from '../src/grid-cell'
const cells: ViewCell[] = [
  { address: 'A1', kind: 'number', rawValue: 2, style: { bold: true }, numberFormat: '0.00' },
  { address: 'B1', kind: 'blank', rawValue: null },
  { address: 'A2', kind: 'number', rawValue: 4 },
  { address: 'B2', kind: 'number', rawValue: 0 },
]
const block = createCopiedBlock(
  'copy',
  { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
  { cells },
  [0, 1],
  [0, 1],
)
const snapshot: SpecialPasteSnapshot = {
  source: {
    kind: 'internal',
    unitId: 'projection',
    subUnitId: 'source',
    rows: [0, 1],
    cols: [0, 1],
    cut: false,
  },
  block,
  sheetId: 'source',
  widths: [80, 120],
  html: '',
  text: '',
}
const target = {
  sheetId: 'target',
  address: 'D5',
  cells: [{ address: 'D5', kind: 'number' as const, rawValue: 10 }],
  linkSheetName: 'Original Data',
}
it('uses the shared desktop arithmetic, transpose and skip-blank transform in one typed operation', () => {
  const result = specialPasteOperations(
    snapshot,
    { ...defaults, type: 'values', operation: 'add', transpose: true, skipBlanks: true },
    target,
  )
  expect(result).toHaveLength(1)
  expect(result[0]).toMatchObject({
    op: 'paste_cells',
    start: 'D5',
    format: 'none',
    writeValues: true,
    skipCells: ['D6'],
    cells: [
      [{ rawValue: 12 }, { rawValue: 4 }],
      [{ rawValue: null }, { rawValue: 0 }],
    ],
  })
})
it('links to original source coordinates across sheets even when Formats is selected', () => {
  expect(
    specialPasteOperations(snapshot, { ...defaults, type: 'formats', link: true }, target)[0],
  ).toMatchObject({
    writeValues: true,
    format: 'none',
    cells: [
      [{ formula: "='Original Data'!A1" }, { formula: "='Original Data'!B1" }],
      [{ formula: "='Original Data'!A2" }, { formula: "='Original Data'!B2" }],
    ],
  })
})
it('preserves literal strings and errors and represents division by zero as an error cell', () => {
  expect(
    specialPasteOperations(
      snapshot,
      { ...defaults, type: 'values', operation: 'divide' },
      target,
    )[0],
  ).toMatchObject({
    cells: [
      [{ rawValue: 5 }, expect.anything()],
      [expect.anything(), { kind: 'error', rawValue: '#DIV/0!' }],
    ],
  })
  const source = structuredClone(snapshot)
  source.block!.cells[0]![0] = { kind: 'string', rawValue: '=literal' }
  source.block!.cells[0]![1] = { kind: 'error', rawValue: '#N/A' }
  expect(specialPasteOperations(source, { ...defaults, type: 'values' }, target)[0]).toMatchObject({
    cells: [
      [
        { kind: 'string', rawValue: '=literal' },
        { kind: 'error', rawValue: '#N/A' },
      ],
      expect.anything(),
    ],
  })
})
it('keeps column widths separate from cell flags and rejects unknown formula values', () => {
  expect(
    specialPasteOperations(
      snapshot,
      { ...defaults, type: 'col-widths', transpose: true, operation: 'add' },
      target,
    ),
  ).toEqual([
    { op: 'set_col_width', sheetId: 'target', column: 'D', count: 1, widthPx: 80 },
    { op: 'set_col_width', sheetId: 'target', column: 'E', count: 1, widthPx: 120 },
  ])
  const pending = structuredClone(snapshot)
  pending.block!.cells[0]![0] = {
    kind: 'blank',
    rawValue: null,
    formula: '=A3',
    calculationState: 'pending',
  }
  expect(() => specialPasteOperations(pending, { ...defaults, type: 'values' }, target)).toThrow()
  expect(
    specialPasteOperations(pending, { ...defaults, type: 'formulas' }, target)[0],
  ).toMatchObject({ cells: [[{ formula: '=D7' }, expect.anything()], expect.anything()] })
})

it('does not demand computed values for transposed Formats or Paste Link with arithmetic flags', () => {
  const pending = structuredClone(snapshot)
  pending.block!.cells[0]![0] = {
    kind: 'blank',
    rawValue: null,
    formula: '=A3',
    calculationState: 'pending',
    style: { bold: true },
  }
  const formulaTarget = {
    ...target,
    cells: [
      {
        address: 'D5',
        kind: 'blank' as const,
        rawValue: null,
        formula: '=1+1',
        calculationState: 'pending' as const,
      },
    ],
  }
  expect(
    specialPasteOperations(
      pending,
      { ...defaults, type: 'formats', transpose: true, operation: 'add' },
      formulaTarget,
    )[0],
  ).toMatchObject({ writeValues: false, format: 'all' })
  expect(
    specialPasteOperations(
      pending,
      { ...defaults, type: 'col-widths', link: true, operation: 'add' },
      formulaTarget,
    )[0],
  ).toMatchObject({
    op: 'paste_cells',
    writeValues: true,
    cells: [[{ formula: "='Original Data'!A1" }, expect.anything()], expect.anything()],
  })
})
