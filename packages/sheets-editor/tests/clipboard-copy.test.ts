import { expect, it } from 'vitest'
import { clipboardCopyId, copiedBlockOperation, createCopiedBlock } from '../src/clipboard-copy'
import type { ViewCell } from '../src/grid-cell'

const bounds = { startRow: 1, endRow: 1, startColumn: 0, endColumn: 1 }
const cells: ViewCell[] = [
  { address: 'A2', kind: 'string', rawValue: '=literal', style: { bold: true } },
  {
    address: 'B2',
    kind: 'number',
    rawValue: 2,
    formula: '=A2+$A$1',
    calculationState: 'ready',
    cachedValue: 2,
  },
]

it('copies a detached typed snapshot and shifts mixed references across sheets', () => {
  const read = structuredClone(cells)
  const block = createCopiedBlock('identity', bounds, { cells: read }, [1], [0, 1])
  read[0]!.rawValue = 'Changed'
  read[0]!.style!.bold = false
  const operation = copiedBlockOperation(block, 'another-sheet', 'D5')
  expect(operation).toMatchObject({
    op: 'paste_cells',
    sheetId: 'another-sheet',
    start: 'D5',
    cells: [
      [{ kind: 'string', rawValue: '=literal', style: { bold: true } }, { formula: '=D5+$A$1' }],
    ],
    merges: [],
  })
  expect(block.cells[0]![1]!.formula).toBe('=A2+$A$1')
  expect(clipboardCopyId('<table data-copy-id="identity">')).toBe('identity')
  expect(clipboardCopyId("<table data-copy-id='identity'>")).toBe('identity')
  expect(clipboardCopyId('<table>')).toBeNull()
})

it('clears covered source values, repeats merges and rejects partial merges or missing source cells', () => {
  const merge = { ...bounds }
  const block = createCopiedBlock('merged', bounds, { cells, merges: [merge] }, [1], [0, 1])
  expect(block.cells[0]![1]).toEqual({ kind: 'blank', rawValue: null })
  expect(copiedBlockOperation(block, 'sheet', 'D5:G6')).toMatchObject({
    merges: ['D5:E5', 'F5:G5', 'D6:E6', 'F6:G6'],
  })
  expect(() =>
    createCopiedBlock('partial', { ...bounds, endColumn: 0 }, { cells, merges: [merge] }, [1], [0]),
  ).toThrow()
  expect(() =>
    createCopiedBlock('missing', bounds, { cells: cells.slice(0, 1) }, [1], [0, 1]),
  ).toThrow()
  expect(() =>
    createCopiedBlock('filtered-merge', bounds, { cells, merges: [merge] }, [1], [0]),
  ).toThrow()
})

it('bounds repeated writes and uses source dimensions when a selected target contains merges', () => {
  const block = createCopiedBlock('identity', bounds, { cells }, [1], [0, 1])
  expect(() => copiedBlockOperation(block, 'sheet', 'A1:B1001')).toThrow()
  expect(() => copiedBlockOperation(block, 'sheet', 'XFD1048576')).toThrow()
  const operation = copiedBlockOperation(block, 'sheet', 'D5:G6', [
    { startRow: 4, endRow: 4, startColumn: 3, endColumn: 4 },
  ])
  expect(operation).toMatchObject({ start: 'D5', cells: [[expect.anything(), expect.anything()]] })
})
