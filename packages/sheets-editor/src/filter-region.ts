import { ObjectMatrix, type ICellData } from '@univerjs/core'
import { expandToContinuousRange } from '@univerjs/sheets'
import {
  formatAddress,
  parseAddress,
  rangeCellCount,
} from '@genoffice/xlsx-gateway/domain/cell-address'
import type { GridRange } from './ports'
import { gridCell, type ViewCell } from './grid-cell'

/** Read a verified halo before each expansion. Unloaded viewport cells are
 * never treated as evidence that the source table ends. The expansion itself
 * is the exact function used by the desktop smart filter command. */
export async function smartFilterRange(
  selection: GridRange,
  read: (range: GridRange) => Promise<{ cells: ViewCell[]; merges?: GridRange[] }>,
  current: () => boolean,
): Promise<GridRange> {
  let range = { ...selection }
  const single = rangeCellCount(range) === 1
  if (!single && range.startRow !== range.endRow) return range
  const directions = single ? { up: true, down: true, left: true, right: true } : { down: true }
  const matrix = new ObjectMatrix<ICellData>()
  const merges = new Map<string, GridRange>()
  let covered: GridRange | undefined
  for (;;) {
    if (!current()) throw new Error('The document changed while finding the filter range.')
    if (rangeCellCount(range) > 2000)
      throw new Error('Filter supports ranges of at most 2000 cells.')
    const halo = {
      startRow: single ? Math.max(0, range.startRow - 8) : range.startRow,
      endRow: Math.min(1048575, range.endRow + 8),
      startColumn: single ? Math.max(0, range.startColumn - 8) : range.startColumn,
      endColumn: single ? Math.min(16383, range.endColumn + 8) : range.endColumn,
    }
    for (let column = halo.startColumn; column <= halo.endColumn; column += 40) {
      const endColumn = Math.min(halo.endColumn, column + 39)
      const step = Math.max(1, Math.floor(2000 / (endColumn - column + 1)))
      for (let row = halo.startRow; row <= halo.endRow; row += step) {
        const part = {
          startRow: row,
          endRow: Math.min(halo.endRow, row + step - 1),
          startColumn: column,
          endColumn,
        }
        if (
          covered &&
          part.startRow >= covered.startRow &&
          part.endRow <= covered.endRow &&
          part.startColumn >= covered.startColumn &&
          part.endColumn <= covered.endColumn
        )
          continue
        const source = await read(part)
        if (!current()) throw new Error('The document changed while finding the filter range.')
        for (const cell of source.cells) {
          const point = parseAddress(cell.address)
          matrix.setValue(point.row, point.column, gridCell(cell))
        }
        for (const merge of source.merges ?? []) {
          if (
            merge.startRow > halo.endRow ||
            merge.endRow < halo.startRow ||
            merge.startColumn > halo.endColumn ||
            merge.endColumn < halo.startColumn
          )
            continue
          if (rangeCellCount(merge) > 2000)
            throw new Error('The filter touches an oversized merged range.')
          merges.set(
            `${formatAddress(merge.startRow, merge.startColumn)}:${formatAddress(merge.endRow, merge.endColumn)}`,
            merge,
          )
        }
      }
    }
    covered = halo
    const mergeData = [...merges.values()]
    if (mergeData.reduce((count, merge) => count + rangeCellCount(merge), 0) > 100_000)
      throw new Error('The filter contains too many merged cells.')
    const worksheet = {
      getMaxRows: () => 1048576,
      getMaxColumns: () => 16384,
      getCellMatrix: () => matrix,
      getSpanModel: () => ({ getMergedCellRange: () => mergeData }),
      getSnapshot: () => ({ mergeData }),
    } as unknown as Parameters<typeof expandToContinuousRange>[2]
    const next = expandToContinuousRange(range, directions, worksheet)
    if (
      next.startRow === range.startRow &&
      next.endRow === range.endRow &&
      next.startColumn === range.startColumn &&
      next.endColumn === range.endColumn
    )
      return range
    range = next
  }
}
