import { cellRecordSchema } from '@genoffice/xlsx-gateway/shared/cell-record'
import {
  workbookOperationSchema,
  type WorkbookOperation,
} from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { formatAddress, parseRange } from '@genoffice/xlsx-gateway/domain/cell-address'
import { offsetFormulaRefs } from '@genoffice/xlsx-gateway/domain/formula-shift'
import { t } from '../../../apps/sheets/src/renderer/i18n/locale'
import type { GridRange } from './ports'
import type { ViewCell } from './grid-cell'
import type { z } from 'zod'

type Cell = z.infer<typeof cellRecordSchema>
export interface CopiedBlock {
  id: string
  rows: number[]
  columns: number[]
  cells: Cell[][]
  merges: GridRange[]
}
const intersects = (a: GridRange, b: GridRange) =>
  a.startRow <= b.endRow &&
  a.endRow >= b.startRow &&
  a.startColumn <= b.endColumn &&
  a.endColumn >= b.startColumn
const contains = (a: GridRange, b: GridRange) =>
  a.startRow <= b.startRow &&
  a.endRow >= b.endRow &&
  a.startColumn <= b.startColumn &&
  a.endColumn >= b.endColumn

export function createCopiedBlock(
  id: string,
  bounds: GridRange,
  read: { cells: ViewCell[]; merges?: GridRange[] },
  rows: number[],
  columns: number[],
): CopiedBlock {
  if (!rows.length || !columns.length || rows.length * columns.length > 2000)
    throw new Error(t('appClipboardUnsupported'))
  if (
    rows.some(
      (r, i) =>
        !Number.isInteger(r) ||
        r < bounds.startRow ||
        r > bounds.endRow ||
        (i > 0 && r <= rows[i - 1]!),
    ) ||
    columns.some(
      (c, i) =>
        !Number.isInteger(c) ||
        c < bounds.startColumn ||
        c > bounds.endColumn ||
        (i > 0 && c <= columns[i - 1]!),
    )
  )
    throw new Error(t('appClipboardUnsupported'))
  const indexed = new Map(read.cells.map((cell) => [cell.address, cell]))
  const sourceMerges = (read.merges ?? []).filter((m) => intersects(m, bounds))
  for (const m of sourceMerges) {
    if (!contains(bounds, m)) throw new Error(t('appClipboardUnsupported'))
    for (let r = m.startRow; r <= m.endRow; r++)
      if (!rows.includes(r)) throw new Error(t('appClipboardUnsupported'))
    for (let c = m.startColumn; c <= m.endColumn; c++)
      if (!columns.includes(c)) throw new Error(t('appClipboardUnsupported'))
  }
  return {
    id,
    rows: [...rows],
    columns: [...columns],
    cells: rows.map((row) =>
      columns.map((column) => {
        const cell = indexed.get(formatAddress(row, column))
        if (!cell) throw new Error(t('appClipboardUnsupported'))
        const { address: _address, displayValue: _display, ...data } = cell
        const value = cellRecordSchema.parse(data)
        if (
          sourceMerges.some(
            (m) =>
              row >= m.startRow &&
              row <= m.endRow &&
              column >= m.startColumn &&
              column <= m.endColumn &&
              (row !== m.startRow || column !== m.startColumn),
          )
        )
          return {
            kind: 'blank' as const,
            rawValue: null,
            ...(value.style ? { style: value.style } : {}),
            ...(value.numberFormat ? { numberFormat: value.numberFormat } : {}),
          }
        return value
      }),
    ),
    merges: sourceMerges.map((m) => ({
      startRow: rows.indexOf(m.startRow),
      endRow: rows.indexOf(m.endRow),
      startColumn: columns.indexOf(m.startColumn),
      endColumn: columns.indexOf(m.endColumn),
    })),
  }
}

export function copiedBlockOperation(
  block: CopiedBlock,
  sheetId: string,
  address: string,
  targetMerges: readonly GridRange[] = [],
): WorkbookOperation {
  const selection = parseRange(address),
    height = block.cells.length,
    width = block.cells[0]!.length
  const repeat =
    (selection.endRow - selection.startRow + 1) % height === 0 &&
    (selection.endColumn - selection.startColumn + 1) % width === 0 &&
    !targetMerges.some((m) => intersects(m, selection))
  const rows = repeat ? selection.endRow - selection.startRow + 1 : height,
    columns = repeat ? selection.endColumn - selection.startColumn + 1 : width
  if (
    rows * columns > 2000 ||
    selection.startRow + rows > 1048576 ||
    selection.startColumn + columns > 16384
  )
    throw new Error(t('appClipboardUnsupported'))
  const cells = Array.from({ length: rows }, (_, r) =>
    Array.from({ length: columns }, (_, c) => {
      const value = structuredClone(block.cells[r % height]![c % width]!)
      if (value.formula)
        value.formula = offsetFormulaRefs(
          value.formula,
          selection.startRow + r - block.rows[r % height]!,
          selection.startColumn + c - block.columns[c % width]!,
        )
      return value
    }),
  )
  const merges: string[] = []
  for (let r = 0; r < rows; r += height)
    for (let c = 0; c < columns; c += width)
      for (const m of block.merges)
        merges.push(
          `${formatAddress(selection.startRow + r + m.startRow, selection.startColumn + c + m.startColumn)}:${formatAddress(selection.startRow + r + m.endRow, selection.startColumn + c + m.endColumn)}`,
        )
  return workbookOperationSchema.parse({
    op: 'paste_cells',
    sheetId,
    start: formatAddress(selection.startRow, selection.startColumn),
    cells,
    merges,
  })
}

/** Native copy HTML keeps this opaque identity even after a view reinstall. */
export function clipboardCopyId(html: string): string | null {
  return /\bdata-copy-id=["']([^"']+)["']/.exec(html)?.[1] ?? null
}
