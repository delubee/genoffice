import * as numfmt from 'numfmt'
import type { PivotLayout, PivotScalar } from './pivot-layout'

/** Excel's report-filter caption when no member is selected. */
export const PIVOT_ALL_ITEMS_CAPTION = '(All)'
const DATE_1904_OFFSET = 1462

function isCalendarDatePattern(pattern: string): boolean {
  try {
    const type = (numfmt.getFormatInfo(pattern) as { type?: string }).type
    return type === 'date' || type === 'datetime'
  } catch {
    return false
  }
}

export interface PivotSourceCell {
  readonly value: string | number | boolean | null
  readonly pattern?: string | undefined
  /** A literal (not formula) serial in a 1904 date-system workbook. */
  readonly static1904?: boolean | undefined
}

/**
 * Width-independent pivot source values: numbers are labelled through their
 * own pattern (dates group by their displayed text), except measure columns,
 * which stay numeric. Same rule as the desktop's readPivotSourceGrid.
 */
export function pivotSourceGrid(
  cells: readonly (readonly PivotSourceCell[])[],
  valueFields: readonly string[] = [],
): PivotScalar[][] {
  const valueColumns = new Set(
    (cells[0] ?? []).flatMap((header, index) =>
      valueFields.includes(String(header.value ?? '').trim()) ? [index] : [],
    ),
  )
  return cells.map((row, rowOffset) =>
    row.map(({ value, pattern, static1904 }, columnOffset) => {
      if (typeof value !== 'number') return value as PivotScalar
      if (rowOffset > 0 && valueColumns.has(columnOffset)) return value
      if (!pattern || pattern === 'General') return value
      const shift = static1904 && isCalendarDatePattern(pattern) ? DATE_1904_OFFSET : 0
      try {
        return numfmt.format(pattern, value + shift, { nbsp: true, throws: false })
      } catch {
        return value
      }
    }),
  )
}

/** Rows the report-filter block occupies above the table body. */
export function pivotPageRows(pageFieldCount: number): number {
  return pageFieldCount > 0 ? pageFieldCount + 1 : 0
}

/** The written block: "Field | (All)" report-filter rows, a separator, then the body. */
export function pivotBlockMatrix(layout: PivotLayout): (string | number | null)[][] {
  const { definition, width } = layout
  const pages = (definition.pageFieldIndices ?? []).map((fieldIndex, page) => {
    const item = definition.pageItems?.[page] ?? null
    const line: (string | number | null)[] = new Array(width).fill(null)
    line[0] = definition.fieldNames[fieldIndex] ?? ''
    line[1] =
      item === null ? PIVOT_ALL_ITEMS_CAPTION : (definition.pageLevelItems?.[page]?.[item] ?? '')
    return line
  })
  if (pages.length) pages.push(new Array(width).fill(null))
  return [...pages, ...layout.matrix.map((line) => [...line])]
}

/** Excel's "numeric field": every non-blank value is a number and the column is not date-formatted. */
export function isNumericPivotColumn(cells: readonly PivotSourceCell[]): boolean {
  let seen = false
  for (const { value, pattern } of cells) {
    if (value === null || value === '') continue
    if (typeof value !== 'number') return false
    if (pattern && isCalendarDatePattern(pattern)) return false
    seen = true
  }
  return seen
}
