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

type PivotDefinitionShape = import('../gateway/xlsx-pivot').PivotDefinition
type Area = { startRow: number; startColumn: number; endRow: number; endColumn: number }

export function pivotMemberCaption(
  definition: PivotDefinitionShape,
  field: number,
  member: number,
): string {
  const item = definition.fieldItems[field]?.[member]
  if (!item || item.x === null) return ''
  return String(definition.fields[field]?.sharedItems[item.x] ?? '')
}

export function pivotColLineLabel(
  definition: PivotDefinitionShape,
  line: PivotDefinitionShape['colLines'][number],
  subtotalLabel = 'Subtotal',
): string {
  if (line.t === 'grand') return 'Grand Total'
  if (line.t === 'blank') return ''
  const levels = line.t === 'data' ? definition.colFields.length : line.depth
  const parts: string[] = []
  for (let level = 0; level < levels; level += 1) {
    const field = definition.colFields[level]
    const member = line.members[level]
    if (field === undefined || member === null || member === undefined) continue
    parts.push(
      field === -2
        ? (definition.dataFields[member]?.name ?? '')
        : pivotMemberCaption(definition, field, member),
    )
  }
  if (line.t === 'default') parts.push(subtotalLabel)
  // Without column dimensions, the single data column's caption is the
  // data-field name.
  if (parts.length === 0) return definition.dataFields[line.dataField]?.name ?? ''
  return parts.join(' ')
}

/** Refresh growth outside the compact single-header layout the writer can widen. */
export class PivotGrowError extends Error {}

/**
 * The whole rewritten output after refresh folded new members into the layout:
 * header row, row labels in the baked form, and the union with the old area so
 * a shrunken layout clears its leftovers. `growth` lists newly occupied areas,
 * which must be empty before writing.
 */
export function grownPivotBlock(
  grown: PivotDefinitionShape,
  data: readonly (readonly (number | null)[])[],
  oldBounds: Area,
  subtotalLabel = 'Subtotal',
) {
  const labelCols = Math.max(1, grown.rowFields.length)
  if (grown.firstDataRow !== 1 || grown.firstDataCol !== labelCols) throw new PivotGrowError()
  const width = labelCols + (data[0]?.length ?? 0)
  const height = 1 + data.length
  const newEndRow = oldBounds.startRow + height - 1
  const newEndColumn = oldBounds.startColumn + width - 1
  const growth: Area[] = []
  if (newEndRow > oldBounds.endRow)
    growth.push({
      startRow: oldBounds.endRow + 1,
      endRow: newEndRow,
      startColumn: oldBounds.startColumn,
      endColumn: newEndColumn,
    })
  const sideEnd = Math.min(oldBounds.endRow, newEndRow)
  if (newEndColumn > oldBounds.endColumn && sideEnd >= oldBounds.startRow)
    growth.push({
      startRow: oldBounds.startRow,
      endRow: sideEnd,
      startColumn: oldBounds.endColumn + 1,
      endColumn: newEndColumn,
    })
  // Header row: row-field names + each column line's caption.
  const headerRow: (string | number | null)[] = [
    ...(grown.rowFields.length === 0
      ? ['']
      : grown.rowFields.map((field) =>
          field === -2 ? 'Values' : (grown.fields[field]?.name ?? ''),
        )),
    ...grown.colLines.map((line) => pivotColLineLabel(grown, line, subtotalLabel)),
  ]
  // Row labels: data rows match the baked form (prefix members equal to the
  // previous data row stay blank), subtotal rows write the subtotal label at
  // their depth, and the grand-total row writes "Grand Total" in column 1.
  const block: (string | number | null)[][] = [headerRow]
  let previousMembers: readonly (number | null)[] | null = null
  grown.rowLines.forEach((line, index) => {
    const labels: (string | null)[] = new Array(labelCols).fill(null)
    if (line.t === 'grand') {
      labels[0] = 'Grand Total'
      previousMembers = null
    } else if (line.t === 'default') {
      for (let level = 0; level < line.depth && level < labelCols; level += 1) {
        const field = grown.rowFields[level]
        const member = line.members[level]
        if (field === undefined || member === null || member === undefined) continue
        labels[level] =
          field === -2
            ? (grown.dataFields[member]?.name ?? '')
            : pivotMemberCaption(grown, field, member)
      }
      if (line.depth < labelCols) labels[line.depth] = subtotalLabel
    } else if (line.t === 'data') {
      for (let level = 0; level < labelCols; level += 1) {
        const field = grown.rowFields[level]
        const member = line.members[level]
        if (field === undefined || member === null || member === undefined) continue
        if (previousMembers !== null && level < labelCols - 1) {
          let samePrefix = true
          for (let k = 0; k <= level; k += 1) {
            if (previousMembers[k] !== line.members[k]) {
              samePrefix = false
              break
            }
          }
          if (samePrefix) continue
        }
        labels[level] =
          field === -2
            ? (grown.dataFields[member]?.name ?? '')
            : pivotMemberCaption(grown, field, member)
      }
      previousMembers = line.members
    }
    block.push([...labels, ...(data[index] ?? [])])
  })
  const totalRows = Math.max(height, oldBounds.endRow - oldBounds.startRow + 1)
  const totalColumns = Math.max(width, oldBounds.endColumn - oldBounds.startColumn + 1)
  const padded = Array.from({ length: totalRows }, (_, rowIndex) => {
    const row = block[rowIndex] ?? []
    return Array.from({ length: totalColumns }, (_, colIndex) => row[colIndex] ?? null)
  })
  const label = (column: number) => {
    let text = ''
    for (let value = column + 1; value > 0; value = Math.floor((value - 1) / 26))
      text = String.fromCharCode(65 + ((value - 1) % 26)) + text
    return text
  }
  return {
    block: padded,
    growth,
    outputRef: `${label(oldBounds.startColumn)}${oldBounds.startRow + 1}:${label(newEndColumn)}${newEndRow + 1}`,
  }
}
