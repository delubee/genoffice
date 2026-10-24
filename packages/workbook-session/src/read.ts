import { z } from 'zod'
import * as numfmt from 'numfmt'
import {
  columnIndex,
  formatAddress,
  parseRange,
  rangeCellCount,
} from '@genoffice/xlsx-gateway/domain/cell-address'
import { MAX_GRID_COLUMNS, MAX_GRID_ROWS } from '@genoffice/xlsx-gateway/shared/grid-bounds'
import { renameSheetInFormula } from '@genoffice/xlsx-gateway/gateway/xlsx-sheets'
import { shiftFormulaText } from '@genoffice/xlsx-gateway/gateway/xlsx-structure'
import { workbookStyleEditSchema } from '@genoffice/xlsx-gateway/shared/edit-schemas'
import {
  LIMITS,
  assertNotCancelled,
  WorkbookError,
  type CellRecord,
  type WorkbookState,
} from './domain'
import type { WorkbookSource } from './ports'
import { logicalSheet } from './sheet-map'
import { indexedFeatures } from './source-features'
import { workbookMergeRangesResultSchema } from '@genoffice/xlsx-gateway/shared/workbook-protocol'
import { originalAxisIndex } from '@genoffice/xlsx-gateway/domain/session-coordinates'

const sourceCells = z.object({
  rows: z
    .array(
      z.object({
        row: z.number().int().nonnegative(),
        height: z.number().nonnegative().optional(),
        hidden: z.boolean().optional(),
        customHeight: z.boolean().optional(),
      }),
    )
    .default([]),
  cells: z.array(
    z.object({
      row: z.number().int().nonnegative(),
      column: z.number().int().nonnegative(),
      value: z.union([z.string(), z.number().finite(), z.boolean(), z.null()]),
      isError: z.boolean().optional(),
      formula: z.string().optional(),
      styleIndex: z.number().int().nonnegative().optional(),
    }),
  ),
})
const sourceFeatures = z.object({
  styles: z.array(z.object({ numberFormat: z.string().optional() }).passthrough()).default([]),
})
const blank = (): CellRecord => ({ kind: 'blank', rawValue: null })

export function baseCoordinate(
  state: Readonly<WorkbookState>,
  sheetId: string,
  row: number,
  column: number,
): { row: number; column: number } | null {
  for (const op of [...state.structuralLog].reverse()) {
    if (!('sheetId' in op) || op.sheetId !== sheetId) continue
    const rows = op.op === 'insert_rows' || op.op === 'delete_rows'
    const columns = op.op === 'insert_cols' || op.op === 'delete_cols'
    if (!rows && !columns) continue
    const index = 'row' in op ? op.row - 1 : 'column' in op ? columnIndex(op.column) : 0
    let coordinate = rows ? row : column
    const count = 'count' in op ? op.count : 0
    if (op.op.startsWith('insert_')) {
      if (coordinate >= index && coordinate < index + count) return null
      if (coordinate >= index + count) coordinate -= count
    } else if (coordinate >= index) coordinate += count
    if (rows) row = coordinate
    else column = coordinate
  }
  return { row, column }
}

function shiftedFormula(state: Readonly<WorkbookState>, sheetId: string, formula: string) {
  const names = new Map(
    state.sheets.map((sheet) => [
      sheet.id,
      (sheet.features?.originalMetadata as { name?: string } | undefined)?.name ?? sheet.name,
    ]),
  )
  for (const op of state.structuralLog) {
    if (!('sheetId' in op)) continue
    const target = state.sheets.find((sheet) => sheet.id === op.sheetId)
    if (!target) continue
    if (op.op === 'rename_sheet') {
      formula = renameSheetInFormula(formula, names.get(op.sheetId)!, op.name)
      names.set(op.sheetId, op.name)
      continue
    }
    const rows = op.op === 'insert_rows' || op.op === 'delete_rows'
    const columns = op.op === 'insert_cols' || op.op === 'delete_cols'
    if (!rows && !columns) continue
    const index = 'row' in op ? op.row - 1 : 'column' in op ? columnIndex(op.column) : 0
    const count = 'count' in op ? op.count : 0,
      deleting = op.op.startsWith('delete_')
    formula = shiftFormulaText(
      formula,
      names.get(op.sheetId)!,
      {
        boundary: deleting ? index + count : index,
        delta: deleting ? -count : count,
        deleted: deleting ? { start: index, end: index + count - 1 } : null,
      },
      rows ? 'row' : 'column',
      sheetId !== op.sheetId,
    )
  }
  return formula
}

/** Holds the supplied immutable root throughout all asynchronous source reads. */
export async function readCurrentRange(input: {
  state: Readonly<WorkbookState>
  sessionId: string
  source: WorkbookSource
  logicalSheetId: string
  range: string
  signal?: AbortSignal
  limit?: number
  /** Presentation includes merge geometry and off-page anchor values. */
  includeMerges?: boolean
}) {
  const { state, source } = input,
    sheet = logicalSheet(state, input.logicalSheetId),
    bounds = parseRange(input.range)
  const limit = Math.min(input.limit ?? LIMITS.readCells, LIMITS.rangeCells)
  if (rangeCellCount(bounds) > limit)
    throw new WorkbookError('range_limit', `Read exceeds ${limit} cells; request a smaller page.`)
  if (bounds.endRow >= MAX_GRID_ROWS || bounds.endColumn >= MAX_GRID_COLUMNS)
    throw new WorkbookError('invalid_range', 'Read exceeds the Excel grid.')
  assertNotCancelled(input.signal)
  const points: { address: string; original: { row: number; column: number } | null }[] = []
  for (let row = bounds.startRow; row <= bounds.endRow; row++)
    for (let column = bounds.startColumn; column <= bounds.endColumn; column++)
      points.push({
        address: formatAddress(row, column),
        original: sheet.originalSheetId ? baseCoordinate(state, sheet.id, row, column) : null,
      })
  // The visible grid can extend beyond the original worksheet's used area.
  // Bounds belong to the immutable source, not to today's structural dimensions.
  const originalMetadata = sheet.features?.originalMetadata as
    { rowCount?: number; columnCount?: number } | undefined
  const needed = points.filter(
    (point) =>
      point.original &&
      (input.includeMerges || !Object.hasOwn(sheet.cells, point.address)) &&
      (originalMetadata?.rowCount === undefined ||
        point.original.row < originalMetadata.rowCount) &&
      (originalMetadata?.columnCount === undefined ||
        point.original.column < originalMetadata.columnCount),
  )
  const base = new Map<string, CellRecord>()
  const sourceRows = new Map<
    number,
    { height?: number; hidden?: boolean; customHeight?: boolean }
  >()
  if (needed.length && sheet.originalSheetId) {
    const features = sourceFeatures.parse(
      await source.readFeatures(input.sessionId, sheet.originalSheetId),
    )
    // Ordinary pages need one source read. Discontinuous mappings use bounded
    // row groups rather than expanding a page across deleted million-row gaps.
    const rowGroups = new Map<number, { min: number; max: number }>()
    for (const point of needed) {
      const { row, column } = point.original!
      const existing = rowGroups.get(row)
      rowGroups.set(row, {
        min: Math.min(existing?.min ?? column, column),
        max: Math.max(existing?.max ?? column, column),
      })
    }
    const rows = [...rowGroups.keys()],
      columns = [...rowGroups.values()]
    const minRow = Math.min(...rows),
      maxRow = Math.max(...rows),
      minCol = Math.min(...columns.map((item) => item.min)),
      maxCol = Math.max(...columns.map((item) => item.max))
    const ranges =
      (maxRow - minRow + 1) * (maxCol - minCol + 1) <= 100_000
        ? [{ startRow: minRow, endRow: maxRow, startColumn: minCol, endColumn: maxCol }]
        : [...rowGroups].map(([row, span]) => ({
            startRow: row,
            endRow: row,
            startColumn: span.min,
            endColumn: span.max,
          }))
    for (const range of ranges) {
      assertNotCancelled(input.signal)
      const response = sourceCells.parse(
        await source.readBaseRange(
          input.sessionId,
          sheet.originalSheetId,
          `${formatAddress(range.startRow, range.startColumn)}:${formatAddress(range.endRow, range.endColumn)}`,
        ),
      )
      for (const row of response.rows)
        sourceRows.set(row.row, {
          ...(row.height === undefined ? {} : { height: row.height }),
          ...(row.hidden === undefined ? {} : { hidden: row.hidden }),
          ...(row.customHeight === undefined ? {} : { customHeight: row.customHeight }),
        })
      for (const cell of response.cells) {
        const sourceStyle = features.styles[cell.styleIndex ?? 0]
        const style = workbookStyleEditSchema
          .partial()
          .parse(
            Object.fromEntries(
              Object.entries(sourceStyle ?? {}).filter(([key]) =>
                Object.hasOwn(workbookStyleEditSchema.shape, key),
              ),
            ),
          )
        const numberFormat = sourceStyle?.numberFormat
        const date =
          typeof cell.value === 'number' && numberFormat && numfmt.isDateFormat(numberFormat)
        const kind: CellRecord['kind'] = cell.isError
          ? 'error'
          : cell.value === null
            ? 'blank'
            : date
              ? 'date'
              : typeof cell.value === 'number'
                ? 'number'
                : typeof cell.value === 'boolean'
                  ? 'boolean'
                  : 'string'
        base.set(formatAddress(cell.row, cell.column), {
          kind,
          rawValue: cell.value,
          ...(Object.keys(style).length ? { style } : {}),
          ...(date ? { dateSystem: state.dateSystem } : {}),
          ...(numberFormat ? { numberFormat } : {}),
          ...(cell.formula
            ? {
                formula: shiftedFormula(state, sheet.id, cell.formula),
                ...(cell.value === null ? {} : { cachedValue: cell.value }),
                calculationState:
                  cell.value === null ||
                  state.structuralLog.length ||
                  state.sheets.some((item) => Object.keys(item.cells).length)
                    ? 'pending'
                    : cell.isError
                      ? 'error'
                      : 'ready',
              }
            : {}),
        })
      }
    }
  }
  assertNotCancelled(input.signal)
  const cells = points.map(({ address, original }) => {
    const record =
      sheet.cells[address] ??
      (original ? base.get(formatAddress(original.row, original.column)) : undefined) ??
      blank()
    let displayValue = record.rawValue === null ? '' : String(record.rawValue)
    if (typeof record.rawValue === 'number' && record.numberFormat)
      displayValue = numfmt.format(
        record.numberFormat,
        record.rawValue + (record.kind === 'date' && record.dateSystem === '1904' ? 1462 : 0),
        { throws: false },
      )
    return { address, ...structuredClone(record), displayValue }
  })
  const result = { stateId: state.stateId, logicalSheetId: sheet.id, range: input.range, cells }
  if (!input.includeMerges) return result
  const features = sheet.originalSheetId
    ? await indexedFeatures(source, input.sessionId, sheet.originalSheetId, input.signal)
    : { indexingComplete: true, merges: [] }
  const { merges } = workbookMergeRangesResultSchema.parse({
    indexingComplete: features.indexingComplete,
    merges: sheet.features?.mergeEdits ?? features.merges ?? [],
  })
  // Coordinate-changing edits on merged sheets are rejected by the candidate
  // interpreter. Source merge coordinates therefore remain current coordinates.
  for (const merge of merges) {
    if (
      merge.startRow > bounds.endRow ||
      merge.endRow < bounds.startRow ||
      merge.startColumn > bounds.endColumn ||
      merge.endColumn < bounds.startColumn
    )
      continue
    if (merge.startRow >= bounds.startRow && merge.startColumn >= bounds.startColumn) continue
    const anchor = await readCurrentRange({
      ...input,
      includeMerges: false,
      range: formatAddress(merge.startRow, merge.startColumn),
    })
    cells.push(...anchor.cells)
  }
  const rows = []
  for (let row = bounds.startRow; row <= bounds.endRow; row++) {
    const original = originalAxisIndex(state.structuralLog, sheet.id, row, 'row')
    const properties = original === null ? undefined : sourceRows.get(original)
    if (properties) rows.push({ row, ...properties })
  }
  return { ...result, merges, rows, pageSetup: features.pageSetup ?? null }
}
