import type { SaveCalculationValue } from './calculation/save-values'
import { sheetProtectionOf } from '@genoffice/xlsx-gateway/domain/session-candidate'
import {
  projectDataValidationRules,
  type DataValidationEdits,
} from '@genoffice/xlsx-gateway/domain/data-validation-projection'
import { workbookFilterStateSchema } from '@genoffice/xlsx-gateway/shared/workbook-protocol'
import {
  visibilityEditsSchema,
  axisHiddenOverride,
} from '@genoffice/xlsx-gateway/domain/session-visibility'
import {
  rewriteNameReferences,
  type NameReferenceChange,
} from '@genoffice/xlsx-gateway/domain/defined-name-references'
import { shiftedSessionFormula } from '@genoffice/xlsx-gateway/domain/session-formula'
import { z } from 'zod'
import * as numfmt from 'numfmt'
import {
  columnIndex,
  formatAddress,
  parseRange,
  rangeCellCount,
} from '@genoffice/xlsx-gateway/domain/cell-address'
import { MAX_GRID_COLUMNS, MAX_GRID_ROWS } from '@genoffice/xlsx-gateway/shared/grid-bounds'
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
import {
  originalAxisIndex,
  currentAxisIndex,
} from '@genoffice/xlsx-gateway/domain/session-coordinates'

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
  /**
   * Fresh formula results of this root, requested only when the range holds a formula whose
   * cached value is not current (tools and search want results, the editor view has its own).
   */
  calculate?: () => Promise<ReadonlyMap<string, SaveCalculationValue['value']> | null>
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
                formula: rewriteNameReferences(
                  shiftedSessionFormula(state, sheet.id, cell.formula),
                  sheet.id,
                  {
                    changes: (state.features?.definedNameReferences ??
                      []) as unknown as NameReferenceChange[],
                    sheets: state.sheets.map(({ id, name }) => ({ id, name })),
                  },
                ),
                ...(cell.value === null ? {} : { cachedValue: cell.value }),
                calculationState:
                  cell.value === null ||
                  state.structuralLog.length ||
                  !!state.features?.definedNamesEdits ||
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
  const recordOf = ({ address, original }: (typeof points)[number]) =>
    sheet.cells[address] ??
    (original ? base.get(formatAddress(original.row, original.column)) : undefined) ??
    blank()
  const pending = points.some((point) => {
    const record = recordOf(point)
    return !!record.formula && record.calculationState !== 'ready'
  })
  const calculated = pending && input.calculate ? await input.calculate() : null
  assertNotCancelled(input.signal)
  const cells = points.map((point) => {
    const { address } = point
    let record = recordOf(point)
    const fresh = record.formula ? calculated?.get(`${sheet.id}!${address}`) : undefined
    if (record.formula && record.calculationState !== 'ready' && fresh !== undefined) {
      const error = typeof fresh === 'object' && fresh !== null
      const value = error ? fresh.error : fresh
      record = {
        ...record,
        kind: error
          ? 'error'
          : value === null
            ? 'blank'
            : typeof value === 'number'
              ? record.kind === 'date'
                ? 'date'
                : 'number'
              : typeof value === 'boolean'
                ? 'boolean'
                : 'string',
        rawValue: value,
        cachedValue: value,
        calculationState: error ? 'error' : 'ready',
      }
    }
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
  const filterState = sheet.features?.filterEdits
    ? workbookFilterStateSchema.parse(sheet.features.filterEdits)
    : undefined
  const filterHidden = new Set(filterState?.hiddenRows)
  const visibility = visibilityEditsSchema.parse(sheet.features?.visibilityEdits ?? [])
  const rows = []
  for (let row = bounds.startRow; row <= bounds.endRow; row++) {
    const original = originalAxisIndex(state.structuralLog, sheet.id, row, 'row')
    const properties = original === null ? undefined : sourceRows.get(original)
    const filtered =
      filterState &&
      row > filterState.visibilityRange.startRow &&
      row <= filterState.visibilityRange.endRow
        ? filterHidden.has(row)
        : properties?.hidden
    const hidden = axisHiddenOverride(visibility, 'rows', row) ?? filtered
    if (properties || hidden !== undefined)
      rows.push({ row, ...properties, ...(hidden === undefined ? {} : { hidden }) })
  }
  const links = new Map<string, string>()
  const originalLinks = z
    .array(
      z.object({
        row: z.number().int().nonnegative(),
        column: z.number().int().nonnegative(),
        target: z.string(),
      }),
    )
    .parse('hyperlinks' in features ? features.hyperlinks : [])
  for (const link of originalLinks) {
    const row = currentAxisIndex(state.structuralLog, sheet.id, link.row, 'row')
    const column = currentAxisIndex(state.structuralLog, sheet.id, link.column, 'column')
    if (row !== null && column !== null) links.set(formatAddress(row, column), link.target)
  }
  for (const [address, target] of Object.entries(
    (sheet.features?.hyperlinkEdits ?? {}) as Record<string, string | null>,
  )) {
    if (target === null) links.delete(address)
    else links.set(address, target)
  }
  const hyperlinks = [...links]
    .filter(([address]) => {
      const cell = parseRange(address)
      return (
        cell.startRow >= bounds.startRow &&
        cell.endRow <= bounds.endRow &&
        cell.startColumn >= bounds.startColumn &&
        cell.endColumn <= bounds.endColumn
      )
    })
    .map(([address, target]) => ({ address, target }))
  const filter = filterState?.filter
  return {
    ...result,
    autoFilter:
      filter === undefined
        ? 'autoFilter' in features
          ? features.autoFilter
          : null
        : (filter?.range ?? null),
    autoFilterColumns:
      filter === undefined
        ? 'autoFilterColumns' in features
          ? features.autoFilterColumns
          : []
        : (filter?.columns ?? []),
    merges,
    dataValidations: projectDataValidationRules(
      Array.isArray(features.dataValidations)
        ? (features.dataValidations as import('@genoffice/xlsx-gateway/shared/workbook-protocol').WorkbookRangeResult['dataValidations'])
        : [],
      sheet.features?.dataValidationEdits as unknown as DataValidationEdits | undefined,
    ),
    rows,
    hyperlinks,
    // File rules sit at source coordinates; a session rule set replaces them.
    conditionalRules:
      sheet.features?.cfEdits ||
      state.structuralLog.some((op) => 'sheetId' in op && op.sheetId === sheet.id)
        ? []
        : Array.isArray(features.conditionalRules)
          ? features.conditionalRules
          : [],
    sessionConditionalFormats:
      (sheet.features?.cfEdits as { rules: unknown[] } | undefined)?.rules ?? null,
    pageSetup: features.pageSetup ?? null,
    // Effective protection for the Review dialogs; password hashes stay on the host.
    protection: (() => {
      const effective = sheetProtectionOf({ ...sheet.features, ...features })
      if (effective) return { protected: effective.protected, hasPassword: effective.hasPassword }
      // An indexed sheet without a <sheetProtection> element is unprotected.
      return 'sheetProtection' in features ? { protected: false, hasPassword: false } : null
    })(),
    protectedRanges: (
      (sheet.features?.protectedRangeEdits ??
        ('protectedRanges' in features ? features.protectedRanges : [])) as {
        name: string
        sqref: string
        hasPassword?: boolean
      }[]
    ).map(({ name, sqref, hasPassword }) => ({ name, sqref, hasPassword: hasPassword === true })),
    pageBreaks: {
      rowBreaks: 'rowBreaks' in features ? features.rowBreaks : [],
      colBreaks: 'colBreaks' in features ? features.colBreaks : [],
    },
  }
}
