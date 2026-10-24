import { randomUUID } from 'node:crypto'
import {
  copyTargetBounds,
  type WorkbookOperation,
} from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import {
  formatAddress,
  parseRange,
  rangeCellCount,
} from '@genoffice/xlsx-gateway/domain/cell-address'
import { LIMITS, WorkbookError, assertNotCancelled, type WorkbookState } from './domain'
import type { WorkbookSource } from './ports'
import { readCurrentRange } from './read'
import { indexedFeatures } from './source-features'
import { workbookStateSchema } from './schemas'
import type { SheetVisual } from '@genoffice/xlsx-gateway/domain/chart-visual'

const structural = new Set([
  'insert_rows',
  'delete_rows',
  'insert_cols',
  'delete_cols',
  'delete_sheet',
  'rename_sheet',
])
/** Host-only materialization; the browser receives no paths or sidecar handles. */
export async function prepareCandidateMaterial(input: {
  state: Readonly<WorkbookState>
  operations: readonly WorkbookOperation[]
  source: WorkbookSource
  sourceSessionId: string
  signal?: AbortSignal
}) {
  const state = workbookStateSchema.parse(input.state),
    full = input.operations.some((op) => structural.has(op.op))
  const ranges = new Map<string, Set<string>>()
  const addRange = (sheetId: string, range: string) => {
    const selected = ranges.get(sheetId) ?? new Set()
    selected.add(range)
    ranges.set(sheetId, selected)
  }
  if (full)
    for (const sheet of state.sheets) {
      if (Math.max(sheet.rows, 1) * Math.max(sheet.columns, 1) > LIMITS.rangeCells)
        throw new WorkbookError(
          'structural_material_budget',
          'This structural edit needs more than 200,000 cells of reference material; use a smaller workbook or range.',
        )
      addRange(
        sheet.id,
        `A1:${formatAddress(Math.max(0, sheet.rows - 1), Math.max(0, sheet.columns - 1))}`,
      )
    }
  for (const op of input.operations)
    if ('sheetId' in op) {
      if ('address' in op) addRange(op.sheetId, op.address)
      if ('range' in op && typeof op.range === 'string') addRange(op.sheetId, op.range)
      if (op.op === 'set_range' && op.start) {
        const start = parseRange(op.start)
        addRange(
          op.sheetId,
          `${op.start}:${formatAddress(start.startRow + op.values.length - 1, start.startColumn + op.values[0]!.length - 1)}`,
        )
      }
      if (op.op === 'add_chart') addRange(op.sheetId, op.dataRange)
      if ('source' in op && typeof op.source === 'string')
        addRange(
          'sourceSheetId' in op && op.sourceSheetId ? op.sourceSheetId : op.sheetId,
          op.source,
        )
      if (op.op === 'copy_range' || op.op === 'fill_range') {
        const target = op.op === 'copy_range' ? copyTargetBounds(op) : parseRange(op.target)
        addRange(
          op.sheetId,
          `${formatAddress(target.startRow, target.startColumn)}:${formatAddress(target.endRow, target.endColumn)}`,
        )
      }
    }
  // Hydrate exact chart references, not unrelated source worksheets.
  for (const sheet of state.sheets) {
    for (const visual of (sheet.features?.charts ?? []) as unknown as SheetVisual[])
      for (const series of visual.chart.series)
        for (const key of ['valuesRef', 'categoriesRef', 'nameRef'] as const) {
          const reference = series[key]
          if (!reference) continue
          const match = /^(?:'((?:[^']|'')*)'|([^!]+))!(.+)$/.exec(reference)
          const sourceName = match?.[1]?.replace(/''/g, "'") ?? match?.[2]
          const sourceSheet = state.sheets.find((item) => item.name === sourceName)
          if (!sourceSheet || !match?.[3])
            throw new WorkbookError(
              'chart_source_missing',
              'The chart reference is not in the current workbook.',
            )
          addRange(sourceSheet.id, match[3])
        }
  }
  for (const op of input.operations)
    if (op.op === 'edit_chart' && op.seriesData) {
      const owner = state.sheets.find((sheet) =>
        ((sheet.features?.charts ?? []) as unknown as SheetVisual[]).some(
          (visual) => visual.id === op.chartPath,
        ),
      )
      if (!owner)
        throw new WorkbookError('unsupported_operation', 'Original workbook charts are read-only.')
      for (const series of op.seriesData) {
        if (series.valuesRange) addRange(series.sheetId ?? owner.id, series.valuesRange)
        if (series.categoriesRange) addRange(series.sheetId ?? owner.id, series.categoriesRange)
      }
    }
  const completeSheetIds: string[] = [],
    writableSheetIds: string[] = []
  for (const sheet of state.sheets) {
    assertNotCancelled(input.signal)
    const features = sheet.originalSheetId
      ? await indexedFeatures(
          input.source,
          input.sourceSessionId,
          sheet.originalSheetId,
          input.signal,
        )
      : undefined
    const selected = ranges.get(sheet.id) ?? new Set<string>()
    if (
      input.operations.some(
        (op) =>
          'sheetId' in op &&
          op.sheetId === sheet.id &&
          ['set_filter_criteria', 'clear_filter'].includes(op.op),
      )
    ) {
      const edits = sheet.features?.filterEdits as
        | {
            filter?: {
              range: { startRow: number; endRow: number; startColumn: number; endColumn: number }
            }
            visibilityRange?: {
              startRow: number
              endRow: number
              startColumn: number
              endColumn: number
            }
          }
        | undefined
      const range =
        edits?.visibilityRange ??
        edits?.filter?.range ??
        (features?.autoFilter as
          { startRow: number; endRow: number; startColumn: number; endColumn: number } | undefined)
      if (range)
        selected.add(
          `${formatAddress(range.startRow, range.startColumn)}:${formatAddress(range.endRow, range.endColumn)}`,
        )
    }
    for (const range of selected) {
      if (rangeCellCount(parseRange(range)) > LIMITS.rangeCells)
        throw new WorkbookError('range_limit', 'Candidate material exceeds the range budget.')
      const read = await readCurrentRange({
        state: input.state,
        sessionId: input.sourceSessionId,
        source: input.source,
        logicalSheetId: sheet.id,
        range,
        signal: input.signal,
        limit: LIMITS.rangeCells,
      })
      for (const { address, displayValue: _display, ...cell } of read.cells)
        if (
          cell.rawValue !== null ||
          cell.formula ||
          cell.style ||
          Object.hasOwn(sheet.cells, address)
        )
          sheet.cells[address] = cell
    }
    if (!sheet.originalSheetId) {
      writableSheetIds.push(sheet.id)
      completeSheetIds.push(sheet.id)
      continue
    }
    const protection = features!.sheetProtection
    sheet.protected =
      protection !== null && protection !== undefined && protection.protected !== false
    if (!sheet.readOnly && !sheet.protected) writableSheetIds.push(sheet.id)
    const { styles: _styles, sheet: _metadata, ...featureState } = features!
    sheet.features = { ...sheet.features, ...JSON.parse(JSON.stringify(featureState)) }
    if (
      full ||
      selected.has(
        `A1:${formatAddress(Math.max(0, sheet.rows - 1), Math.max(0, sheet.columns - 1))}`,
      )
    )
      completeSheetIds.push(sheet.id)
  }
  assertNotCancelled(input.signal)
  return {
    state: workbookStateSchema.parse(state),
    completeSheetIds,
    referenceRanges: Object.fromEntries([...ranges].map(([id, selected]) => [id, [...selected]])),
    writableSheetIds,
    allocation: {
      stateId: randomUUID(),
      sheetIds: Object.fromEntries(
        input.operations.flatMap((op, index) =>
          op.op === 'add_sheet' ? [[String(index), randomUUID()]] : [],
        ),
      ),
    },
  }
}
