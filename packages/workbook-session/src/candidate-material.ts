import {
  filePivotsOfSheet,
  pivotSourceSheet,
  sessionPivotAdditions,
} from '@genoffice/xlsx-gateway/domain/session-candidate'
import { chartSelectionReferences } from '@genoffice/xlsx-gateway/domain/chart-series-selection'
import {
  applyChartStateEdit,
  type ChartStateEdit,
} from '@genoffice/xlsx-gateway/domain/chart-visual'
import { resolveChartRange } from '@genoffice/xlsx-gateway/domain/chart-series-range'
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
  'subtotal',
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
    full = input.operations.some((op) => structural.has(op.op) || op.op === 'relocate_range')
  const pageSettingsOnly = input.operations.every((op) =>
    [
      'set_page_setup',
      'add_defined_name',
      'delete_defined_name',
      'manage_defined_name',
      'add_defined_names',
    ].includes(op.op),
  )
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
      if ((op.op === 'set_range' || op.op === 'paste_cells') && op.start) {
        const values = op.op === 'paste_cells' ? op.cells : op.values
        const start = parseRange(op.start)
        addRange(
          op.sheetId,
          `${op.start}:${formatAddress(start.startRow + values.length - 1, start.startColumn + values[0]!.length - 1)}`,
        )
      }
      if (op.op === 'add_chart') addRange(op.sheetId, op.dataRange)
      if (op.op === 'add_pivot') addRange(op.sheetId, op.sourceRange)
      if (op.op === 'edit_pivot') {
        const sheet = state.sheets.find((item) => item.id === op.sheetId)
        if (sheet) {
          if (Math.max(sheet.rows, 1) * Math.max(sheet.columns, 1) > LIMITS.rangeCells)
            throw new WorkbookError(
              'structural_material_budget',
              'PivotTable edits need more than 200,000 cells of sheet material.',
            )
          addRange(
            sheet.id,
            `A1:${formatAddress(Math.max(0, sheet.rows - 1), Math.max(0, sheet.columns - 1))}`,
          )
          const area = (range: {
            startRow: number
            startColumn: number
            endRow: number
            endColumn: number
          }) =>
            `${formatAddress(range.startRow, range.startColumn)}:${formatAddress(range.endRow, range.endColumn)}`
          if ('session' in op.pivot) {
            const name = op.pivot.session
            const add = sessionPivotAdditions(state).find(
              (pivot) => pivot.sheetId === sheet.id && pivot.name === name,
            )
            if (add) addRange(add.sourceSheetId, area(add.sourceArea))
          } else {
            const path = op.pivot.file
            const relayout = (
              (sheet.features?.pivotRelayouts ?? {}) as Record<
                string,
                { sourceSheetId: string; sourceArea: Parameters<typeof area>[0] }
              >
            )[path]
            const definition = filePivotsOfSheet(sheet).find(
              (pivot) => pivot.path === path,
            )?.definition
            const source = definition && pivotSourceSheet(state, definition.sourceSheet)
            if (relayout) addRange(relayout.sourceSheetId, area(relayout.sourceArea))
            else if (source && definition) addRange(source.id, definition.sourceRef)
          }
        }
      }
      if (op.op === 'refresh_pivot') {
        // Growth checks need the whole output sheet; recompute reads each source range.
        const sheet = state.sheets.find((item) => item.id === op.sheetId)
        if (sheet) {
          if (Math.max(sheet.rows, 1) * Math.max(sheet.columns, 1) > LIMITS.rangeCells)
            throw new WorkbookError(
              'structural_material_budget',
              'PivotTable refresh needs more than 200,000 cells of sheet material.',
            )
          addRange(
            sheet.id,
            `A1:${formatAddress(Math.max(0, sheet.rows - 1), Math.max(0, sheet.columns - 1))}`,
          )
          for (const pivot of filePivotsOfSheet(sheet)) {
            const source = pivot.definition && pivotSourceSheet(state, pivot.definition.sourceSheet)
            if (source && pivot.definition) addRange(source.id, pivot.definition.sourceRef)
          }
        }
      }
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
  // Print metadata does not consume or recalculate chart data.
  // Hydrate exact chart references, not unrelated source worksheets.
  for (const sheet of pageSettingsOnly ? [] : state.sheets) {
    for (const visual of (sheet.features?.charts ?? []) as unknown as SheetVisual[])
      for (const series of visual.kind === 'chart' ? visual.chart.series : [])
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
    if (op.op === 'edit_chart' && (op.seriesData || op.seriesSelection)) {
      const native = (
        state.features?.originalMetadata as { visuals?: SheetVisual[] } | undefined
      )?.visuals?.find((visual) => visual.chartPath === op.chartPath)
      const owner = state.sheets.find(
        (sheet) =>
          ((sheet.features?.charts ?? []) as unknown as SheetVisual[]).some(
            (visual) => visual.id === op.chartPath,
          ) ||
          (!!native && sheet.originalSheetId === native.sheetId),
      )
      if (!owner) throw new WorkbookError('unsupported_operation', 'unknown_chart')
      if (op.seriesSelection) {
        const own = ((owner.features?.charts ?? []) as unknown as SheetVisual[]).find(
          (v) => v.id === op.chartPath,
        )
        const visual = own ?? native
        if (!visual || visual.kind !== 'chart')
          throw new WorkbookError('unsupported_operation', 'unknown_chart')
        const overlays = (owner.features?.chartEdits ?? {}) as Record<string, ChartStateEdit>
        const chart = applyChartStateEdit(visual.chart, own ? undefined : overlays[op.chartPath])
        for (const reference of chartSelectionReferences(chart, op.seriesSelection)) {
          const resolved = resolveChartRange(reference, state.sheets, owner.id)
          addRange(resolved.sheetId, resolved.range)
        }
      }
      for (const series of op.seriesData ?? []) {
        for (const range of [series.valuesRange, series.categoriesRange]) {
          if (!range) continue
          const resolved = resolveChartRange(range, state.sheets, series.sheetId ?? owner.id)
          addRange(resolved.sheetId, resolved.range)
        }
      }
    }
  // Journaled original-series references remain live after subsequent cell edits.
  for (const sheet of pageSettingsOnly ? [] : state.sheets) {
    const edits = (sheet.features?.chartEdits ?? {}) as Record<
      string,
      {
        series?: { valuesRef?: string; categoriesRef?: string }[]
        seriesSet?: { valuesRef?: string; categoriesRef?: string }[]
      }
    >
    for (const edit of Object.values(edits))
      for (const series of [...(edit.series ?? []), ...(edit.seriesSet ?? [])])
        for (const reference of [series.valuesRef, series.categoriesRef]) {
          if (!reference) continue
          const resolved = resolveChartRange(reference, state.sheets, sheet.id)
          addRange(resolved.sheetId, resolved.range)
        }
  }
  const hiddenRows: Record<string, number[]> = {}
  const completeSheetIds: string[] = [],
    writableSheetIds: string[] = [],
    viewWritableSheetIds: string[] = []
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
          [
            'set_filter_criteria',
            'clear_filter',
            'clear_filter_criteria',
            'reapply_filter',
          ].includes(op.op),
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
    const needsRowVisibility = input.operations.some(
      (op) =>
        'sheetId' in op &&
        op.sheetId === sheet.id &&
        [
          'sort_range',
          'set_filter',
          'set_filter_criteria',
          'clear_filter',
          'clear_filter_criteria',
          'reapply_filter',
        ].includes(op.op),
    )
    if (needsRowVisibility) hiddenRows[sheet.id] = []
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
        includeMerges: needsRowVisibility,
      })
      if (needsRowVisibility && 'rows' in read)
        for (const row of read.rows) if (row.hidden) hiddenRows[sheet.id]!.push(row.row)
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
      if (!sheet.readOnly) viewWritableSheetIds.push(sheet.id)
      completeSheetIds.push(sheet.id)
      continue
    }
    const protection = features!.sheetProtection
    sheet.protected =
      protection !== null && protection !== undefined && protection.protected !== false
    if (!sheet.readOnly && !sheet.protected) writableSheetIds.push(sheet.id)
    if (!sheet.readOnly) viewWritableSheetIds.push(sheet.id)
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
    hiddenRows,
    referenceRanges: Object.fromEntries([...ranges].map(([id, selected]) => [id, [...selected]])),
    writableSheetIds,
    viewWritableSheetIds,
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
