import { isDateFormat } from 'numfmt'
import { InMemoryWorkbookAdapter } from './in-memory-workbook'
import { formatAddress, parseAddress } from './cell-address'
import { shiftIndex, shiftSpecForOp } from './formula-shift'
import type { SheetVisual } from './chart-visual'
import type { CellFormatState, WorkbookSnapshot, WorksheetState } from './workbook.types'
import { isStructuralOp, parseWorkbookCommandBatch } from './workbook-dsl'
import { normalizeStyleColor } from './style-color'
import { renameSheetInFormula } from '../gateway/xlsx-sheets'
import { shiftFormulaText } from '../gateway/xlsx-structure'
import {
  transactionInputSchema,
  workbookStateSchema,
  type cellRecordSchema,
} from '../shared/session-state'
import type { z } from 'zod'
import { refreshChartReferences } from './session-chart-references'
import { stateChangeset, type CandidateState } from '../shared/session-changeset'
import { copySessionRange } from './session-range-copy'
import { SESSION_OPERATION_NAMES, hasMixedSessionOperations } from '../shared/session-operations'
import { assertMergeOperation } from './session-merges'

type TypedCell = z.infer<typeof cellRecordSchema>
export interface CandidateMaterial {
  /** Current values and feature state, resolved against exactly stateId. */
  state: CandidateState
  /** Full cell/formula/feature coverage for structural/reference changes. */
  completeSheetIds: readonly string[]
  /** Exact current ranges hydrated for chart dependencies, separate from full sheets. */
  referenceRanges?: Record<string, string[]>
  /** Host-validated protection/readonly state, independent of renderer focus. */
  writableSheetIds: readonly string[]
  allocation: { stateId: string; sheetIds: Record<string, string> }
}

const whitelist = new Set<string>(SESSION_OPERATION_NAMES)
const changingCoordinates = new Set([
  'insert_rows',
  'delete_rows',
  'insert_cols',
  'delete_cols',
  'delete_sheet',
  'rename_sheet',
])
function cancelled(signal?: AbortSignal) {
  if (signal?.aborted) throw new Error('operation_cancelled')
}

function formatState(cell: TypedCell): CellFormatState {
  const style = cell.style ?? {}
  const keys = [
    'bold',
    'italic',
    'underline',
    'strikethrough',
    'fontFamily',
    'fontSize',
    'fontColor',
    'fillColor',
    'numberFormat',
    'wrapText',
    'indent',
    'textRotation',
  ]
  return JSON.parse(
    JSON.stringify({
      ...Object.fromEntries(Object.entries(style).filter(([key]) => keys.includes(key))),
      ...(style.horizontalAlignment ? { horizontalAlign: style.horizontalAlignment } : {}),
      ...(style.verticalAlignment ? { verticalAlign: style.verticalAlignment } : {}),
      ...(cell.numberFormat ? { numberFormat: cell.numberFormat } : {}),
    }),
  )
}
function typedStyle(
  format: CellFormatState | undefined,
  original?: TypedCell['style'],
): TypedCell['style'] {
  if (!format) return original
  const { horizontalAlign, verticalAlign, border, ...simple } = format
  const result = JSON.parse(
    JSON.stringify({
      ...original,
      ...simple,
      ...(horizontalAlign ? { horizontalAlignment: horizontalAlign } : {}),
      ...(verticalAlign ? { verticalAlignment: verticalAlign } : {}),
    }),
  ) as NonNullable<TypedCell['style']>
  if (border)
    for (const edge of ['top', 'bottom', 'left', 'right'] as const)
      if (border.type === 'all' || border.type === 'none' || border.type === edge)
        result[('border' + edge[0]!.toUpperCase() + edge.slice(1)) as 'borderTop'] =
          border.type === 'none'
            ? null
            : {
                style: 'thin',
                ...(border.color ? { color: normalizeStyleColor(border.color) } : {}),
              }
  return result
}
function snapshot(state: CandidateState): WorkbookSnapshot {
  return {
    revision: 0,
    sheets: state.sheets.map((sheet) => ({
      id: sheet.id,
      name: sheet.name,
      gridRows: Math.max(1, sheet.rows),
      gridColumns: Math.max(1, sheet.columns),
      cells: Object.fromEntries(
        Object.entries(sheet.cells).map(([address, cell]) => [
          address,
          { value: cell.rawValue, ...(cell.formula ? { formula: cell.formula } : {}) },
        ]),
      ),
      styles: Object.fromEntries(
        Object.entries(sheet.cells)
          .filter(([, cell]) => cell.style || cell.numberFormat)
          .map(([address, cell]) => [address, formatState(cell)]),
      ),
      visuals: structuredClone((sheet.features?.charts ?? []) as unknown as SheetVisual[]),
    })),
  }
}

export async function computeCandidateState(
  inputState: Readonly<CandidateState>,
  rawIntent: unknown,
  material: CandidateMaterial,
  options: {
    signal?: AbortSignal
    fault?(point: string): void
    afterStep?(snapshot: WorkbookSnapshot): void | Promise<void>
  } = {},
) {
  const signal = options.signal
  const before = workbookStateSchema.parse(inputState),
    intent = transactionInputSchema.parse(rawIntent)
  cancelled(signal)
  if (before.sheets.some((sheet) => sheet.readOnly)) throw new Error('readonly_workbook')
  if (hasMixedSessionOperations(intent.operations))
    throw new Error('mixed_structural_content: commit structure, read again, then commit content')
  parseWorkbookCommandBatch({
    dslVersion: 1,
    transactionId: intent.transactionId,
    baseRevision: intent.expectedRevision,
    summary: intent.summary,
    operations: intent.operations,
  })
  for (const op of intent.operations)
    if (!whitelist.has(op.op)) throw new Error(`unsupported_operation: ${op.op}`)
  for (const op of intent.operations)
    if (
      (op.op === 'add_chart' || op.op === 'edit_chart') &&
      op.chartType &&
      !['column', 'bar', 'line', 'pie'].includes(op.chartType)
    )
      throw new Error('unsupported_chart_type')
  const current = workbookStateSchema.parse(material.state)
  if (
    current.stateId !== before.stateId ||
    current.sheets.map((sheet) => sheet.id).join() !== before.sheets.map((sheet) => sheet.id).join()
  )
    throw new Error('stale_candidate_material')
  const workbookMetadata = current.features?.originalMetadata as Record<string, unknown> | undefined
  if (
    intent.operations.some((op) => changingCoordinates.has(op.op) || op.op === 'add_sheet') &&
    (workbookMetadata?.workbookProtection as { lockStructure?: boolean } | undefined)?.lockStructure
  )
    throw new Error('protected_workbook_structure')
  if (intent.operations.some((op) => changingCoordinates.has(op.op))) {
    if (current.sheets.some((sheet) => !material.completeSheetIds.includes(sheet.id)))
      throw new Error('incomplete_structural_material')
    if (Array.isArray(workbookMetadata?.definedNames) && workbookMetadata.definedNames.length)
      throw new Error('unsupported_structural_feature: definedNames')
    let formulaCount = 0
    for (const sheet of current.sheets) {
      formulaCount += Object.values(sheet.cells).filter((cell) => cell.formula).length
      const metadata = sheet.features?.originalMetadata as Record<string, unknown> | undefined
      // Features whose references have no implemented transform are rejected
      // before creating a candidate, rather than silently left behind.
      for (const key of ['tables', 'comments', 'pivotRanges', 'pivotTables', 'sparklines'])
        if (Array.isArray(metadata?.[key]) && metadata[key].length)
          throw new Error(`unsupported_structural_feature: ${key}`)
      for (const key of [
        'merges',
        'hyperlinks',
        'conditionalRules',
        'dataValidations',
        'protectedRanges',
      ])
        if (Array.isArray(sheet.features?.[key]) && (sheet.features[key] as unknown[]).length)
          throw new Error(`unsupported_structural_feature: ${key}`)
      if (sheet.features?.nativeVisuals && (sheet.features.nativeVisuals as unknown[]).length)
        throw new Error('unsupported_structural_feature: nativeVisuals')
    }
    if (formulaCount > 50_000) throw new Error('formula_dependency_budget')
  }
  let book = snapshot(current)
  let structuralBaseline = structuredClone(book)
  const overlayKeys = new Map(
    before.sheets.map((sheet) => [sheet.id, new Set(Object.keys(sheet.cells))]),
  )
  const originalKinds = new Map(
    current.sheets.map((sheet) => [sheet.id, structuredClone(sheet.cells)]),
  )
  const typedTemplates = new Map(
    current.sheets.map((sheet) => [sheet.id, structuredClone(sheet.cells)]),
  )
  const dimensions = new Map(
    current.sheets.map((sheet) => [sheet.id, { rows: sheet.rows, columns: sheet.columns }]),
  )
  const log = structuredClone(before.structuralLog)
  const sheetMap = structuredClone(before.sheetMap)
  const affected: string[] = []
  let expandedWrites = 0
  for (const [index, op] of intent.operations.entries()) {
    cancelled(signal)
    const target =
      'sheetId' in op
        ? book.sheets.find((sheet) => sheet.id === op.sheetId)
        : book.sheets.find((sheet) =>
            sheet.visuals?.some((visual) =>
              'chartPath' in op
                ? visual.id === op.chartPath
                : 'visualId' in op
                  ? visual.id === op.visualId
                  : false,
            ),
          )
    const permissions = current.sheets.find((sheet) => sheet.id === target?.id)
    if (target && !material.writableSheetIds.includes(target.id))
      throw new Error('unverified_writable_target')
    if (permissions?.readOnly || permissions?.protected) throw new Error('protected_target')
    assertMergeOperation(current, op)
    if (
      op.op === 'delete_sheet' &&
      book.sheets.some((sheet) =>
        Object.values(sheet.cells).some(
          (cell) =>
            cell.formula &&
            renameSheetInFormula(cell.formula, target!.name, '__removed__') !== cell.formula,
        ),
      )
    )
      throw new Error('referenced_sheet')
    if (isStructuralOp(op) || op.op === 'rename_sheet') {
      const baseline = new InMemoryWorkbookAdapter(structuralBaseline)
      const baselinePlan = baseline.plan({
        dslVersion: 1,
        summary: intent.summary,
        transactionId: `baseline-${index}`,
        baseRevision: structuralBaseline.revision,
        operations: [op],
      })
      baseline.apply(baselinePlan)
      structuralBaseline = baseline.getSnapshot()
      if (op.op === 'rename_sheet' && target)
        for (const sheet of structuralBaseline.sheets)
          for (const cell of Object.values(sheet.cells))
            if (cell.formula)
              (cell as { formula: string }).formula = renameSheetInFormula(
                cell.formula,
                target.name,
                op.name,
              )
    }
    let next: WorkbookSnapshot
    if (op.op === 'copy_range' || op.op === 'fill_range') {
      const copied = copySessionRange(book, op, typedTemplates, signal)
      next = copied.snapshot
      const size = dimensions.get(op.sheetId)!
      size.rows = Math.max(size.rows, copied.target.endRow + 1)
      size.columns = Math.max(size.columns, copied.target.endColumn + 1)
    } else {
      const adapter = new InMemoryWorkbookAdapter(book)
      const plan = adapter.plan({
        dslVersion: 1,
        summary: intent.summary,
        transactionId: `${intent.transactionId}-${index}`,
        baseRevision: book.revision,
        operations: [op],
      })
      expandedWrites += plan.cellChanges.length
      if (expandedWrites > 2000)
        throw new Error('expanded_write_limit: ordinary batches allow 2000 cells')
      if (plan.warnings.length)
        throw new Error(`unsafe_reference_change: ${plan.warnings.join('; ')}`)
      if (
        op.op === 'set_cell' ||
        op.op === 'set_formula' ||
        op.op === 'clear_cell' ||
        op.op === 'set_range' ||
        op.op === 'clear_range'
      )
        for (const change of plan.cellChanges) {
          const template = typedTemplates.get(change.sheetId)?.[change.address]
          if (template)
            template.kind =
              change.after.value === null
                ? 'blank'
                : typeof change.after.value === 'number'
                  ? 'number'
                  : typeof change.after.value === 'boolean'
                    ? 'boolean'
                    : 'string'
        }
      adapter.apply(plan)
      next = adapter.getSnapshot()
    }
    if (op.op === 'add_sheet') {
      const fresh = next.sheets.find(
        (sheet) => !book.sheets.some((previous) => previous.id === sheet.id),
      )!
      const logicalId = material.allocation.sheetIds[String(index)]!
      if (
        !logicalId ||
        sheetMap[logicalId] ||
        current.sheets.some((sheet) => sheet.id === logicalId)
      )
        throw new Error('invalid_sheet_allocation')
      sheetMap[logicalId] = { createdName: op.name, createdAtLogIndex: log.length }
      next = {
        ...next,
        sheets: next.sheets.map((sheet) =>
          sheet.id === fresh.id ? { ...sheet, id: logicalId } : sheet,
        ),
      }
      structuralBaseline = {
        ...structuralBaseline,
        sheets: structuralBaseline.sheets.map((sheet) =>
          sheet.id === fresh.id ? { ...sheet, id: logicalId } : sheet,
        ),
      }
      dimensions.set(logicalId, { rows: fresh.gridRows ?? 1000, columns: fresh.gridColumns ?? 20 })
      affected.push(logicalId)
    }
    if (op.op === 'rename_sheet' && target) {
      for (const sheet of next.sheets)
        for (const cell of Object.values(sheet.cells))
          if (cell.formula)
            (cell as { formula: string }).formula = renameSheetInFormula(
              cell.formula,
              target.name,
              op.name,
            )
    }
    if (isStructuralOp(op)) {
      const spec = shiftSpecForOp(op)
      if (spec && 'sheetId' in op) {
        const templates = typedTemplates.get(op.sheetId) ?? {},
          moved: Record<string, TypedCell> = {}
        for (const [address, cell] of Object.entries(templates)) {
          const point = parseAddress(address),
            shifted = shiftIndex(spec.axis === 'row' ? point.row : point.column, spec)
          if (shifted !== null)
            moved[
              formatAddress(
                spec.axis === 'row' ? shifted : point.row,
                spec.axis === 'column' ? shifted : point.column,
              )
            ] = cell
        }
        typedTemplates.set(op.sheetId, moved)
        const oldKeys = overlayKeys.get(op.sheetId) ?? new Set<string>(),
          nextKeys = new Set<string>()
        const oldKinds = originalKinds.get(op.sheetId) ?? {},
          nextKinds: Record<string, TypedCell> = {}
        for (const address of new Set([...oldKeys, ...Object.keys(oldKinds)])) {
          const point = parseAddress(address),
            shifted = shiftIndex(spec.axis === 'row' ? point.row : point.column, spec)
          if (shifted === null) continue
          const changed = formatAddress(
            spec.axis === 'row' ? shifted : point.row,
            spec.axis === 'column' ? shifted : point.column,
          )
          if (oldKeys.has(address)) nextKeys.add(changed)
          if (oldKinds[address]) nextKinds[changed] = oldKinds[address]!
        }
        overlayKeys.set(op.sheetId, nextKeys)
        originalKinds.set(op.sheetId, nextKinds)
        const size = dimensions.get(op.sheetId)!
        if (spec.axis === 'row') size.rows += spec.delta
        else size.columns += spec.delta
        for (const sheet of next.sheets)
          for (const visual of sheet.visuals ?? []) {
            if (visual.sheetId === op.sheetId)
              for (const key of [
                spec.axis === 'row' ? 'fromRow' : 'fromColumn',
                spec.axis === 'row' ? 'toRow' : 'toColumn',
              ] as const) {
                const value = shiftIndex(visual.anchor[key], spec)
                if (value === null) throw new Error('chart_anchor_deleted')
                visual.anchor[key] = value
              }
            for (const series of visual.chart.series)
              for (const key of ['valuesRef', 'categoriesRef', 'nameRef'] as const)
                if (series[key])
                  series[key] = shiftFormulaText(
                    series[key]!,
                    target!.name,
                    {
                      boundary: spec.delta > 0 ? spec.start : spec.start - spec.delta,
                      delta: spec.delta,
                      deleted:
                        spec.delta < 0
                          ? { start: spec.start, end: spec.start - spec.delta - 1 }
                          : null,
                    },
                    spec.axis,
                    true,
                  )
          }
      }
      log.push(op)
    } else if (op.op === 'rename_sheet') {
      log.push(op)
      for (const sheet of next.sheets)
        for (const visual of sheet.visuals ?? [])
          for (const series of visual.chart.series)
            for (const key of ['valuesRef', 'categoriesRef', 'nameRef'] as const)
              if (series[key])
                series[key] = renameSheetInFormula(series[key]!, target!.name, op.name)
    }
    next = {
      ...next,
      sheets: next.sheets.map((sheet) => ({
        ...sheet,
        gridRows: Math.max(1, dimensions.get(sheet.id)?.rows ?? sheet.gridRows ?? 1000),
        gridColumns: Math.max(1, dimensions.get(sheet.id)?.columns ?? sheet.gridColumns ?? 20),
      })),
    }
    refreshChartReferences(
      next,
      [
        ...material.completeSheetIds,
        ...next.sheets
          .filter((sheet) => !current.sheets.some((previous) => previous.id === sheet.id))
          .map((sheet) => sheet.id),
      ],
      material.referenceRanges,
    )
    await options.afterStep?.(next)
    book = next
    affected.push('sheetId' in op ? op.sheetId : op.op)
    options.fault?.(`afterOperation:${index + 1}`)
  }
  const after: CandidateState = {
    ...before,
    stateId: material.allocation.stateId,
    structuralLog: log,
    sheetMap,
    sheets: book.sheets.map((sheet: WorksheetState) => {
      const original = current.sheets.find((item) => item.id === sheet.id)
      const cells: Record<string, TypedCell> = {}
      for (const address of new Set([
        ...Object.keys(sheet.cells),
        ...Object.keys(sheet.styles ?? {}),
      ])) {
        const cell = sheet.cells[address] ?? { value: null },
          template = typedTemplates.get(sheet.id)?.[address],
          style = typedStyle(sheet.styles?.[address], template?.style)
        const baseline = structuralBaseline.sheets.find((item) => item.id === sheet.id)
        const originalKind = originalKinds.get(sheet.id)?.[address]?.kind
        const numericDate =
          typeof cell.value === 'number' &&
          !!style?.numberFormat &&
          isDateFormat(style.numberFormat)
        const kind: TypedCell['kind'] = numericDate
          ? 'date'
          : template && template.rawValue === cell.value
            ? template.kind
            : cell.value === null
              ? 'blank'
              : typeof cell.value === 'number'
                ? 'number'
                : typeof cell.value === 'boolean'
                  ? 'boolean'
                  : 'string'
        if (
          !overlayKeys.get(sheet.id)?.has(address) &&
          kind === (originalKind ?? 'blank') &&
          JSON.stringify(cell) === JSON.stringify(baseline?.cells[address] ?? { value: null }) &&
          JSON.stringify(sheet.styles?.[address]) === JSON.stringify(baseline?.styles?.[address])
        )
          continue
        cells[address] = {
          kind,
          rawValue: cell.value,
          ...(kind === 'date' ? { dateSystem: before.dateSystem } : {}),
          ...(cell.formula
            ? {
                formula: cell.formula,
                ...(cell.value === null ? {} : { cachedValue: cell.value }),
                calculationState: 'pending',
              }
            : {}),
          ...(style ? { style } : {}),
          ...(style?.numberFormat ? { numberFormat: style.numberFormat } : {}),
        }
      }
      const size = dimensions.get(sheet.id)
      return {
        ...(original ?? {}),
        id: sheet.id,
        name: sheet.name,
        rows: size?.rows ?? sheet.gridRows ?? 1000,
        columns: size?.columns ?? sheet.gridColumns ?? 20,
        protected:
          current.sheets.find((item) => item.id === sheet.id)?.protected ??
          original?.protected ??
          false,
        cells,
        features: JSON.parse(
          JSON.stringify({ ...original?.features, charts: sheet.visuals ?? [] }),
        ),
      }
    }),
  }
  cancelled(signal)
  const changes = stateChangeset(before, after, affected)
  return {
    ...changes,
    calculationState: current.sheets.some((sheet) => !material.completeSheetIds.includes(sheet.id))
      ? ('pending' as const)
      : changes.calculationState,
  }
}
