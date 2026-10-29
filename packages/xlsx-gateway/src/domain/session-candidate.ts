import { buildSubtotalPlan } from './subtotal-plan'
import { sortRowVisibility } from './session-sort'
import {
  visibilityEditSchema,
  visibilityEditsSchema,
  axisHiddenOverride,
} from './session-visibility'
import { isWorksheetViewOperation } from './session-permissions'
import { rewriteNameReferences, type NameReferenceChange } from './defined-name-references'
import {
  editSessionDefinedName,
  sessionDefinedNames,
  manageSessionDefinedName,
  assertNameBindingsComplete,
} from './session-defined-names'
import type { WorkbookVisualObject } from '../shared/workbook-protocol'
import { projectVisualOrder } from './visual-order'
import { buildChartSeriesSelection, chartSelectionReferences } from './chart-series-selection'
import { assertChartFormat } from './chart-format'
import { chartSeriesRangeEdits, resolveChartRange } from './chart-series-range'
import { workbookChartEditSchema, workbookVisualEditSchema } from '../shared/edit-schemas'
import { isDateFormat } from 'numfmt'
import { InMemoryWorkbookAdapter } from './in-memory-workbook'
import {
  formatAddress,
  parseAddress,
  rangeCellCount,
  parseRange,
  columnIndex,
  rangeAddresses,
} from './cell-address'
import { computeSortedRowOrder } from './sort-range'
import { updateSessionFilter } from './session-filter'
import { shiftIndex, shiftSpecForOp } from './formula-shift'
import {
  applyChartStateEdit,
  pruneChartSettingsForType,
  chartKindFromVisual,
  type ChartStateEdit,
  type SheetVisual,
} from './chart-visual'
import type { CellFormatState, WorkbookSnapshot, WorksheetState } from './workbook.types'
import { isStructuralOp, parseWorkbookCommandBatch } from './workbook-dsl'
import { normalizeStyleColor } from './style-color'
import { renameSheetInFormula } from '../gateway/xlsx-sheets'
import { shiftFormulaText } from '../gateway/xlsx-structure'
import {
  dvRuleToWireRule,
  subtractDvRules,
  type DvCellArea,
  type DvWireRule,
} from '../gateway/xlsx-dv'
import {
  transactionInputSchema,
  workbookStateSchema,
  type cellRecordSchema,
} from '../shared/session-state'
import type { z } from 'zod'
import { refreshChartReferences } from './session-chart-references'
import { stateChangeset, type CandidateState } from '../shared/session-changeset'
import { pasteCellFields } from './paste-cell-fields'
import { moveSessionRange } from './session-range-move'
import { moveFormulaRefs } from './formula-shift'
import { copySessionRange } from './session-range-copy'
import { measureNormalFontMdw, type NormalFontDescriptor } from './normal-font-metrics'
import { SESSION_OPERATION_NAMES, hasMixedSessionOperations } from '../shared/session-operations'
import { assertMergeOperation, updateSessionMerges, sessionMerges } from './session-merges'
import {
  PIVOT_SOURCE_COL_LIMIT,
  PIVOT_SOURCE_ROW_LIMIT,
  PivotLayoutError,
  areasOverlap,
  buildPivotLayout,
  pivotOutputArea,
  type PivotLayout,
} from './pivot-layout'
import {
  PivotGrowError,
  grownPivotBlock,
  pivotBlockMatrix,
  pivotPageRows,
  pivotSourceGrid,
} from './pivot-placement'
import { growPivotDefinition, recomputePivotData } from './pivot-engine'
import {
  workbookPivotAddSchema,
  workbookTableAddSchema,
  workbookTableEditSchema,
} from '../shared/workbook-protocol'
import { DEFAULT_SHEET_PROTECTION_ALLOW } from '../gateway/xlsx-protection'
import { hashSheetPassword, verifySheetPassword } from '../gateway/xlsx-protection-hash'
import { MAX_GRID_COLUMNS, MAX_GRID_ROWS } from '../shared/grid-bounds'

type TypedCell = z.infer<typeof cellRecordSchema>
export interface CandidateMaterial {
  /** Current values and feature state, resolved against exactly stateId. */
  state: CandidateState
  /** Full cell/formula/feature coverage for structural/reference changes. */
  completeSheetIds: readonly string[]
  /** Exact current ranges hydrated for chart dependencies, separate from full sheets. */
  referenceRanges?: Record<string, string[]>
  /** Complete host-resolved visibility over requested sort/filter spans. */
  hiddenRows?: Record<string, readonly number[]>
  /** Host-validated protection/readonly state, independent of renderer focus. */
  writableSheetIds: readonly string[]
  /** Indexed source permits view-only changes, including protected sheets; never content writes. */
  viewWritableSheetIds?: readonly string[]
  allocation: { stateId: string; sheetIds: Record<string, string> }
}

const whitelist = new Set<string>(SESSION_OPERATION_NAMES)
// The extension host OS and Webview language can differ. Both executions
// use the same explicit Han/Latin collation, including typed/style movement.
const SESSION_SORT_LOCALE = 'zh-CN'
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
    'underlineStyle',
    'vertAlign',
    'protectionLocked',
    'protectionHidden',
    'strikethrough',
    'fontFamily',
    'fontSize',
    'fontColor',
    'fillColor',
    'fill',
    'numberFormat',
    'wrapText',
    'indent',
    'textRotation',
    'borderTop',
    'borderBottom',
    'borderLeft',
    'borderRight',
  ]
  return JSON.parse(
    JSON.stringify({
      ...Object.fromEntries(Object.entries(style).filter(([key]) => keys.includes(key))),
      ...(style.textRotation !== undefined
        ? {
            textRotation:
              style.textRotation === 255
                ? 'vertical'
                : style.textRotation > 90
                  ? 90 - style.textRotation
                  : style.textRotation,
          }
        : {}),
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
  if (Object.hasOwn(simple, 'fillColor') && !Object.hasOwn(simple, 'fill')) delete result.fill
  if (format.textRotation !== undefined)
    result.textRotation =
      format.textRotation === 'vertical'
        ? 255
        : format.textRotation < 0
          ? 90 - format.textRotation
          : format.textRotation
  if (border)
    for (const edge of ['top', 'bottom', 'left', 'right'] as const)
      if (
        (border.type === 'all' || border.type === 'none' || border.type === edge) &&
        !Object.hasOwn(simple, 'border' + edge[0]!.toUpperCase() + edge.slice(1))
      )
        result[('border' + edge[0]!.toUpperCase() + edge.slice(1)) as 'borderTop'] =
          border.type === 'none'
            ? null
            : {
                style: border.style ?? 'thin',
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
      rowHeights: structuredClone((sheet.features?.rowHeights ?? {}) as Record<string, number>),
      colWidths: structuredClone((sheet.features?.colWidths ?? {}) as Record<string, number>),
      visuals: structuredClone(
        (sheet.features?.visuals ?? sheet.features?.charts ?? []) as unknown as SheetVisual[],
      ),
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
): Promise<ReturnType<typeof stateChangeset>> {
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
  const current = workbookStateSchema.parse(material.state)
  if (
    current.stateId !== before.stateId ||
    current.sheets.map((sheet) => sheet.id).join() !== before.sheets.map((sheet) => sheet.id).join()
  )
    throw new Error('stale_candidate_material')
  const subtotal = intent.operations.find((op) => op.op === 'subtotal')
  if (subtotal) {
    if (intent.operations.length !== 1)
      throw new Error('Subtotal must be a standalone transaction.')
    const sheet = current.sheets.find((sheet) => sheet.id === subtotal.sheetId)
    if (!sheet) throw new Error('unknown_sheet')
    const bounds = parseRange(subtotal.range)
    if ((bounds.endRow - bounds.startRow + 1) * (bounds.endColumn - bounds.startColumn + 1) > 2000)
      throw new Error('Subtotal supports at most 2000 source cells.')
    const groupCol = columnIndex(subtotal.groupColumn),
      valueCol = columnIndex(subtotal.valueColumn)
    const labels: unknown[] = []
    for (let row = bounds.startRow + 1; row <= bounds.endRow; row++) {
      const cell = sheet.cells[formatAddress(row, groupCol)]
      if (cell?.kind === 'error') throw new Error('Subtotal cannot group cells containing errors.')
      if (
        cell?.formula &&
        (cell.calculationState !== 'ready' || !Object.hasOwn(cell, 'cachedValue'))
      )
        throw new Error('Wait for Subtotal group formulas to finish calculating, then retry.')
      labels.push(cell?.formula ? cell.cachedValue : cell?.rawValue)
    }
    for (const item of current.sheets) {
      const metadata = item.features?.originalMetadata as
        { outlineLevelRow?: number; outlineLevelCol?: number } | undefined
      if (metadata?.outlineLevelRow || metadata?.outlineLevelCol)
        throw new Error('unsupported_structural_feature: source_outline')
    }
    const plan = buildSubtotalPlan(bounds, labels, { groupCol, valueCol, agg: subtotal.agg })
    if (sheet.rows + plan.insertBefore.length > 1_048_576)
      throw new Error('Subtotal exceeds the worksheet row limit.')
    const structuralOps = plan.insertBefore.map((row) => ({
      op: 'insert_rows' as const,
      sheetId: sheet.id,
      row: row + 1,
      count: 1,
    }))
    // Private stages are never installed or journaled as separate commits.
    // Hydrated source cells are retained across the coordinate/content boundary.
    let stage = await computeCandidateState(
      current,
      { ...intent, operations: structuralOps },
      material,
      options,
    )
    const writes: import('./workbook-dsl').WorkbookOperation[] = []
    for (const total of [...plan.totals, plan.grand]) {
      const label = formatAddress(total.row, groupCol),
        value = formatAddress(total.row, valueCol)
      writes.push(
        { op: 'set_cell', sheetId: sheet.id, address: label, value: total.label },
        { op: 'set_formula', sheetId: sheet.id, address: value, formula: total.formula },
        { op: 'format_range', sheetId: sheet.id, range: label, format: { bold: true } },
        { op: 'format_range', sheetId: sheet.id, range: value, format: { bold: true } },
      )
    }
    writes.push({
      op: 'set_outline_settings',
      sheetId: sheet.id,
      summaryBelow: true,
      summaryRight:
        (sheet.features?.originalMetadata as { outlineSummaryRight?: boolean } | undefined)
          ?.outlineSummaryRight !== false,
    })
    writes.push({
      op: 'set_outline',
      sheetId: sheet.id,
      axis: 'rows',
      start: bounds.startRow + 2,
      count: plan.grand.row - bounds.startRow - 1,
      level: 1,
      collapsed: false,
    })
    for (const group of plan.totals)
      writes.push({
        op: 'set_outline',
        sheetId: sheet.id,
        axis: 'rows',
        start: group.detailStart + 1,
        count: group.detailEnd - group.detailStart + 1,
        level: 2,
        collapsed: false,
      })
    const affected = [...stage.affectedTargets]
    for (let index = 0; index < writes.length; index += 500) {
      const state = stage.nextState
      stage = await computeCandidateState(
        state,
        { ...intent, operations: writes.slice(index, index + 500) },
        { ...material, state },
        options,
      )
      affected.push(...stage.affectedTargets)
    }
    return stateChangeset(before, stage.nextState, affected)
  }
  if (intent.operations.some((op) => op.op === 'refresh_pivot')) {
    if (!intent.operations.every((op) => op.op === 'refresh_pivot'))
      throw new Error('PivotTable refresh must be a standalone transaction.')
    return refreshSessionPivots(
      before,
      current,
      intent,
      intent.operations as Extract<
        import('./workbook-dsl').WorkbookOperation,
        { op: 'refresh_pivot' }
      >[],
      material,
      options,
    )
  }
  const editPivot = intent.operations.find((op) => op.op === 'edit_pivot')
  if (editPivot) {
    if (intent.operations.length !== 1)
      throw new Error('A PivotTable edit must be a standalone transaction.')
    return editSessionPivot(before, current, intent, editPivot, material, options)
  }
  const pivotOp = intent.operations.find((op) => op.op === 'add_pivot')
  if (pivotOp) {
    if (intent.operations.length !== 1)
      throw new Error('A PivotTable must be created in a standalone transaction.')
    return addSessionPivot(before, current, intent, pivotOp, material, options)
  }
  let workbookFeatures = structuredClone(current.features ?? {})
  const workbookMetadata = current.features?.originalMetadata as Record<string, unknown> | undefined
  if (
    intent.operations.some((op) => changingCoordinates.has(op.op) || op.op === 'add_sheet') &&
    workbookStructureLocked(current.features)
  )
    throw new Error('protected_workbook_structure')
  if (
    intent.operations.some((op) => changingCoordinates.has(op.op)) &&
    current.sheets.some(
      (sheet) =>
        (sheet.features?.tableAdditions as unknown[] | undefined)?.length ||
        sheet.features?.tableEdits,
    )
  )
    throw new Error('unsupported_structural_feature: session_table')
  // A pending PivotTable pins its source and output coordinates and sheet names.
  const pendingPivots = sessionPivotAdditions(current)
  if (pendingPivots.length) {
    if (intent.operations.some((op) => changingCoordinates.has(op.op)))
      throw new Error('unsupported_structural_feature: session_pivot')
    for (const op of intent.operations)
      if (
        (op.op === 'delete_sheet' || op.op === 'rename_sheet') &&
        pendingPivots.some(
          (pivot) => pivot.sheetId === op.sheetId || pivot.sourceSheetId === op.sheetId,
        )
      )
        throw new Error('unsupported_structural_feature: session_pivot')
  }
  if (intent.operations.some((op) => changingCoordinates.has(op.op))) {
    if (current.sheets.some((sheet) => !material.completeSheetIds.includes(sheet.id)))
      throw new Error('incomplete_structural_material')
    if (
      sessionDefinedNames(current).names.length ||
      sessionDefinedNames(current).preserveNames.length ||
      (Array.isArray(current.features?.definedNameReferences) &&
        current.features.definedNameReferences.length > 0) ||
      (Array.isArray(workbookMetadata?.reservedDefinedNames) &&
        workbookMetadata.reservedDefinedNames.length > 0)
    )
      throw new Error('unsupported_structural_feature: definedNames')
    let formulaCount = 0
    for (const sheet of current.sheets) {
      formulaCount += Object.values(sheet.cells).filter((cell) => cell.formula).length
      const metadata = sheet.features?.originalMetadata as Record<string, unknown> | undefined
      if ((sheet.features?.visibilityEdits as unknown[] | undefined)?.length)
        throw new Error('unsupported_structural_feature: session_visibility')
      if ((sheet.features?.outlineEdits as unknown[] | undefined)?.length)
        throw new Error('unsupported_structural_feature: session_outline')
      if (sheet.features?.mergeEdits)
        throw new Error('unsupported_structural_feature: session_merges')
      if (Object.keys((sheet.features?.noteEdits ?? {}) as object).length)
        throw new Error('unsupported_structural_feature: session_notes')
      if (sheet.features?.hyperlinkEdits)
        throw new Error('unsupported_structural_feature: session_hyperlinks')
      if (sheet.features?.filterEdits)
        throw new Error('unsupported_structural_feature: session_filter')
      if (sheet.features?.autoFilter) throw new Error('unsupported_structural_feature: autoFilter')
      const page = sheet.features?.pageSetupEdits as
        { printArea?: unknown; printTitles?: unknown } | undefined
      if (page?.printArea || page?.printTitles)
        throw new Error('unsupported_structural_feature: print_references')
      const freeze = sheet.features?.freezeEdits as
        { frozenRows?: number; frozenColumns?: number } | undefined
      if (freeze?.frozenRows || freeze?.frozenColumns)
        throw new Error('unsupported_structural_feature: freeze')
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
  const featureEdits = new Map<string, Record<string, unknown>>()
  const log = structuredClone(before.structuralLog)
  const sheetMap = structuredClone(before.sheetMap)
  const nativeVisuals =
    (
      before.features?.originalMetadata as
        | {
            visuals?: WorkbookVisualObject[]
          }
        | undefined
    )?.visuals ?? []
  const affected: string[] = []
  let expandedWrites = 0
  for (const [index, op] of intent.operations.entries()) {
    cancelled(signal)
    if (op.op === 'add_defined_names') {
      assertNameBindingsComplete({ features: workbookFeatures })
      for (const entry of op.names) {
        cancelled(signal)
        workbookFeatures = {
          ...workbookFeatures,
          definedNamesEdits: JSON.parse(
            JSON.stringify(
              editSessionDefinedName(
                { ...current, features: workbookFeatures },
                { op: 'add_defined_name', ...entry },
              ),
            ),
          ),
        }
      }
      affected.push(...op.names.map((name) => name.name))
      await options.afterStep?.(book)
      options.fault?.(`afterOperation:${index + 1}`)
      continue
    }
    if (op.op === 'manage_defined_name') {
      const managed = manageSessionDefinedName({ ...current, features: workbookFeatures }, op)
      const previous = (workbookFeatures.definedNameReferences ??
        []) as unknown as NameReferenceChange[]
      if (managed.change) {
        const plan = {
          changes: [managed.change],
          sheets: current.sheets.map(({ id, name }) => ({ id, name })),
        }
        book = {
          ...book,
          sheets: book.sheets.map((sheet) => {
            cancelled(signal)
            return {
              ...sheet,
              cells: Object.fromEntries(
                Object.entries(sheet.cells).map(([address, cell]) => {
                  const formula = cell.formula
                    ? rewriteNameReferences(cell.formula, sheet.id, plan)
                    : undefined
                  return [
                    address,
                    formula && formula !== cell.formula ? { ...cell, formula, value: null } : cell,
                  ]
                }),
              ),
            }
          }),
        }
      }
      workbookFeatures = {
        ...workbookFeatures,
        definedNamesEdits: JSON.parse(JSON.stringify(managed.names)),
        definedNameReferences: JSON.parse(
          JSON.stringify([...previous, ...(managed.change ? [managed.change] : [])]),
        ),
      }
      affected.push(op.name)
      await options.afterStep?.(book)
      options.fault?.(`afterOperation:${index + 1}`)
      continue
    }
    if (op.op === 'protect_workbook') {
      const original = workbookMetadata?.workbookProtection as { hasPassword?: boolean } | undefined
      if (original?.hasPassword) throw new Error('protection:workbook_password')
      workbookFeatures = {
        ...workbookFeatures,
        workbookProtectionEdits: { lockStructure: op.locked },
      }
      affected.push('workbook')
      continue
    }
    if (op.op === 'add_defined_name' || op.op === 'delete_defined_name') {
      workbookFeatures = {
        ...workbookFeatures,
        definedNamesEdits: JSON.parse(
          JSON.stringify(editSessionDefinedName({ ...current, features: workbookFeatures }, op)),
        ),
      }
      affected.push(op.name)
      await options.afterStep?.(book)
      options.fault?.(`afterOperation:${index + 1}`)
      continue
    }
    cancelled(signal)
    const target =
      'sheetId' in op
        ? book.sheets.find((sheet) => sheet.id === op.sheetId)
        : book.sheets.find(
            (sheet) =>
              sheet.visuals?.some((visual) =>
                'chartPath' in op
                  ? visual.id === op.chartPath
                  : 'visualId' in op
                    ? visual.id === op.visualId
                    : false,
              ) ||
              (('visualId' in op || 'chartPath' in op) &&
                nativeVisuals.some(
                  (visual) =>
                    ('visualId' in op
                      ? visual.id === op.visualId
                      : visual.chartPath === op.chartPath) &&
                    visual.sheetId ===
                      current.sheets.find((item) => item.id === sheet.id)?.originalSheetId,
                )),
          )
    const permissions = current.sheets.find((sheet) => sheet.id === target?.id)
    const viewWrite =
      !!target &&
      isWorksheetViewOperation(op) &&
      material.viewWritableSheetIds?.includes(target.id) === true
    if (target && !material.writableSheetIds.includes(target.id) && !viewWrite)
      throw new Error('unverified_writable_target')
    if (
      workbookMetadata?.readOnly ||
      permissions?.readOnly ||
      (permissions?.protected && !viewWrite)
    )
      throw new Error('protected_target')
    assertMergeOperation(
      {
        ...current,
        sheets: current.sheets.map((sheet) => ({
          ...sheet,
          features: featureEdits.get(sheet.id) ?? sheet.features,
        })),
      },
      op,
    )
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
    const checkedChartRange = (range: string, sheetId: string) => {
      const resolved = resolveChartRange(range, book.sheets, sheetId)
      const covered =
        material.completeSheetIds.includes(resolved.sheetId) ||
        (material.referenceRanges?.[resolved.sheetId] ?? []).some((area) => {
          const bounds = parseRange(area)
          return (
            bounds.startRow <= resolved.bounds.startRow &&
            bounds.endRow >= resolved.bounds.endRow &&
            bounds.startColumn <= resolved.bounds.startColumn &&
            bounds.endColumn >= resolved.bounds.endColumn
          )
        })
      if (!covered) throw new Error('incomplete_chart_material')
      const typed = current.sheets.find((sheet) => sheet.id === resolved.sheetId)!
      const source = book.sheets.find((sheet) => sheet.id === resolved.sheetId)!
      for (const address of rangeAddresses(resolved.bounds)) {
        const cell = typed.cells[address]
        if (
          source.cells[address]?.formula &&
          (cell?.formula !== source.cells[address]?.formula ||
            (cell.calculationState !== 'ready' && cell.calculationState !== 'error'))
        )
          throw new Error('chart_formula_pending')
      }
      return resolved
    }
    if (op.op === 'edit_chart' && target) {
      if (op.seriesData && op.seriesSelection) throw new Error('chart_series_edit_conflict')
      for (const entry of op.seriesData ?? [])
        for (const range of [entry.valuesRange, entry.categoriesRange])
          if (range) checkedChartRange(range, entry.sheetId ?? target.id)
      if (op.seriesSelection) {
        const native = nativeVisuals.find((v) => v.chartPath === op.chartPath)
        const own = target.visuals?.find((v) => v.id === op.chartPath)
        const overlays = (featureEdits.get(target.id)?.chartEdits ??
          permissions?.features?.chartEdits ??
          {}) as Record<string, ChartStateEdit>
        const chart =
          own?.chart ??
          (native?.chart ? applyChartStateEdit(native.chart, overlays[op.chartPath]) : undefined)
        if (!chart) throw new Error('unknown_chart')
        for (const reference of chartSelectionReferences(chart, op.seriesSelection))
          checkedChartRange(reference, target.id)
      }
    }
    let next: WorkbookSnapshot
    if (op.op === 'edit_chart' && op.chartPath.startsWith('xl/')) {
      const matches = nativeVisuals.filter((visual) => visual.chartPath === op.chartPath)
      if (matches.length !== 1 || !matches[0]?.chart || !target) throw new Error('unknown_chart')
      const native = matches[0]
      const features = featureEdits.get(target.id) ?? structuredClone(permissions?.features ?? {})
      if (
        (features.visualEdits as Record<string, { remove?: true }> | undefined)?.[native.id]?.remove
      )
        throw new Error('visual_removed')
      const { op: _op, seriesData, seriesSelection, ...patch } = op
      const chart = native.chart!
      const edits = (features.chartEdits ?? {}) as Record<string, Record<string, unknown>>
      let previous = edits[op.chartPath] ?? {}
      const currentChart = applyChartStateEdit(chart, previous as ChartStateEdit)
      const seriesSet = seriesSelection
        ? buildChartSeriesSelection(currentChart, seriesSelection, (reference) => {
            const resolved = checkedChartRange(reference, target.id)
            const source = book.sheets.find((sheet) => sheet.id === resolved.sheetId)!
            return {
              ref: resolved.ref,
              vector: rangeAddresses(resolved.bounds).map(
                (address) => source.cells[address]?.value ?? null,
              ),
            }
          })
        : undefined
      if (seriesSet) {
        previous = { ...previous }
        for (const key of [
          'series',
          'seriesSet',
          'seriesColors',
          'pointColors',
          'pointExplosions',
          'explosionPct',
        ])
          delete previous[key]
      }
      const kind = chartKindFromVisual(chart)
      if (
        op.chartType !== undefined &&
        (!kind || !['column', 'bar', 'line', 'area', 'pie', 'doughnut'].includes(kind))
      )
        throw new Error('unsupported_chart_conversion')
      const series = chartSeriesRangeEdits(
        book,
        target.id,
        currentChart.series.length,
        seriesData ?? [],
        (sheetId, address) => {
          const cell = current.sheets.find((sheet) => sheet.id === sheetId)?.cells[address]
          if (
            cell?.formula &&
            cell.calculationState !== 'ready' &&
            cell.calculationState !== 'error'
          )
            throw new Error('chart_formula_pending')
        },
      )
      let merged = {
        ...previous,
        ...patch,
        ...(seriesSet ? { seriesSet } : {}),
        ...(series.length
          ? {
              series: (() => {
                const previousSeries = (previous.series ?? []) as NonNullable<
                  ReturnType<typeof chartSeriesRangeEdits>
                >
                const merged = new Map(previousSeries.map((entry) => [entry.index, entry]))
                for (const entry of series)
                  merged.set(entry.index, { ...merged.get(entry.index), ...entry })
                return [...merged.values()]
              })(),
            }
          : {}),
        ...(op.seriesColors
          ? { seriesColors: { ...(previous.seriesColors as object), ...op.seriesColors } }
          : {}),
        ...(op.pointColors
          ? {
              pointColors: Object.fromEntries(
                Array.from(
                  new Set([
                    ...Object.keys((previous.pointColors ?? {}) as object),
                    ...Object.keys(op.pointColors),
                  ]),
                  (index) => [
                    index,
                    {
                      ...((previous.pointColors as Record<string, object> | undefined)?.[index] ??
                        {}),
                      ...op.pointColors?.[index],
                    },
                  ],
                ),
              ),
            }
          : {}),
        ...(op.valueAxis
          ? { valueAxis: { ...(previous.valueAxis as object), ...op.valueAxis } }
          : {}),
        ...(op.pointExplosions
          ? { pointExplosions: { ...(previous.pointExplosions as object), ...op.pointExplosions } }
          : {}),
        ...(op.axisTitles
          ? { axisTitles: { ...(previous.axisTitles as object), ...op.axisTitles } }
          : {}),
      }
      assertChartFormat(currentChart, {
        ...patch,
        ...(seriesSet ? { seriesSet } : {}),
        ...(seriesSet ? { seriesSet } : {}),
        ...(series.length ? { series } : {}),
      })
      if (op.chartType !== undefined) merged = pruneChartSettingsForType(merged, op.chartType)
      const edit = workbookChartEditSchema.parse(merged)
      const effective = applyChartStateEdit(chart, edit)
      if (op.axisTitles && effective.chartTypes.some((type) => /pie|doughnut/i.test(type)))
        throw new Error('unsupported_chart_axes')
      if (
        op.grouping &&
        !effective.chartTypes.some((type) => /^(barChart|lineChart|areaChart)$/.test(type))
      )
        throw new Error('unsupported_chart_grouping')
      features.chartEdits = { ...edits, [op.chartPath]: edit }
      featureEdits.set(target.id, features)
      next = structuredClone(book)
    } else if (
      op.op === 'set_visual_anchor' ||
      op.op === 'delete_visual' ||
      op.op === 'edit_visual' ||
      op.op === 'edit_shape'
    ) {
      const own = target?.visuals?.find((visual) => visual.id === op.visualId)
      const native = nativeVisuals.find(
        (visual) => visual.id === op.visualId && visual.sheetId === permissions?.originalSheetId,
      )
      if (!own && native) {
        if (op.op === 'edit_shape') throw new Error('unknown_shape')
        if (!native.drawingPath || native.drawingIndex === undefined)
          throw new Error('unsupported_visual_anchor')
        const features =
          featureEdits.get(target!.id) ?? structuredClone(permissions?.features ?? {})
        const edits = (features.visualEdits ?? {}) as Record<
          string,
          { anchor?: unknown; remove?: true }
        >
        if (edits[native.id]?.remove) throw new Error('visual_removed')
        if (op.op === 'edit_visual') {
          if ((op.fillColor !== undefined || op.lineColor !== undefined) && native.kind !== 'shape')
            throw new Error('unsupported_visual_paint')
          if (
            (op.rotation !== undefined ||
              op.flipH !== undefined ||
              op.flipV !== undefined ||
              op.frameSize !== undefined) &&
            native.kind !== 'shape' &&
            native.kind !== 'image'
          )
            throw new Error('unsupported_visual_transform')
          if (op.frameSize !== undefined && !edits[native.id]?.anchor)
            throw new Error('visual_frame_requires_anchor')

          if (nativeVisuals.filter((v) => v.id === op.visualId).length !== 1)
            throw new Error('ambiguous_visual')
          if (op.editAs !== undefined && native.editAs === undefined)
            throw new Error('unsupported_visual_placement')
          if (op.zIndex !== undefined) {
            const part = nativeVisuals.filter(
              (v) =>
                v.drawingPath === native.drawingPath &&
                v.drawingIndex !== undefined &&
                !edits[v.id]?.remove,
            )
            if (op.zIndex >= part.length) throw new Error('invalid_visual_z_index')
          }
        }
        const patch =
          op.op === 'delete_visual'
            ? { remove: true as const }
            : op.op === 'edit_visual'
              ? (({ op: _op, visualId: _id, ...properties }) => properties)(op)
              : { anchor: structuredClone(op.anchor) }
        workbookVisualEditSchema.parse({
          ...edits[native.id],
          ...patch,
          drawingPath: native.drawingPath,
          drawingIndex: native.drawingIndex,
        })
        const mergedEdits = {
          ...edits,
          [native.id]:
            op.op === 'delete_visual'
              ? { remove: true as const }
              : { ...edits[native.id], ...patch },
        }
        if (op.op === 'edit_visual' && op.zIndex !== undefined) {
          const part = projectVisualOrder(
            nativeVisuals.filter(
              (v) =>
                v.sheetId === native.sheetId &&
                v.drawingPath === native.drawingPath &&
                v.drawingIndex !== undefined &&
                !edits[v.id]?.remove,
            ),
            (v) => (edits[v.id] as { zIndex?: number } | undefined)?.zIndex,
          )
          const at = part.findIndex((v) => v.id === native.id)
          const reordered = [...part]
          reordered.splice(at, 1)
          reordered.splice(op.zIndex, 0, native)
          reordered.forEach((visual, zIndex) => {
            mergedEdits[visual.id] = { ...mergedEdits[visual.id], zIndex }
          })
        }
        features.visualEdits = mergedEdits
        featureEdits.set(target!.id, features)
        next = structuredClone(book)
      } else {
        if (!own) throw new Error('unknown_visual')
        const adapter = new InMemoryWorkbookAdapter(book)
        adapter.apply(
          adapter.plan({
            dslVersion: 1,
            transactionId: `${intent.transactionId}-${index}`,
            baseRevision: book.revision,
            summary: intent.summary,
            operations: [op],
          }),
        )
        next = adapter.getSnapshot()
      }
    } else if (
      op.op === 'merge_cells' ||
      op.op === 'unmerge_cells' ||
      op.op === 'set_page_setup' ||
      op.op === 'set_freeze' ||
      op.op === 'set_note' ||
      op.op === 'set_hyperlink' ||
      op.op === 'set_filter' ||
      op.op === 'clear_filter' ||
      op.op === 'clear_filter_criteria' ||
      op.op === 'reapply_filter' ||
      op.op === 'set_filter_criteria' ||
      op.op === 'set_data_validation' ||
      op.op === 'set_rows_hidden' ||
      op.op === 'set_cols_hidden' ||
      op.op === 'set_outline' ||
      op.op === 'set_outline_settings' ||
      op.op === 'protect_sheet' ||
      op.op === 'set_protected_ranges' ||
      op.op === 'set_thread' ||
      op.op === 'add_table' ||
      op.op === 'edit_table'
    ) {
      const features =
        featureEdits.get(op.sheetId) ??
        structuredClone(current.sheets.find((sheet) => sheet.id === op.sheetId)!.features ?? {})
      if (op.op === 'merge_cells' || op.op === 'unmerge_cells') {
        const cells = book.sheets.find((sheet) => sheet.id === op.sheetId)!.cells
        features.mergeEdits = updateSessionMerges(
          features,
          op,
          Object.fromEntries(
            Object.entries(cells).map(([address, cell]) => [
              address,
              { rawValue: cell.value, ...(cell.formula ? { formula: cell.formula } : {}) },
            ]),
          ),
        )
        if (op.op === 'merge_cells') {
          const range = parseRange(op.range),
            size = dimensions.get(op.sheetId)!
          size.rows = Math.max(size.rows, range.endRow + 1)
          size.columns = Math.max(size.columns, range.endColumn + 1)
        }
      } else if (
        op.op === 'set_filter' ||
        op.op === 'clear_filter' ||
        op.op === 'clear_filter_criteria' ||
        op.op === 'reapply_filter' ||
        op.op === 'set_filter_criteria'
      ) {
        const nextFilter = updateSessionFilter(
          features,
          op,
          book,
          current.dateSystem,
          (current.features?.originalMetadata as { themeColors?: string[] } | undefined)
            ?.themeColors ?? undefined,
        )
        // A source row hidden without live filter criteria is manually hidden.
        // Record that provenance before canonical filter visibility replaces the
        // file cache; the same journal survives projection, sorting and saving.
        if (!features.filterEdits) {
          const visibility = visibilityEditsSchema.parse(features.visibilityEdits ?? [])
          const originalFilter = features.autoFilter as
            { startRow: number; endRow: number } | undefined
          const hasCriteria =
            Array.isArray(features.autoFilterColumns) && features.autoFilterColumns.length > 0
          for (const row of material.hiddenRows?.[op.sheetId] ?? []) {
            if (
              row <= nextFilter.visibilityRange.startRow ||
              row > nextFilter.visibilityRange.endRow ||
              axisHiddenOverride(visibility, 'rows', row) !== undefined ||
              (hasCriteria &&
                originalFilter &&
                row > originalFilter.startRow &&
                row <= originalFilter.endRow)
            )
              continue
            visibility.push({ axis: 'rows', start: row, count: 1, hidden: true })
          }
          if (visibility.length) features.visibilityEdits = visibility
        }
        if ((op.op === 'set_filter_criteria' || op.op === 'reapply_filter') && nextFilter.filter) {
          const typed = current.sheets.find((sheet) => sheet.id === op.sheetId)!
          const source = book.sheets.find((sheet) => sheet.id === op.sheetId)!
          const contentOps = new Set([
            'set_cell',
            'set_formula',
            'clear_cell',
            'set_range',
            'clear_range',
            'paste_cells',
            'sort_range',
            'find_replace',
            'fill_range',
            'copy_range',
            'relocate_range',
            'add_defined_name',
            'add_defined_names',
            'manage_defined_name',
            'delete_defined_name',
            'set_rows_hidden',
            'set_cols_hidden',
          ])
          const changedEarlier = intent.operations
            .slice(0, index)
            .some((earlier) => contentOps.has(earlier.op))
          for (const column of nextFilter.filter.columns.filter((column) => !column.colors))
            for (
              let row = nextFilter.filter.range.startRow + 1;
              row <= nextFilter.filter.range.endRow;
              row++
            ) {
              const address = formatAddress(row, nextFilter.filter.range.startColumn + column.colId)
              const cell = typed.cells[address]
              if (
                source.cells[address]?.formula &&
                (changedEarlier ||
                  cell?.formula !== source.cells[address]!.formula ||
                  (cell.calculationState !== 'ready' && cell.calculationState !== 'error') ||
                  !Object.hasOwn(cell, 'cachedValue'))
              )
                throw new Error(
                  'filter_formula_pending: recalculate current source values before filtering',
                )
            }
        }
        features.filterEdits = nextFilter
      } else if (op.op === 'set_data_validation') {
        const area = parseRange(op.range),
          edits = (features.dataValidationEdits ?? { rules: [], remove: [] }) as {
            rules: DvWireRule[]
            remove: DvCellArea[]
          }
        edits.rules = subtractDvRules(edits.rules, [area])
        if (!edits.remove.some((range) => JSON.stringify(range) === JSON.stringify(area)))
          edits.remove.push(area)
        if (op.validation !== null)
          edits.rules.push({ ranges: [area], rule: dvRuleToWireRule(op.validation) })
        features.dataValidationEdits = edits
      } else if (op.op === 'set_rows_hidden' || op.op === 'set_cols_hidden') {
        const edit = visibilityEditSchema.parse({
          axis: op.op === 'set_rows_hidden' ? 'rows' : 'cols',
          start: op.op === 'set_rows_hidden' ? op.row - 1 : columnIndex(op.column),
          count: op.count,
          hidden: op.hidden,
        })
        features.visibilityEdits = [...((features.visibilityEdits ?? []) as unknown[]), edit]
      } else if (op.op === 'set_outline' || op.op === 'set_outline_settings') {
        features.outlineEdits = JSON.parse(
          JSON.stringify([...((features.outlineEdits ?? []) as unknown[]), op]),
        )
      } else if (op.op === 'add_table' || op.op === 'edit_table') {
        const owner = current.sheets.find((sheet) => sheet.id === op.sheetId)!
        if (!owner.originalSheetId) throw new Error('table:addedSheet')
        const tables = (sheetId: string) => {
          const sheet = current.sheets.find((item) => item.id === sheetId)!
          return sessionTablesOf(
            sheet.id === op.sheetId
              ? { ...sheet, features }
              : {
                  ...sheet,
                  features: featureEdits.get(sheet.id) ?? sheet.features,
                },
          )
        }
        const cells = book.sheets.find((sheet) => sheet.id === op.sheetId)!.cells
        const headerNames = (area: TableArea) => {
          const names: string[] = []
          for (let column = area.startColumn; column <= area.endColumn; column++) {
            const cell = cells[formatAddress(area.startRow, column)]
            const value = cell?.formula ? null : (cell?.rawValue ?? cell?.value)
            const text = value === null || value === undefined ? '' : String(value).trim()
            if (!text || names.some((name) => name.toLowerCase() === text.toLowerCase()))
              throw new Error('table:headers')
            names.push(text.slice(0, 255))
          }
          return names
        }
        const allNames = (except?: string) => [
          ...current.sheets.map((sheet) => sheet.name.toLowerCase()),
          ...current.sheets.flatMap((sheet) =>
            tables(sheet.id)
              .filter((table) => table.name !== except || sheet.id !== op.sheetId)
              .map((table) => table.name.toLowerCase()),
          ),
        ]
        const overlaps = (area: TableArea, except?: string) =>
          tables(op.sheetId).some(
            (table) =>
              table.name !== except &&
              !(
                area.endRow < table.area.startRow ||
                table.area.endRow < area.startRow ||
                area.endColumn < table.area.startColumn ||
                table.area.endColumn < area.startColumn
              ),
          )
        if (op.op === 'add_table') {
          const area = parseRange(op.range)
          if (area.endRow <= area.startRow) throw new Error('table:needsRows')
          if (area.endColumn - area.startColumn + 1 > 1000) throw new Error('table:tooWide')
          if (overlaps(area)) throw new Error('table:overlap')
          const taken = allNames()
          let name = op.name
          if (name === undefined) {
            let index = 1
            while (taken.includes(`table${index}`)) index++
            name = `Table${index}`
          } else if (taken.includes(name.toLowerCase()))
            throw new Error(`table:nameUsed:${JSON.stringify({ name })}`)
          features.tableAdditions = [
            ...((features.tableAdditions ?? []) as unknown[]),
            JSON.parse(
              JSON.stringify(
                workbookTableAddSchema.parse({
                  sheetId: op.sheetId,
                  area,
                  name,
                  columnNames: headerNames(area),
                  ...(op.style === undefined ? {} : { style: op.style }),
                  bandedRows: op.bandedRows ?? true,
                }),
              ),
            ),
          ]
        } else {
          const table = tables(op.sheetId).find((item) => item.name === op.table)
          if (!table) throw new Error(`table:notFound:${JSON.stringify({ name: op.table })}`)
          if (op.name !== undefined && op.name !== table.name) {
            if (allNames(table.name).includes(op.name.toLowerCase()))
              throw new Error(`table:nameUsed:${JSON.stringify({ name: op.name })}`)
          }
          const area = op.area ? parseRange(op.area) : undefined
          if (area && overlaps(area, table.name)) throw new Error('table:overlap')
          const headerRow = op.headerRow ?? table.options.headerRow
          const columnNames =
            area && headerRow
              ? headerNames(area)
              : area
                ? Array.from(
                    { length: area.endColumn - area.startColumn + 1 },
                    (_, index) =>
                      table.columnNames[area.startColumn + index - table.area.startColumn] ??
                      `Column${index + 1}`,
                  )
                : op.headerRow === true
                  ? headerNames({ ...table.area })
                  : undefined
          const patch = {
            ...(op.name === undefined ? {} : { name: op.name }),
            ...(area ? { area } : {}),
            ...(columnNames ? { columnNames } : {}),
            ...(op.style === undefined ? {} : { style: op.style }),
            ...Object.fromEntries(
              (
                [
                  'headerRow',
                  'totalsRow',
                  'bandedRows',
                  'bandedColumns',
                  'firstColumn',
                  'lastColumn',
                  'filterButton',
                ] as const
              ).flatMap((key) => (op[key] === undefined ? [] : [[key, op[key]]])),
            ),
          }
          if (table.kind === 'session') {
            const additions = (features.tableAdditions ?? []) as { name: string }[]
            features.tableAdditions = op.remove
              ? additions.filter((item) => item.name !== table.name)
              : additions.map((item) =>
                  item.name === table.name
                    ? JSON.parse(
                        JSON.stringify(workbookTableAddSchema.parse({ ...item, ...patch })),
                      )
                    : item,
                )
          } else {
            const edits = (features.tableEdits ?? {}) as Record<string, Record<string, unknown>>
            features.tableEdits = {
              ...edits,
              [table.originalName]: JSON.parse(
                JSON.stringify(
                  workbookTableEditSchema.parse({
                    ...edits[table.originalName],
                    sheetId: op.sheetId,
                    tableName: table.originalName,
                    ...patch,
                    ...(op.remove ? { remove: true } : {}),
                  }),
                ),
              ),
            }
          }
        }
      } else if (op.op === 'set_thread') {
        features.threadEdits = {
          ...((features.threadEdits ?? {}) as Record<string, unknown>),
          [op.address]: op.thread === null ? null : JSON.parse(JSON.stringify(op.thread)),
        }
      } else if (op.op === 'set_protected_ranges') {
        const source = (features.protectedRanges ?? []) as { hasPassword?: boolean }[]
        // A rewrite cannot preserve per-range password hashes; fail closed.
        if (!features.protectedRangeEdits && source.some((range) => range.hasPassword))
          throw new Error('protection:range_password')
        features.protectedRangeEdits = op.ranges.map(({ name, sqref }) => ({ name, sqref }))
      } else if (op.op === 'protect_sheet') {
        features.protectionEdits = await sheetProtectionEdit(features, op, intent.transactionId)
      } else if (op.op === 'set_page_setup') {
        const { op: _op, sheetId: _id, ...patch } = op
        const page = { ...((features.pageSetupEdits ?? {}) as Record<string, unknown>), ...patch }
        if (op.scale !== undefined) {
          delete page.fitToWidth
          delete page.fitToHeight
        }
        if (op.fitToWidth !== undefined || op.fitToHeight !== undefined) delete page.scale
        features.pageSetupEdits = page
      } else if (op.op === 'set_freeze')
        features.freezeEdits = { frozenRows: op.rows, frozenColumns: op.columns }
      else if (op.op === 'set_hyperlink')
        features.hyperlinkEdits = {
          ...((features.hyperlinkEdits ?? {}) as Record<string, string | null>),
          [op.address]: op.target,
        }
      else {
        const metadata = features.originalMetadata as
          { comments?: { image?: unknown }[] } | undefined
        if (metadata?.comments?.some((note) => note.image))
          throw new Error('unsupported_note_image: notes with images must be preserved')
        features.noteEdits = {
          ...((features.noteEdits ?? {}) as Record<string, string | null>),
          [op.address]: op.text,
        }
      }
      featureEdits.set(op.sheetId, features)
      next = structuredClone(book)
    } else if (op.op === 'relocate_range') {
      if (current.sheets.some((sheet) => !material.completeSheetIds.includes(sheet.id)))
        throw new Error('incomplete_move_material')
      const features = new Map(
        current.sheets.map((sheet) => [
          sheet.id,
          structuredClone(featureEdits.get(sheet.id) ?? sheet.features ?? {}),
        ]),
      )
      const moved = moveSessionRange(
        book,
        op,
        typedTemplates,
        features,
        material.writableSheetIds,
        signal,
      )
      const defined = sessionDefinedNames({ ...current, features: workbookFeatures })
      if (defined.preserveNames.length) throw new Error('unsupported_move_feature: definedNames')
      for (const entry of defined.names) {
        const owner =
          current.sheets.find((sheet) => sheet.id === entry.scopeSheetId) ?? current.sheets[0]!
        if (
          moveFormulaRefs(entry.formula, moved.spec, owner.name) !== entry.formula ||
          (entry.scopeSheetId !== undefined && op.sourceSheetId !== op.sheetId)
        )
          throw new Error('unsupported_move_feature: definedNames')
      }
      for (const visual of nativeVisuals)
        if (visual.kind === 'chart' && visual.chart)
          for (const series of visual.chart.series)
            for (const key of ['valuesRef', 'categoriesRef', 'nameRef'] as const)
              if (series[key] && moveFormulaRefs(series[key]!, moved.spec, '') !== series[key])
                throw new Error('unsupported_move_feature: native_chart_reference')
      for (const [id, data] of features) featureEdits.set(id, data)
      expandedWrites += moved.writes
      if (expandedWrites > 2000)
        throw new Error('expanded_write_limit: ordinary batches allow 2000 cells')
      affected.push(...moved.affected)
      next = moved.book
      const size = dimensions.get(op.sheetId)!
      size.rows = Math.max(size.rows, moved.target.endRow + 1)
      size.columns = Math.max(size.columns, moved.target.endColumn + 1)
    } else if (op.op === 'paste_cells') {
      const origin = parseAddress(op.start),
        sheet = book.sheets.find((s) => s.id === op.sheetId)!
      const cells = { ...sheet.cells },
        styles = { ...sheet.styles },
        templates = typedTemplates.get(sheet.id) ?? {}
      for (const [r, row] of op.cells.entries())
        for (const [c, input] of row.entries()) {
          const address = formatAddress(origin.row + r, origin.column + c)
          if (op.skipCells?.includes(address)) continue
          const prior = templates[address] ?? { kind: 'blank' as const, rawValue: null }
          const actual = sheet.cells[address] ?? { value: null }
          const oldStyle = typedStyle(sheet.styles?.[address], prior.style)
          const currentCell: TypedCell = {
            ...prior,
            rawValue: actual.value,
            kind:
              prior.rawValue === actual.value
                ? prior.kind
                : actual.value === null
                  ? 'blank'
                  : typeof actual.value === 'number'
                    ? 'number'
                    : typeof actual.value === 'boolean'
                      ? 'boolean'
                      : 'string',
            ...(oldStyle ? { style: oldStyle } : {}),
          }
          if (actual.formula) currentCell.formula = actual.formula
          else delete currentCell.formula
          if (oldStyle?.numberFormat) currentCell.numberFormat = oldStyle.numberFormat
          const value = pasteCellFields(currentCell, input, op)
          if (
            op.writeValues !== false &&
            input.kind === 'date' &&
            input.dateSystem !== (before.dateSystem ?? '1900')
          )
            throw new Error('date_system_mismatch')
          if (op.writeValues !== false) {
            if (value.formula) {
              value.kind = 'blank'
              value.rawValue = null
              value.calculationState = 'pending'
              delete value.cachedValue
            }
            cells[address] = {
              value: value.rawValue,
              ...(value.formula ? { formula: value.formula } : {}),
              ...(value.kind === 'error' ? { error: true } : {}),
            }
          }
          styles[address] = formatState(value)
          templates[address] = value
        }
      if (op.merges !== undefined) {
        const area = {
          startRow: origin.row,
          startColumn: origin.column,
          endRow: origin.row + op.cells.length - 1,
          endColumn: origin.column + op.cells[0]!.length - 1,
        }
        const features = {
          ...(featureEdits.get(sheet.id) ??
            current.sheets.find((s) => s.id === sheet.id)!.features ??
            {}),
        }
        features.mergeEdits = sessionMerges(features).filter(
          (m) =>
            !(
              m.startRow <= area.endRow &&
              m.endRow >= area.startRow &&
              m.startColumn <= area.endColumn &&
              m.endColumn >= area.startColumn
            ),
        )
        for (const range of op.merges)
          features.mergeEdits = updateSessionMerges(
            features,
            { op: 'merge_cells', sheetId: sheet.id, range },
            templates,
          )
        featureEdits.set(sheet.id, features)
      }
      typedTemplates.set(sheet.id, templates)
      expandedWrites += op.cells.length * op.cells[0]!.length
      if (expandedWrites > 2000)
        throw new Error('expanded_write_limit: ordinary batches allow 2000 cells')
      next = {
        ...book,
        sheets: book.sheets.map((s) => (s.id === sheet.id ? { ...s, cells, styles } : s)),
      }
      const size = dimensions.get(sheet.id)!
      size.rows = Math.max(size.rows, origin.row + op.cells.length)
      size.columns = Math.max(size.columns, origin.column + op.cells[0]!.length)
    } else if (op.op === 'copy_range' || op.op === 'fill_range') {
      const copied = copySessionRange(book, op, typedTemplates, signal)
      next = copied.snapshot
      const size = dimensions.get(op.sheetId)!
      size.rows = Math.max(size.rows, copied.target.endRow + 1)
      size.columns = Math.max(size.columns, copied.target.endColumn + 1)
    } else {
      const sortVisible =
        op.op === 'sort_range'
          ? sortRowVisibility(
              featureEdits.get(op.sheetId) ?? permissions?.features ?? {},
              material.hiddenRows?.[op.sheetId],
            )
          : undefined
      const adapter = new InMemoryWorkbookAdapter(book, {
        sortLocale: SESSION_SORT_LOCALE,
        ...(sortVisible ? { sortRowVisible: sortVisible } : {}),
      })
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
      if (op.op === 'format_range' && op.reset) {
        const range = parseRange(op.range),
          templates = typedTemplates.get(op.sheetId)!
        for (let row = range.startRow; row <= range.endRow; row++)
          for (let column = range.startColumn; column <= range.endColumn; column++) {
            const address = formatAddress(row, column)
            const template = templates[address] ?? { kind: 'blank' as const, rawValue: null }
            templates[address] = { ...template, style: {}, styleReset: true }
            delete templates[address]!.numberFormat
          }
      }
      if (op.op === 'sort_range') {
        const range = parseRange(op.range),
          first = range.startRow + (op.hasHeader ? 1 : 0)
        const sourceSheet = book.sheets.find((sheet) => sheet.id === op.sheetId)!
        const targetSheet = next.sheets.find((sheet) => sheet.id === op.sheetId)!
        const rowIndices = Array.from(
          { length: range.endRow - first + 1 },
          (_, r) => first + r,
        ).filter((row) => sortVisible!(row))
        const matrix = rowIndices.map((sourceRow) =>
          Array.from(
            { length: range.endColumn - range.startColumn + 1 },
            (_, c) =>
              sourceSheet.cells[formatAddress(sourceRow, range.startColumn + c)]?.value ?? null,
          ),
        )
        const order = computeSortedRowOrder(
          matrix,
          columnIndex(op.byColumn) - range.startColumn,
          op.order === 'asc',
          SESSION_SORT_LOCALE,
          op.levels?.map((level) => ({
            keyOffset: columnIndex(level.byColumn) - range.startColumn,
            ascending: level.order === 'asc',
          })),
        )
        const templates = typedTemplates.get(op.sheetId)!,
          old = structuredClone(templates)
        const styles = { ...targetSheet.styles }
        for (const [row, source] of order.entries())
          for (let column = range.startColumn; column <= range.endColumn; column++) {
            const from = formatAddress(rowIndices[source]!, column),
              to = formatAddress(rowIndices[row]!, column)
            if (old[from]) templates[to] = structuredClone(old[from]!)
            else delete templates[to]
            if (sourceSheet.styles?.[from]) styles[to] = structuredClone(sourceSheet.styles[from]!)
            else delete styles[to]
          }
        next = {
          ...next,
          sheets: next.sheets.map((sheet) =>
            sheet.id === op.sheetId ? { ...sheet, styles } : sheet,
          ),
        }
      }
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
    if (op.op === 'add_shape' || op.op === 'add_picture') {
      next = {
        ...next,
        sheets: next.sheets.map((sheet) =>
          sheet.id === op.sheetId
            ? {
                ...sheet,
                visuals: sheet.visuals?.map((visual, at) =>
                  at === sheet.visuals!.length - 1
                    ? {
                        ...visual,
                        id: `added-${op.op === 'add_picture' ? 'image' : 'shape'}-${material.allocation.stateId}-${index}`,
                      }
                    : visual,
                ),
              }
            : sheet,
        ),
      }
    }
    if (op.op === 'set_row_height')
      dimensions.get(op.sheetId)!.rows = Math.max(
        dimensions.get(op.sheetId)!.rows,
        op.row + op.count - 1,
      )
    if (op.op === 'set_col_width')
      dimensions.get(op.sheetId)!.columns = Math.max(
        dimensions.get(op.sheetId)!.columns,
        columnIndex(op.column) + op.count,
      )
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
      if (spec?.axis === 'column' && 'sheetId' in op) {
        const features =
          featureEdits.get(op.sheetId) ??
          structuredClone(current.sheets.find((sheet) => sheet.id === op.sheetId)!.features ?? {})
        const widths = features.columnWidthsCharacters as Record<string, number> | undefined
        if (widths) {
          const shifted: Record<string, number> = {}
          for (const [column, value] of Object.entries(widths)) {
            const nextColumn = shiftIndex(columnIndex(column), spec)
            if (nextColumn !== null) shifted[formatAddress(0, nextColumn).replace(/1$/, '')] = value
          }
          features.columnWidthsCharacters = shifted
          featureEdits.set(op.sheetId, features)
        }
      }
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
            for (const series of visual.kind === 'chart' ? visual.chart.series : [])
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
          for (const series of visual.kind === 'chart' ? visual.chart.series : [])
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
    if (op.op === 'set_col_width') {
      const features =
        featureEdits.get(op.sheetId) ??
        structuredClone(current.sheets.find((sheet) => sheet.id === op.sheetId)!.features ?? {})
      const widths = { ...((features.columnWidthsCharacters ?? {}) as Record<string, number>) }
      for (let index = 0; index < op.count; index++) {
        const column = formatAddress(0, columnIndex(op.column) + index).replace(/1$/, '')
        widths[column] =
          op.widthCharacters ??
          Math.max(
            0,
            Math.round(
              (op.widthPx /
                measureNormalFontMdw(
                  (current.features?.originalMetadata ?? {}) as NormalFontDescriptor,
                )) *
                256,
            ) / 256,
          )
      }
      if (
        Object.values(widths).some((width) => width > 255 || width < 0 || !Number.isFinite(width))
      )
        throw new Error('column_width_out_of_range: Excel supports 0 to 255 characters')
      features.columnWidthsCharacters = widths
      featureEdits.set(op.sheetId, features)
    }
    if (op.op !== 'set_page_setup')
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
  // Keep journaled original-chart caches aligned with their current references.
  for (const owner of intent.operations.every((op) =>
    [
      'set_page_setup',
      'add_defined_name',
      'delete_defined_name',
      'manage_defined_name',
      'add_defined_names',
    ].includes(op.op),
  )
    ? []
    : book.sheets) {
    const features =
      featureEdits.get(owner.id) ??
      structuredClone(current.sheets.find((sheet) => sheet.id === owner.id)?.features ?? {})
    const edits = (features.chartEdits ?? {}) as Record<
      string,
      ChartStateEdit & { chartPath?: string }
    >
    let refreshed = false
    for (const [path, edit] of Object.entries(edits)) {
      if (!edit.series?.length && !edit.seriesSet?.length) continue
      const native = nativeVisuals.find(
        (visual) =>
          visual.chartPath === path &&
          visual.sheetId === current.sheets.find((sheet) => sheet.id === owner.id)?.originalSheetId,
      )
      if (
        !native?.chart ||
        (features.visualEdits as Record<string, { remove?: true }> | undefined)?.[native.id]?.remove
      )
        continue
      const entries = edit.seriesSet
        ? applyChartStateEdit(native.chart, edit).series.map((entry, index) => ({
            ...entry,
            index,
          }))
        : (edit.series ?? [])
      const updates = entries.map((entry) => {
        if (!entry.valuesRef && !entry.categoriesRef) return entry
        for (const ref of [entry.valuesRef, entry.categoriesRef]) {
          if (!ref) continue
          const resolved = resolveChartRange(ref, book.sheets, owner.id)
          const covered =
            material.completeSheetIds.includes(resolved.sheetId) ||
            (material.referenceRanges?.[resolved.sheetId] ?? []).some((area) => {
              const bounds = parseRange(area)
              return (
                bounds.startRow <= resolved.bounds.startRow &&
                bounds.endRow >= resolved.bounds.endRow &&
                bounds.startColumn <= resolved.bounds.startColumn &&
                bounds.endColumn >= resolved.bounds.endColumn
              )
            })
          if (!covered) throw new Error('incomplete_chart_material')
        }
        const data = chartSeriesRangeEdits(
          book,
          owner.id,
          edit.seriesSet?.length ?? native.chart!.series.length,
          [
            {
              index: entry.index,
              ...(entry.valuesRef ? { valuesRange: entry.valuesRef } : {}),
              ...(entry.categoriesRef ? { categoriesRange: entry.categoriesRef } : {}),
            },
          ],
          (sheetId, address) => {
            const cell = current.sheets.find((sheet) => sheet.id === sheetId)?.cells[address]
            if (
              book.sheets.find((sheet) => sheet.id === sheetId)?.cells[address]?.formula &&
              cell?.calculationState !== 'ready' &&
              cell?.calculationState !== 'error'
            )
              throw new Error('chart_formula_pending')
          },
        )[0]!
        return { ...entry, ...data }
      })
      edits[path] = workbookChartEditSchema.parse({
        ...edit,
        chartPath: path,
        ...(edit.seriesSet
          ? {
              series: undefined,
              seriesSet: updates.map((series) => ({
                name: series.name,
                values: series.values,
                categories: series.categories,
                valuesRef: series.valuesRef,
                categoriesRef: series.categoriesRef,
                color: 'color' in series ? series.color : undefined,
              })),
            }
          : { series: updates }),
      })
      refreshed = true
    }
    if (refreshed) {
      features.chartEdits = edits
      featureEdits.set(owner.id, features)
    }
  }
  const after: CandidateState = {
    ...before,
    stateId: material.allocation.stateId,
    structuralLog: log,
    sheetMap,
    ...(Object.keys(workbookFeatures).length ? { features: workbookFeatures } : {}),
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
          ...(kind === 'date'
            ? { dateSystem: before.dateSystem ?? template?.dateSystem ?? '1900' }
            : {}),
          ...(cell.formula
            ? {
                formula: cell.formula,
                ...(cell.value === null ? {} : { cachedValue: cell.value }),
                calculationState: 'pending',
              }
            : {}),
          ...(style ? { style } : {}),
          ...(template?.styleReset ? { styleReset: true } : {}),
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
          (featureEdits.get(sheet.id)?.protectionEdits as { protected: boolean } | undefined)
            ?.protected ??
          current.sheets.find((item) => item.id === sheet.id)?.protected ??
          original?.protected ??
          false,
        cells,
        features: JSON.parse(
          JSON.stringify({
            ...original?.features,
            ...featureEdits.get(sheet.id),
            visuals: sheet.visuals ?? [],
            charts: (sheet.visuals ?? []).filter((v) => v.kind === 'chart'),
            rowHeights: sheet.rowHeights ?? {},
            colWidths: sheet.colWidths ?? {},
          }),
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

type PivotArea = { startRow: number; endRow: number; startColumn: number; endColumn: number }

/** Pending session pivots, by target sheet (written as native parts on save). */
export function sessionPivotAdditions(state: { sheets: readonly { features?: unknown }[] }) {
  return state.sheets.flatMap(
    (sheet) =>
      ((sheet.features as { pivotAdditions?: unknown[] } | undefined)?.pivotAdditions ??
        []) as import('../shared/workbook-protocol').WorkbookPivotAdd[],
  )
}

type PivotLayoutSpec = Pick<
  Extract<import('./workbook-dsl').WorkbookOperation, { op: 'add_pivot' }>,
  'rowFields' | 'columnField' | 'pageFields' | 'values' | 'groupings' | 'filters' | 'subtotalLabel'
>

/**
 * Lays out a PivotTable block at targetCell from current source cells, like
 * the desktop's applyAiPivotAdd. An edit passes the pivot's previous block:
 * it is excluded from overlap checks, newly occupied cells must be empty and
 * the union of old and new blocks is rewritten so a shrunken layout clears.
 */
function placeSessionPivot(input: {
  current: ReturnType<typeof workbookStateSchema.parse>
  material: CandidateMaterial
  sourceSheetId: string
  targetSheetId: string
  sourceRange: string
  targetCell: string
  spec: PivotLayoutSpec
  previous?: { block: PivotArea; sessionName?: string }
}) {
  const { current, material, previous } = input
  const sourceSheet = current.sheets.find((sheet) => sheet.id === input.sourceSheetId),
    targetSheet = current.sheets.find((sheet) => sheet.id === input.targetSheetId)
  if (!sourceSheet || !targetSheet) throw new Error('unknown_sheet')
  if (!sourceSheet.originalSheetId || !targetSheet.originalSheetId)
    throw new Error('pivot_layout:onAddedSheet')
  const source = parseRange(input.sourceRange)
  const sourceRows = source.endRow - source.startRow + 1,
    sourceColumns = source.endColumn - source.startColumn + 1
  if (sourceRows < 2) throw new Error('pivot_layout:sourceNeedsRows')
  if (sourceRows > PIVOT_SOURCE_ROW_LIMIT) throw new Error('pivot_layout:sourceRowLimit')
  if (sourceColumns > PIVOT_SOURCE_COL_LIMIT) throw new Error('pivot_layout:sourceColLimit')
  if (!material.completeSheetIds.includes(sourceSheet.id) && sourceRows * sourceColumns > 200_000)
    throw new Error('incomplete_pivot_source')
  const cells = []
  for (let row = source.startRow; row <= source.endRow; row++) {
    const line = []
    for (let column = source.startColumn; column <= source.endColumn; column++) {
      const cell = sourceSheet.cells[formatAddress(row, column)]
      if (
        cell?.formula &&
        (cell.calculationState !== 'ready' || !Object.hasOwn(cell, 'cachedValue'))
      )
        throw new Error('Wait for PivotTable source formulas to finish calculating, then retry.')
      const value = cell?.formula ? cell.cachedValue : cell?.rawValue
      if (cell?.kind === 'error' || (cell?.formula && cell.calculationState === 'error'))
        throw new Error('pivot_layout:sourceError')
      line.push({
        value: (value ?? null) as string | number | boolean | null,
        pattern: cell?.numberFormat ?? cell?.style?.numberFormat,
        static1904: current.dateSystem === '1904' && !cell?.formula,
      })
    }
    cells.push(line)
  }
  const grid = pivotSourceGrid(
    cells,
    input.spec.values.filter((value) => value.formula === undefined).map((value) => value.field),
  )
  let layout: PivotLayout
  try {
    layout = buildPivotLayout(grid, input.spec, {
      subtotal: input.spec.subtotalLabel ?? 'Subtotal',
      grandTotal: 'Grand Total',
    })
  } catch (error) {
    if (error instanceof PivotLayoutError)
      throw new Error(`pivot_layout:${error.code}:${JSON.stringify(error.params)}`, {
        cause: error,
      })
    throw error
  }
  const matrix = pivotBlockMatrix(layout),
    pageRows = pivotPageRows(layout.definition.pageFieldIndices?.length ?? 0)
  const blockAnchor = parseAddress(input.targetCell)
  const location = pivotOutputArea(
    { row: blockAnchor.row + pageRows, column: blockAnchor.column },
    layout,
  )
  const block: PivotArea = { ...location, startRow: blockAnchor.row }
  if (block.endRow >= MAX_GRID_ROWS || block.endColumn >= MAX_GRID_COLUMNS)
    throw new Error('The PivotTable would extend past the worksheet grid.')
  if (matrix.length * layout.width > 50_000)
    throw new Error('The PivotTable output exceeds 50,000 cells.')
  if (input.targetSheetId === input.sourceSheetId && areasOverlap(block, source))
    throw new Error('pivot_layout:overlapSource')
  // A file pivot's recorded range may predate refresh growth: match it by anchor column.
  const self = (area: PivotArea) =>
    previous !== undefined &&
    previous.sessionName === undefined &&
    area.startColumn === previous.block.startColumn &&
    areasOverlap(area, previous.block)
  const originalPivots = ((
    targetSheet.features?.originalMetadata as { pivotRanges?: PivotArea[] } | undefined
  )?.pivotRanges ?? []) as PivotArea[]
  if (originalPivots.some((area) => !self(area) && areasOverlap(block, area)))
    throw new Error('pivot_layout:overlapExisting')
  for (const pivot of sessionPivotAdditions(current))
    if (
      pivot.sheetId === input.targetSheetId &&
      pivot.name !== previous?.sessionName &&
      areasOverlap(block, pivot.location)
    )
      throw new Error(`pivot_layout:overlapSession:${JSON.stringify({ name: pivot.name })}`)
  let written = matrix,
    origin = { row: blockAnchor.row, column: blockAnchor.column }
  if (previous) {
    // Grown parts must be empty; the union of old and new blocks is rewritten.
    const old = previous.block
    for (let row = block.startRow; row <= block.endRow; row++)
      for (let column = block.startColumn; column <= block.endColumn; column++) {
        if (
          row >= old.startRow &&
          row <= old.endRow &&
          column >= old.startColumn &&
          column <= old.endColumn
        )
          continue
        const cell = targetSheet.cells[formatAddress(row, column)]
        if (cell && (cell.formula || (cell.rawValue !== null && cell.rawValue !== '')))
          throw new Error('pivot_layout:relayoutOverlap')
      }
    const union = {
      startRow: Math.min(old.startRow, block.startRow),
      startColumn: Math.min(old.startColumn, block.startColumn),
      endRow: Math.max(old.endRow, block.endRow),
      endColumn: Math.max(old.endColumn, block.endColumn),
    }
    written = Array.from({ length: union.endRow - union.startRow + 1 }, (_, r) =>
      Array.from(
        { length: union.endColumn - union.startColumn + 1 },
        (_, c) =>
          matrix[union.startRow + r - block.startRow]?.[
            union.startColumn + c - block.startColumn
          ] ?? null,
      ),
    )
    origin = { row: union.startRow, column: union.startColumn }
  }
  const writes = matrixWrites(input.targetSheetId, origin.row, origin.column, written)
  // Value columns' numFmt covers the data rows only (no header, no grand total).
  const dataRows = layout.height - 2
  for (const { columnOffset, format } of layout.numberFormats)
    for (let row = 1; row <= dataRows; row += 2000)
      writes.push({
        op: 'format_range',
        sheetId: input.targetSheetId,
        range: `${formatAddress(location.startRow + row, location.startColumn + columnOffset)}:${formatAddress(location.startRow + Math.min(dataRows, row + 1999), location.startColumn + columnOffset)}`,
        format: { numberFormat: format },
      })
  return { writes, layout, source, location, block }
}

/**
 * Create PivotTable: bakes the computed block like the desktop and records the
 * native definition for the package writer. Private write stages stay within
 * the ordinary batch; only the final root is published.
 */
async function addSessionPivot(
  before: Parameters<typeof stateChangeset>[0],
  current: ReturnType<typeof workbookStateSchema.parse>,
  intent: ReturnType<typeof transactionInputSchema.parse>,
  op: Extract<import('./workbook-dsl').WorkbookOperation, { op: 'add_pivot' }>,
  material: CandidateMaterial,
  options: Parameters<typeof computeCandidateState>[3],
) {
  const targetId = op.targetSheetId ?? op.sheetId
  const placed = placeSessionPivot({
    current,
    material,
    sourceSheetId: op.sheetId,
    targetSheetId: targetId,
    sourceRange: op.sourceRange,
    targetCell: op.targetCell,
    spec: op,
  })
  const pending = sessionPivotAdditions(current)
  let name = op.name
  if (name === undefined) {
    const taken = new Set(pending.map((pivot) => pivot.name.toLowerCase()))
    let index = pending.length + 1
    while (taken.has(`pivot${index}`)) index++
    name = `Pivot${index}`
  } else if (pending.some((pivot) => pivot.name.toLowerCase() === name!.toLowerCase()))
    throw new Error(`pivot_layout:nameUsed:${JSON.stringify({ name })}`)
  const written = await applyBoundedWrites(current, intent, placed.writes, material, options)
  const addition = workbookPivotAddSchema.parse({
    sheetId: targetId,
    sourceSheetId: op.sheetId,
    sourceArea: placed.source,
    location: placed.location,
    name,
    ...placed.layout.definition,
  })
  const next = structuredClone(written.state)
  const target = next.sheets.find((sheet) => sheet.id === targetId)!
  target.features = {
    ...target.features,
    pivotAdditions: [
      ...((target.features?.pivotAdditions ?? []) as unknown[]),
      JSON.parse(JSON.stringify(addition)),
    ],
  }
  const { block } = placed
  return stateChangeset(before, next, [
    ...written.affected,
    `${targetId}!${formatAddress(block.startRow, block.startColumn)}:${formatAddress(block.endRow, block.endColumn)}`,
  ])
}

function areaRef(area: PivotArea) {
  return `${formatAddress(area.startRow, area.startColumn)}:${formatAddress(area.endRow, area.endColumn)}`
}

/** File pivots with a pending session relayout: definition is stale until save. */
function fileRelayouts(sheet: { features?: unknown }) {
  return ((sheet.features as { pivotRelayouts?: unknown } | undefined)?.pivotRelayouts ??
    {}) as Record<
    string,
    import('../shared/workbook-protocol').WorkbookPivotAdd & { cachePath: string }
  >
}

/**
 * Edit PivotTable (Fields pane): relayouts a session pivot (its pending
 * addition is replaced) or a file pivot (written on save as a regenerated
 * definition with refreshOnLoad), keeping the block's top-left anchor.
 */
async function editSessionPivot(
  before: Parameters<typeof stateChangeset>[0],
  current: ReturnType<typeof workbookStateSchema.parse>,
  intent: ReturnType<typeof transactionInputSchema.parse>,
  op: Extract<import('./workbook-dsl').WorkbookOperation, { op: 'edit_pivot' }>,
  material: CandidateMaterial,
  options: Parameters<typeof computeCandidateState>[3],
) {
  const sheet = current.sheets.find((item) => item.id === op.sheetId)
  if (!sheet) throw new Error('unknown_sheet')
  if (!material.completeSheetIds.includes(sheet.id)) throw new Error('incomplete_pivot_material')
  let sourceSheetId: string, sourceRange: string, body: PivotArea, pageRows: number
  let sessionName: string | undefined, file: { path: string; cachePath: string } | undefined
  if ('session' in op.pivot) {
    const add = sessionPivotAdditions(current).find(
      (pivot) =>
        pivot.sheetId === sheet.id && pivot.name === (op.pivot as { session: string }).session,
    )
    if (!add) throw new Error('pivot_layout:pivotStale')
    sessionName = add.name
    sourceSheetId = add.sourceSheetId
    sourceRange = areaRef(add.sourceArea)
    body = add.location
    pageRows = pivotPageRows(add.pageFieldIndices?.length ?? 0)
  } else {
    const path = op.pivot.file
    const relayout = fileRelayouts(sheet)[path]
    const pivot = filePivotsOfSheet(sheet).find((item) => item.path === path)
    if (!pivot) throw new Error('pivot_layout:pivotStale')
    if (relayout) {
      file = { path, cachePath: relayout.cachePath }
      sourceSheetId = relayout.sourceSheetId
      sourceRange = areaRef(relayout.sourceArea)
      body = relayout.location
      pageRows = pivotPageRows(relayout.pageFieldIndices?.length ?? 0)
    } else {
      const definition = pivot.definition
      if (!definition || !pivot.cachePath) throw new Error('pivot_layout:defNotLoaded')
      if (definition.unsupported.length)
        throw new Error(
          `pivot_layout:editUnsupported:${JSON.stringify({ reasons: definition.unsupported.join('; ') })}`,
        )
      const source = pivotSourceSheet(current, definition.sourceSheet)
      if (!source)
        throw new Error(
          `pivot_layout:sourceSheetMissing:${JSON.stringify({ name: definition.sourceSheet })}`,
        )
      file = { path, cachePath: pivot.cachePath }
      sourceSheetId = source.id
      sourceRange = definition.sourceRef
      body = parseRange(definition.outputRef)
      pageRows = Math.min(body.startRow, pivotPageRows(definition.pageFields.length))
    }
  }
  const oldBlock = { ...body, startRow: body.startRow - pageRows }
  const placed = placeSessionPivot({
    current,
    material,
    sourceSheetId,
    targetSheetId: sheet.id,
    sourceRange,
    targetCell: formatAddress(oldBlock.startRow, oldBlock.startColumn),
    spec: op,
    previous: { block: oldBlock, ...(sessionName === undefined ? {} : { sessionName }) },
  })
  const written = await applyBoundedWrites(current, intent, placed.writes, material, options)
  const addition = workbookPivotAddSchema.parse({
    sheetId: sheet.id,
    sourceSheetId,
    sourceArea: placed.source,
    location: placed.location,
    name: sessionName ?? 'PivotEdit',
    ...placed.layout.definition,
  })
  const next = structuredClone(written.state)
  const target = next.sheets.find((item) => item.id === sheet.id)!
  const features = (target.features ?? {}) as {
    pivotAdditions?: { name: string }[]
    pivotRelayouts?: Record<string, unknown>
    pivotDefinitionEdits?: Record<string, unknown>
    pivotRefresh?: { cachePaths: string[]; outputRefs: Record<string, string> }
  }
  if (sessionName !== undefined)
    target.features = {
      ...target.features,
      pivotAdditions: (features.pivotAdditions ?? []).map((pivot) =>
        pivot.name === sessionName ? JSON.parse(JSON.stringify(addition)) : pivot,
      ),
    }
  else {
    const { path, cachePath } = file!
    const { [path]: _stale, ...definitions } = features.pivotDefinitionEdits ?? {}
    target.features = {
      ...target.features,
      pivotDefinitionEdits: JSON.parse(JSON.stringify(definitions)),
      pivotRelayouts: {
        ...features.pivotRelayouts,
        [path]: JSON.parse(JSON.stringify({ ...addition, cachePath })),
      },
      pivotRefresh: {
        cachePaths: [...new Set([...(features.pivotRefresh?.cachePaths ?? []), cachePath])],
        outputRefs: { ...features.pivotRefresh?.outputRefs, [cachePath]: areaRef(placed.location) },
      },
    }
  }
  return stateChangeset(before, next, [...written.affected, `${sheet.id}!${areaRef(placed.block)}`])
}

/** Writes in private ordinary-batch stages (at most 2,000 cells each); only the final root is used. */
async function applyBoundedWrites(
  current: ReturnType<typeof workbookStateSchema.parse>,
  intent: ReturnType<typeof transactionInputSchema.parse>,
  writes: readonly import('./workbook-dsl').WorkbookOperation[],
  material: CandidateMaterial,
  options: Parameters<typeof computeCandidateState>[3],
) {
  let state = current
  const affected: string[] = []
  let stageOps: import('./workbook-dsl').WorkbookOperation[] = [],
    budget = 0
  const flush = async () => {
    if (!stageOps.length) return
    const stage = await computeCandidateState(
      state,
      { ...intent, operations: stageOps },
      { ...material, state },
      options,
    )
    affected.push(...stage.affectedTargets)
    state = stage.nextState
    stageOps = []
    budget = 0
  }
  for (const write of writes) {
    const cells =
      write.op === 'set_range'
        ? write.values.length * write.values[0]!.length
        : write.op === 'format_range'
          ? rangeCellCount(parseRange(write.range))
          : 1
    if (budget + cells > 2000) await flush()
    stageOps.push(write)
    budget += cells
  }
  await flush()
  return { state, affected }
}

function matrixWrites(
  sheetId: string,
  startRow: number,
  startColumn: number,
  matrix: readonly (readonly (string | number | boolean | null)[])[],
) {
  const writes: import('./workbook-dsl').WorkbookOperation[] = []
  const width = matrix[0]?.length ?? 0
  for (let column = 0; column < width; column += 100) {
    const span = Math.min(100, width - column),
      band = Math.max(1, Math.min(500, Math.floor(2000 / span)))
    for (let row = 0; row < matrix.length; row += band)
      writes.push({
        op: 'set_range',
        sheetId,
        start: formatAddress(startRow + row, startColumn + column),
        values: matrix.slice(row, row + band).map((line) => line.slice(column, column + span)),
      })
  }
  return writes
}

export interface FilePivot {
  readonly path: string
  readonly cachePath?: string
  readonly definition?: import('../shared/workbook-protocol').WorkbookPivotDefinition
  readonly error?: string
}

/** A sheet's file PivotTables, with definitions grown by earlier session refreshes. */
export function filePivotsOfSheet(sheet: { features?: unknown }): FilePivot[] {
  const features = (sheet.features ?? {}) as {
    originalMetadata?: { pivotDefinitions?: FilePivot[]; pivotTables?: { path: string }[] }
    pivotDefinitionEdits?: Record<string, FilePivot['definition']>
  }
  const parsed = features.originalMetadata?.pivotDefinitions ?? []
  return (features.originalMetadata?.pivotTables ?? []).map(({ path }) => {
    const entry = parsed.find((item) => item.path === path) ?? { path, error: 'not_loaded' }
    if (fileRelayouts(sheet)[path]) return { path, error: 'relayout_pending' }
    const edited = features.pivotDefinitionEdits?.[path]
    return edited ? { ...entry, definition: edited } : entry
  })
}

/** The current sheet holding a file pivot's source (the definition names the original sheet). */
export function pivotSourceSheet<
  T extends { sheets: readonly { id: string; name: string }[]; sheetMap?: unknown },
>(state: T, sourceSheet: string): T['sheets'][number] | undefined {
  const map = (state.sheetMap ?? {}) as Record<string, { originalName?: string }>
  return state.sheets.find((sheet) => (map[sheet.id]?.originalName ?? sheet.name) === sourceSheet)
}

/**
 * Data → Refresh / Refresh All for file PivotTables, as the desktop does:
 * recompute the data area from fresh source values, or fold new members into
 * the layout and rewrite the widened output. Saves set refreshOnLoad and the
 * new location; the cache records are rebuilt by Excel.
 */
async function refreshSessionPivots(
  before: Parameters<typeof stateChangeset>[0],
  current: ReturnType<typeof workbookStateSchema.parse>,
  intent: ReturnType<typeof transactionInputSchema.parse>,
  ops: readonly Extract<import('./workbook-dsl').WorkbookOperation, { op: 'refresh_pivot' }>[],
  material: CandidateMaterial,
  options: Parameters<typeof computeCandidateState>[3],
) {
  const writes: import('./workbook-dsl').WorkbookOperation[] = []
  const edits = new Map<
    string,
    {
      definitions: Record<string, unknown>
      cachePaths: Set<string>
      outputRefs: Record<string, string>
    }
  >()
  let refreshed = 0
  for (const op of ops) {
    const sheet = current.sheets.find((item) => item.id === op.sheetId)
    if (!sheet) throw new Error('unknown_sheet')
    const pivots = filePivotsOfSheet(sheet)
    if (!pivots.length) throw new Error('pivot_layout:sheetNoPivot')
    if (!material.completeSheetIds.includes(sheet.id)) throw new Error('incomplete_pivot_material')
    const edit = edits.get(sheet.id) ?? {
      definitions: {},
      cachePaths: new Set<string>(),
      outputRefs: {},
    }
    edits.set(sheet.id, edit)
    for (const pivot of pivots) {
      const definition = pivot.definition
      if (pivot.error === 'relayout_pending') throw new Error('pivot_layout:relayoutPending')
      if (!definition || !pivot.cachePath) throw new Error('pivot_layout:defNotLoaded')
      const sourceSheet = pivotSourceSheet(current, definition.sourceSheet)
      if (!sourceSheet)
        throw new Error(
          `pivot_layout:sourceSheetMissing:${JSON.stringify({ name: definition.sourceSheet })}`,
        )
      const source = parseRange(definition.sourceRef)
      const cells = []
      for (let row = source.startRow; row <= source.endRow; row++) {
        const line = []
        for (let column = source.startColumn; column <= source.endColumn; column++) {
          const cell = sourceSheet.cells[formatAddress(row, column)]
          if (
            cell?.formula &&
            (cell.calculationState !== 'ready' || !Object.hasOwn(cell, 'cachedValue'))
          )
            throw new Error(
              'Wait for PivotTable source formulas to finish calculating, then retry.',
            )
          line.push({
            value: ((cell?.formula ? cell.cachedValue : cell?.rawValue) ?? null) as
              string | number | boolean | null,
            pattern: cell?.numberFormat ?? cell?.style?.numberFormat,
            static1904: current.dateSystem === '1904' && !cell?.formula,
          })
        }
        cells.push(line)
      }
      const grid = pivotSourceGrid(
        cells,
        definition.dataFields.flatMap(({ field }) => {
          const owner = definition.fields[field]
          return owner && owner.formula === undefined ? [owner.name] : []
        }),
      )
      const growth = growPivotDefinition(definition, grid)
      const { data } = recomputePivotData(growth.definition, grid)
      const bounds = parseRange(definition.outputRef)
      if (!growth.grown) {
        const startRow = bounds.startRow + definition.firstDataRow,
          startColumn = bounds.startColumn + definition.firstDataCol
        if (
          startRow + data.length - 1 !== bounds.endRow ||
          startColumn + (data[0]?.length ?? 0) - 1 !== bounds.endColumn
        )
          throw new Error('pivot_layout:layoutMismatch')
        if (data.length) writes.push(...matrixWrites(sheet.id, startRow, startColumn, data))
      } else {
        let output: ReturnType<typeof grownPivotBlock>
        try {
          output = grownPivotBlock(growth.definition, data, bounds, op.subtotalLabel ?? 'Subtotal')
        } catch (error) {
          if (error instanceof PivotGrowError)
            throw new Error('pivot_layout:growUnsupported', { cause: error })
          throw error
        }
        for (const area of output.growth)
          for (let row = area.startRow; row <= area.endRow; row++)
            for (let column = area.startColumn; column <= area.endColumn; column++) {
              const cell = sheet.cells[formatAddress(row, column)]
              if (cell && (cell.formula || (cell.rawValue !== null && cell.rawValue !== '')))
                throw new Error('pivot_layout:growConflict')
            }
        writes.push(...matrixWrites(sheet.id, bounds.startRow, bounds.startColumn, output.block))
        edit.definitions[pivot.path] = { ...growth.definition, outputRef: output.outputRef }
        edit.outputRefs[pivot.cachePath] = output.outputRef
      }
      edit.cachePaths.add(pivot.cachePath)
      refreshed++
    }
  }
  if (!refreshed) throw new Error('pivot_layout:sheetNoPivot')
  const written = await applyBoundedWrites(current, intent, writes, material, options)
  const next = structuredClone(written.state)
  for (const [sheetId, edit] of edits) {
    const sheet = next.sheets.find((item) => item.id === sheetId)!
    const features = (sheet.features ?? {}) as {
      pivotDefinitionEdits?: Record<string, unknown>
      pivotRefresh?: { cachePaths: string[]; outputRefs: Record<string, string> }
    }
    sheet.features = {
      ...sheet.features,
      pivotDefinitionEdits: JSON.parse(
        JSON.stringify({ ...features.pivotDefinitionEdits, ...edit.definitions }),
      ),
      pivotRefresh: {
        cachePaths: [
          ...new Set([...(features.pivotRefresh?.cachePaths ?? []), ...edit.cachePaths]),
        ],
        outputRefs: { ...features.pivotRefresh?.outputRefs, ...edit.outputRefs },
      },
    }
  }
  return stateChangeset(before, next, [...written.affected, ...edits.keys()])
}

type ProtectionEdit = {
  protected: boolean
  hasPassword: boolean
  allow: import('../gateway/xlsx-protection').SheetProtectionAllow
  password: import('../gateway/xlsx-protection-hash').SheetPasswordHash | null
  verified: boolean
}

/** The sheet's effective protection: a session edit, else the source element. */
export function sheetProtectionOf(features: Record<string, unknown> | undefined) {
  const edit = features?.protectionEdits as ProtectionEdit | undefined
  if (edit) return edit
  const source = features?.sheetProtection as
    | {
        protected: boolean
        hasPassword: boolean
        allow?: ProtectionEdit['allow']
        password?: ProtectionEdit['password']
      }
    | null
    | undefined
  return source
    ? {
        protected: source.protected,
        hasPassword: source.hasPassword,
        allow: { ...DEFAULT_SHEET_PROTECTION_ALLOW, ...source.allow },
        password: source.password ?? null,
        verified: false,
      }
    : null
}

/**
 * Review → Protect/Unprotect Sheet. Unprotecting a password-protected sheet
 * verifies the password; protecting hashes it with a salt derived from the
 * transaction so the view candidate and host execution agree.
 */
async function sheetProtectionEdit(
  features: Record<string, unknown>,
  op: Extract<import('./workbook-dsl').WorkbookOperation, { op: 'protect_sheet' }>,
  transactionId: string,
): Promise<ProtectionEdit> {
  if (!Object.hasOwn(features, 'sheetProtection') && !features.protectionEdits)
    throw new Error('protection:not_indexed')
  const current = sheetProtectionOf(features)
  if (!op.protected && current?.protected && current.hasPassword) {
    if (!op.password) throw new Error('protection:password_required')
    if (!current.password) throw new Error('protection:password_unavailable')
    if (!(await verifySheetPassword(op.password, current.password)))
      throw new Error('protection:wrong_password')
  }
  let password: ProtectionEdit['password'] = null
  if (op.protected && op.password) {
    const digest = new Uint8Array(
      await globalThis.crypto.subtle.digest(
        'SHA-256',
        new TextEncoder().encode(`${transactionId}:${op.sheetId}`),
      ),
    )
    password = await hashSheetPassword(op.password, undefined, digest.slice(0, 16))
  } else if (op.protected && current?.protected) password = current.password
  return {
    protected: op.protected,
    hasPassword: password !== null,
    allow: Object.fromEntries(
      Object.entries(DEFAULT_SHEET_PROTECTION_ALLOW).map(([key, value]) => [
        key,
        (op.allow ?? current?.allow)?.[key as keyof ProtectionEdit['allow']] ?? value,
      ]),
    ) as ProtectionEdit['allow'],
    password,
    verified: !op.protected,
  }
}

/** The workbook structure lock: a session edit, else the file's workbookProtection. */
export function workbookStructureLocked(features: Record<string, unknown> | undefined) {
  const edit = features?.workbookProtectionEdits as { lockStructure: boolean } | undefined
  if (edit) return edit.lockStructure
  const original = (features?.originalMetadata as Record<string, unknown> | undefined)
    ?.workbookProtection as { lockStructure?: boolean } | undefined
  return original?.lockStructure === true
}

type TableArea = { startRow: number; endRow: number; startColumn: number; endColumn: number }
export interface SessionTable {
  readonly kind: 'session' | 'file'
  /// Current name (after renames).
  readonly name: string
  /// File tables: the name the save locates the part by.
  readonly originalName: string
  readonly area: TableArea
  readonly columnNames: readonly string[]
  readonly style: string | undefined
  readonly options: {
    headerRow: boolean
    totalsRow: boolean
    bandedRows: boolean
    bandedColumns: boolean
    firstColumn: boolean
    lastColumn: boolean
    filterButton: boolean
  }
  /// File table metadata (palette etc.) for projection.
  readonly file?: Record<string, unknown>
}

/** The sheet's current tables: file tables (with session edits) and tables created this session. */
export function sessionTablesOf(sheet: { features?: unknown }): SessionTable[] {
  const features = (sheet.features ?? {}) as {
    originalMetadata?: { tables?: Record<string, unknown>[] }
    tableEdits?: Record<string, Record<string, unknown>>
    tableAdditions?: import('../shared/workbook-protocol').WorkbookTableAdd[]
  }
  const result: SessionTable[] = []
  for (const table of features.originalMetadata?.tables ?? []) {
    const originalName = String(table.name ?? '')
    if (!originalName) continue
    const edit = features.tableEdits?.[originalName] ?? {}
    if (edit.remove) continue
    const range = table.range as TableArea
    result.push({
      kind: 'file',
      name: String(edit.name ?? originalName),
      originalName,
      area: (edit.area as TableArea | undefined) ?? range,
      columnNames:
        (edit.columnNames as string[] | undefined) ?? (table.columns as string[] | undefined) ?? [],
      style: (edit.style as string | undefined) ?? (table.styleName as string | undefined),
      options: {
        headerRow: (edit.headerRow as boolean | undefined) ?? Number(table.headerRowCount ?? 1) > 0,
        totalsRow: (edit.totalsRow as boolean | undefined) ?? Number(table.totalsRowCount ?? 0) > 0,
        bandedRows: (edit.bandedRows as boolean | undefined) ?? table.showRowStripes === true,
        bandedColumns:
          (edit.bandedColumns as boolean | undefined) ?? table.showColumnStripes === true,
        firstColumn: (edit.firstColumn as boolean | undefined) ?? false,
        lastColumn: (edit.lastColumn as boolean | undefined) ?? false,
        filterButton: (edit.filterButton as boolean | undefined) ?? true,
      },
      file: table,
    })
  }
  for (const add of features.tableAdditions ?? [])
    result.push({
      kind: 'session',
      name: add.name,
      originalName: add.name,
      area: add.area,
      columnNames: add.columnNames,
      style: add.style,
      options: {
        headerRow: add.headerRow ?? true,
        totalsRow: add.totalsRow ?? false,
        bandedRows: add.bandedRows,
        bandedColumns: add.bandedColumns ?? false,
        firstColumn: add.firstColumn ?? false,
        lastColumn: add.lastColumn ?? false,
        filterButton: add.filterButton ?? true,
      },
    })
  return result
}
