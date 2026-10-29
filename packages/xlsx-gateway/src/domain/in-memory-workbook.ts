import { buildChartSeriesSelection } from './chart-series-selection'
import { resolveChartRange } from './chart-series-range'
import { assertChartFormat } from './chart-format'
import { chartSeriesRangeEdits } from './chart-series-range'
import {
  copyTargetBounds,
  expandToPrimitiveOps,
  filteredCopySourceRows,
  formatOpLabel,
  isLayoutOp,
  isStructuralOp,
  layoutOpLabel,
  matchableCellText,
  MAX_EXPANDED_CELL_OPS,
  parseWorkbookCommandBatch,
  structuralOpLabel,
  type BorderPatch,
  type CellFormatPatch,
  type FillPatch,
  type FormatRangeOperation,
  type StyleColorInput,
  type LayoutOperation,
  type StructuralOperation,
} from './workbook-dsl'
import {
  fillDisplayColor,
  normalizeStyleColor,
  resolveStyleColor,
  type FillSpec,
} from './style-color'
import {
  columnIndex,
  columnLabel,
  formatAddress,
  parseAddress,
  parseRange,
  rangeAddresses,
  rangeCellCount,
} from './cell-address'
import {
  applyChartStateEdit,
  buildChartVisual,
  CHART_EDIT_TYPES,
  type ChartGridValue,
  type ChartStateEdit,
  type SheetVisual,
} from './chart-visual'
import {
  offsetFormulaRefs,
  shiftFormulaRefs,
  shiftIndex,
  shiftSpecForOp,
  type ShiftSpec,
} from './formula-shift'
import type {
  CellChange,
  CellFormatState,
  CellState,
  ChangePlan,
  CommitReceipt,
  FormatChange,
  SheetRename,
  StructuralChange,
  WorkbookAdapter,
  WorkbookSnapshot,
  WorksheetState,
} from './workbook.types'

export class WorkbookConflictError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'WorkbookConflictError'
  }
}

export interface AuditRecord {
  readonly transactionId: string
  readonly action: 'apply' | 'undo'
  readonly previousRevision: number
  readonly revision: number
  readonly changedCellCount: number
  readonly renamedSheetCount: number
  readonly structuralChangeCount: number
}

export class InMemoryWorkbookAdapter implements WorkbookAdapter {
  private snapshot: WorkbookSnapshot
  private readonly history: WorkbookSnapshot[] = []
  private readonly committedTransactions = new Set<string>()
  private readonly auditLog: AuditRecord[] = []

  constructor(
    snapshot: WorkbookSnapshot,
    private readonly options: {
      sortLocale?: string
      sortRowVisible?: (row: number, sheetId: string) => boolean
    } = {},
  ) {
    this.snapshot = structuredClone(snapshot)
  }

  getSnapshot(): WorkbookSnapshot {
    return structuredClone(this.snapshot)
  }

  getAuditLog(): readonly AuditRecord[] {
    return structuredClone(this.auditLog)
  }

  plan(input: unknown): ChangePlan {
    const batch = parseWorkbookCommandBatch(input)
    if (batch.baseRevision !== this.snapshot.revision) {
      throw new WorkbookConflictError('The workbook changed after the AI context was captured.')
    }
    if (this.committedTransactions.has(batch.transactionId)) {
      throw new WorkbookConflictError('This transaction was already committed.')
    }

    const working = structuredClone(this.snapshot)
    const cellChanges: CellChange[] = []
    const sheetRenames: SheetRename[] = []
    const structuralChanges: StructuralChange[] = []
    const formatChanges: FormatChange[] = []
    const warnings: string[] = []

    const readCell = (address: string, sheetId: string): CellState => {
      const sheet = this.snapshot.sheets.find((candidate) => candidate.id === sheetId)
      if (!sheet) throw new WorkbookConflictError(`Unknown sheet: ${sheetId}`)
      return sheet.cells[address] ?? { value: null }
    }

    for (const operation of expandToPrimitiveOps(batch.operations, readCell, this.options)) {
      if (operation.op === 'fill_range') {
        if (operation.copyFormat)
          throw new WorkbookConflictError(
            'Formatting fills require an authoritative workbook session.',
          )
        // Demo workbooks clone every cell per revision, so fills stay at the
        // per-cell preview scale; imported files get the range-level executor
        // with the 200k cap instead.
        const target = parseRange(operation.target)
        if (rangeCellCount(target) > MAX_EXPANDED_CELL_OPS) {
          throw new WorkbookConflictError(
            `fill_range on an in-memory workbook is limited to ${MAX_EXPANDED_CELL_OPS} cells per operation (imported xlsx files allow up to 200,000).`,
          )
        }
        const source = parseRange(operation.source)
        // Captured once: when source and target overlap they share their
        // top-left corner (validated at expansion), so the only source cells
        // rewritten during the loop are identity copies of themselves.
        const sourceSheet = findSheet(working, operation.sourceSheetId ?? operation.sheetId)
        const targetBefore = findSheet(working, operation.sheetId)
        const sourceRows = source.endRow - source.startRow + 1
        const sourceColumns = source.endColumn - source.startColumn + 1
        for (let row = target.startRow; row <= target.endRow; row += 1) {
          for (let column = target.startColumn; column <= target.endColumn; column += 1) {
            const sourceRow = source.startRow + ((row - target.startRow) % sourceRows)
            const sourceColumn =
              source.startColumn + ((column - target.startColumn) % sourceColumns)
            const cell = sourceSheet.cells[formatAddress(sourceRow, sourceColumn)] ?? {
              value: null,
            }
            const after: CellState = cell.formula
              ? {
                  value: null,
                  formula: offsetFormulaRefs(cell.formula, row - sourceRow, column - sourceColumn),
                }
              : { value: cell.value }
            const address = formatAddress(row, column)
            const before = targetBefore.cells[address] ?? { value: null }
            cellChanges.push({ sheetId: targetBefore.id, address, before, after })
            replaceCell(working, targetBefore.id, address, after)
          }
        }
        continue
      }
      if (operation.op === 'copy_range') {
        // Same per-cell scale as fill_range: demo workbooks clone every cell
        // per revision, so big-block copies belong to imported files.
        const target = copyTargetBounds(operation)
        if (rangeCellCount(target) > MAX_EXPANDED_CELL_OPS) {
          throw new WorkbookConflictError(
            `copy_range on an in-memory workbook is limited to ${MAX_EXPANDED_CELL_OPS} cells per operation (imported xlsx files allow up to 200,000).`,
          )
        }
        const source = parseRange(operation.source)
        const sourceSheet = findSheet(working, operation.sourceSheetId ?? operation.sheetId)
        const targetBefore = findSheet(working, operation.sheetId)
        const rowDelta = target.startRow - source.startRow
        const columnDelta = target.startColumn - source.startColumn
        // Filtered copies extract matching rows, compacted at the target, as
        // static values (the snapshot only knows written values, not what
        // formulas evaluate to — same values-only contract as the lazy path).
        const sourceRows = filteredCopySourceRows(operation, (row, column) =>
          matchableCellText(sourceSheet.cells[formatAddress(row, column)]?.value ?? null),
        )
        if (sourceRows === null) {
          throw new WorkbookConflictError(
            `copy_range filter matched no source rows — none of the ${operation.filterColumn} cells equal ${(operation.filterValues ?? []).join(', ')} (matching is trimmed and case-insensitive).`,
          )
        }
        for (const [targetOffset, row] of sourceRows.entries()) {
          for (let column = source.startColumn; column <= source.endColumn; column += 1) {
            const cell = sourceSheet.cells[formatAddress(row, column)] ?? { value: null }
            const after: CellState =
              operation.filterColumn !== undefined
                ? { value: cell.value }
                : cell.formula
                  ? { value: null, formula: offsetFormulaRefs(cell.formula, rowDelta, columnDelta) }
                  : { value: cell.value }
            const address = formatAddress(target.startRow + targetOffset, column + columnDelta)
            const before = targetBefore.cells[address] ?? { value: null }
            cellChanges.push({ sheetId: targetBefore.id, address, before, after })
            replaceCell(working, targetBefore.id, address, after)
          }
        }
        continue
      }
      if (operation.op === 'convert_to_values') {
        // The AI propose path pre-expands this into set_cell ops using the
        // live grid's computed values (this snapshot only stores what was
        // written, not what formulas evaluate to).
        throw new WorkbookConflictError(
          'convert_to_values reaches the in-memory adapter only above the ' +
            `${MAX_EXPANDED_CELL_OPS}-cell demo limit — convert smaller ranges, or work on an imported xlsx file.`,
        )
      }
      if (operation.op === 'find_replace') {
        // Only >MAX_EXPANDED_CELL_OPS ranges arrive range-level; smaller
        // ones were already expanded into plain set_cell edits.
        throw new WorkbookConflictError(
          `find_replace on an in-memory workbook is limited to ${MAX_EXPANDED_CELL_OPS} cells per operation (imported xlsx files allow up to 200,000).`,
        )
      }
      if (operation.op === 'sort_range') {
        throw new WorkbookConflictError(
          `sort_range on an in-memory workbook is limited to ${MAX_EXPANDED_CELL_OPS} cells per operation (imported xlsx files allow up to 200,000).`,
        )
      }
      if (operation.op === 'clear_range') {
        // Range-level clear (only ranges above the per-cell expansion cap
        // arrive in this form): existing cells become ordinary cell changes,
        // empty ones need no work.
        const bounds = parseRange(operation.range)
        const sheet = findSheet(working, operation.sheetId)
        for (const [address, cell] of Object.entries(sheet.cells)) {
          const parsed = parseAddress(address)
          if (parsed.row < bounds.startRow || parsed.row > bounds.endRow) continue
          if (parsed.column < bounds.startColumn || parsed.column > bounds.endColumn) continue
          if (cell.value === null && cell.formula === undefined) continue
          cellChanges.push({ sheetId: sheet.id, address, before: cell, after: { value: null } })
          replaceCell(working, sheet.id, address, { value: null })
        }
        continue
      }
      if (isLayoutOp(operation)) {
        applyLayoutOp(working, operation)
        structuralChanges.push({ op: operation, label: layoutOpLabel(operation) })
        continue
      }
      if (operation.op === 'format_range') {
        findSheet(working, operation.sheetId)
        applyFormatRange(working, operation)
        formatChanges.push({
          sheetId: operation.sheetId,
          range: operation.range,
          format: operation.format,
          ...(operation.reset ? { reset: true } : {}),
          label: formatOpLabel(operation),
        })
        continue
      }
      if (isStructuralOp(operation)) {
        const refErrors = applyStructuralOp(working, operation)
        structuralChanges.push({ op: operation, label: structuralOpLabel(operation) })
        if (refErrors.length > 0) {
          warnings.push(
            `Formulas reference deleted targets and will break (#REF!): ${refErrors.join(', ')}`,
          )
        }
        continue
      }
      if (operation.op === 'rename_sheet') {
        const sheet = findSheet(working, operation.sheetId)
        if (working.sheets.some((candidate) => candidate.name === operation.name)) {
          throw new WorkbookConflictError(`A sheet named "${operation.name}" already exists.`)
        }
        sheetRenames.push({ sheetId: sheet.id, before: sheet.name, after: operation.name })
        replaceSheet(working, { ...sheet, name: operation.name })
        continue
      }

      const sheet = findSheet(working, operation.sheetId)
      const before = sheet.cells[operation.address] ?? { value: null }
      if (operation.op === 'set_cell' && operation.expectedValue !== undefined) {
        if (before.value !== operation.expectedValue) {
          throw new WorkbookConflictError(
            `${sheet.name}!${operation.address} no longer has the expected value.`,
          )
        }
      }
      if (operation.op === 'set_formula' && operation.expectedFormula !== undefined) {
        if (before.formula !== operation.expectedFormula) {
          throw new WorkbookConflictError(
            `${sheet.name}!${operation.address} no longer has the expected formula.`,
          )
        }
      }

      const after =
        operation.op === 'set_cell'
          ? { value: operation.value }
          : operation.op === 'set_formula'
            ? { value: null, formula: operation.formula }
            : { value: null }
      cellChanges.push({ sheetId: sheet.id, address: operation.address, before, after })
      replaceCell(working, sheet.id, operation.address, after)
    }

    return {
      transactionId: batch.transactionId,
      baseRevision: batch.baseRevision,
      cellChanges,
      sheetRenames,
      structuralChanges,
      formatChanges,
      warnings,
    }
  }

  apply(plan: ChangePlan): CommitReceipt {
    if (plan.baseRevision !== this.snapshot.revision) {
      throw new WorkbookConflictError('The workbook changed after preview.')
    }
    if (this.committedTransactions.has(plan.transactionId)) {
      throw new WorkbookConflictError('This transaction was already committed.')
    }

    const previous = structuredClone(this.snapshot)
    const next = structuredClone(this.snapshot)
    for (const rename of plan.sheetRenames) {
      const sheet = findSheet(next, rename.sheetId)
      replaceSheet(next, { ...sheet, name: rename.after })
    }
    for (const change of plan.structuralChanges) {
      if (isLayoutOp(change.op)) applyLayoutOp(next, change.op)
      else if (isStructuralOp(change.op)) applyStructuralOp(next, change.op)
      // fill_range / clear_range never reach the demo adapter as range-level
      // entries — plan() expands them into ordinary cellChanges above.
    }
    for (const change of plan.cellChanges) {
      replaceCell(next, change.sheetId, change.address, change.after)
    }
    for (const change of plan.formatChanges) {
      applyFormatRange(next, {
        op: 'format_range',
        sheetId: change.sheetId,
        range: change.range,
        format: change.format,
        ...(change.reset ? { reset: true } : {}),
      })
    }

    this.history.push(previous)
    this.snapshot = { ...next, revision: next.revision + 1 }
    this.committedTransactions.add(plan.transactionId)
    const receipt = {
      transactionId: plan.transactionId,
      previousRevision: previous.revision,
      revision: this.snapshot.revision,
    }
    this.auditLog.push({
      ...receipt,
      action: 'apply',
      changedCellCount: plan.cellChanges.length,
      renamedSheetCount: plan.sheetRenames.length,
      structuralChangeCount: plan.structuralChanges.length,
    })
    return receipt
  }

  findVisual(visualId: string): SheetVisual | null {
    for (const sheet of this.snapshot.sheets) {
      const visual = sheet.visuals?.find((candidate) => candidate.id === visualId)
      if (visual) return structuredClone(visual)
    }
    return null
  }

  /// Direct visual mutation for interactive edits (ribbon, drag, delete).
  /// Deliberately no history entry and no revision bump: these ride the
  /// editor's own undo stack via closures holding the previous state, and
  /// a pending AI preview's CAS base must not be invalidated by them.
  /// `next: null` removes; a missing id with a non-null `next` re-inserts
  /// (the undo of a removal). Returns the replaced visual.
  replaceVisual(visualId: string, next: SheetVisual | null): SheetVisual | null {
    for (const sheet of this.snapshot.sheets) {
      const visuals = sheet.visuals ?? []
      const index = visuals.findIndex((candidate) => candidate.id === visualId)
      if (index < 0) continue
      const previous = structuredClone(visuals[index]) ?? null
      replaceSheet(this.snapshot, {
        ...sheet,
        visuals:
          next === null
            ? visuals.filter((candidate) => candidate.id !== visualId)
            : visuals.map((candidate) =>
                candidate.id === visualId ? structuredClone(next) : candidate,
              ),
      })
      return previous
    }
    if (next !== null) {
      const sheet = this.snapshot.sheets.find((candidate) => candidate.id === next.sheetId)
      if (sheet) {
        replaceSheet(this.snapshot, {
          ...sheet,
          visuals: [...(sheet.visuals ?? []), structuredClone(next)],
        })
      }
    }
    return null
  }

  /** Whether undo() has a transaction to revert (drives the QAT button gray state). */
  get canUndo(): boolean {
    return this.history.length > 0
  }

  undo(): CommitReceipt {
    const previous = this.history.pop()
    if (!previous) {
      throw new WorkbookConflictError('There is no transaction to undo.')
    }
    const currentRevision = this.snapshot.revision
    this.snapshot = { ...structuredClone(previous), revision: currentRevision + 1 }
    const receipt = {
      transactionId: `undo-${currentRevision}`,
      previousRevision: currentRevision,
      revision: this.snapshot.revision,
    }
    this.auditLog.push({
      ...receipt,
      action: 'undo',
      changedCellCount: 0,
      renamedSheetCount: 0,
      structuralChangeCount: 0,
    })
    return receipt
  }
}

/// Row/column shifts re-key cell addresses and rewrite A1 references in every
/// formula of the workbook (deleted targets become #REF!, ranges shrink).
/// Returns the cells whose formulas ended up with #REF!.
function applyStructuralOp(snapshot: WorkbookSnapshot, op: StructuralOperation): string[] {
  if (op.op === 'duplicate_sheet' || op.op === 'set_sheet_hidden' || op.op === 'move_sheet') {
    throw new WorkbookConflictError(
      `${op.op} needs an imported xlsx file — the demo workbook does not support it.`,
    )
  }
  if (op.op === 'add_sheet') {
    if (snapshot.sheets.some((candidate) => candidate.name === op.name)) {
      throw new WorkbookConflictError(`A sheet named "${op.name}" already exists.`)
    }
    const sheets = snapshot.sheets as WorksheetState[]
    let serial = sheets.length + 1
    while (sheets.some((sheet) => sheet.id === `sheet-${serial}`)) serial += 1
    sheets.push({
      id: `sheet-${serial}`,
      name: op.name,
      cells: {},
      ...(op.rows !== undefined ? { gridRows: op.rows } : {}),
      ...(op.columns !== undefined ? { gridColumns: op.columns } : {}),
    })
    return []
  }

  if (op.op === 'delete_sheet') {
    const sheet = findSheet(snapshot, op.sheetId)
    if (snapshot.sheets.length === 1) {
      throw new WorkbookConflictError('The workbook needs at least one sheet.')
    }
    const sheets = snapshot.sheets as WorksheetState[]
    sheets.splice(
      sheets.findIndex((candidate) => candidate.id === op.sheetId),
      1,
    )
    // Formulas referencing the deleted sheet keep their text (Excel would
    // show #REF!); surface them so the plan can warn.
    const danglers: string[] = []
    const namePattern = new RegExp(
      `(?:'${sheet.name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}'|\\b${sheet.name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')})!`,
    )
    for (const candidate of snapshot.sheets) {
      for (const [address, cell] of Object.entries(candidate.cells)) {
        if (cell.formula && namePattern.test(cell.formula)) {
          danglers.push(`${candidate.name}!${address}`)
        }
      }
    }
    return danglers
  }

  const sheet = findSheet(snapshot, op.sheetId)
  const spec = shiftSpecForOp(op)!
  const cells = shiftAddressedRecord(sheet.cells, op)
  const styles = sheet.styles ? shiftAddressedRecord(sheet.styles, op) : undefined
  const merges = sheet.merges
    ? sheet.merges
        .map((merge) => shiftRangeString(merge, spec))
        .filter((merge): merge is string => merge !== null)
    : undefined
  const rowHeights =
    sheet.rowHeights && spec.axis === 'row'
      ? shiftNumericKeyRecord(
          sheet.rowHeights,
          spec,
          (key) => Number(key) - 1,
          (index) => String(index + 1),
        )
      : sheet.rowHeights
  const colWidths =
    sheet.colWidths && spec.axis === 'column'
      ? shiftNumericKeyRecord(sheet.colWidths, spec, columnIndex, columnLabel)
      : sheet.colWidths
  replaceSheet(snapshot, { ...sheet, cells, styles, merges, rowHeights, colWidths })

  const refErrors: string[] = []
  for (const candidate of snapshot.sheets) {
    let nextCells: Record<string, CellState> | null = null
    for (const [address, cell] of Object.entries(candidate.cells)) {
      if (!cell.formula) continue
      const result = shiftFormulaRefs(cell.formula, op, candidate.id === op.sheetId, sheet.name)
      if (!result.changed) continue
      nextCells ??= { ...candidate.cells }
      nextCells[address] = { ...cell, formula: result.formula }
      if (result.hasRefError) refErrors.push(`${candidate.name}!${address}`)
    }
    if (nextCells) replaceSheet(snapshot, { ...candidate, cells: nextCells })
  }
  return refErrors
}

/** re-keys an address→value record for a row/column insert or delete */
function shiftAddressedRecord<T>(
  record: Readonly<Record<string, T>>,
  op: Extract<
    StructuralOperation,
    { op: 'insert_rows' | 'delete_rows' | 'insert_cols' | 'delete_cols' }
  >,
): Record<string, T> {
  const next: Record<string, T> = {}
  for (const [address, value] of Object.entries(record)) {
    const { row, column } = parseAddress(address)
    let nextRow = row
    let nextColumn = column
    if (op.op === 'insert_rows') {
      if (row >= op.row - 1) nextRow = row + op.count
    } else if (op.op === 'delete_rows') {
      if (row >= op.row - 1 && row < op.row - 1 + op.count) continue
      if (row >= op.row - 1 + op.count) nextRow = row - op.count
    } else if (op.op === 'insert_cols') {
      const pivot = columnIndex(op.column)
      if (column >= pivot) nextColumn = column + op.count
    } else {
      const pivot = columnIndex(op.column)
      if (column >= pivot && column < pivot + op.count) continue
      if (column >= pivot + op.count) nextColumn = column - op.count
    }
    next[formatAddress(nextRow, nextColumn)] = value
  }
  return next
}

/** shifts a range string ("B2:D4") for a row/column insert or delete;
 * null when the whole range fell inside a deleted region */
function shiftRangeString(range: string, spec: ShiftSpec): string | null {
  const bounds = parseRange(range)
  const startValue = spec.axis === 'row' ? bounds.startRow : bounds.startColumn
  const endValue = spec.axis === 'row' ? bounds.endRow : bounds.endColumn
  let start = shiftIndex(startValue, spec)
  let end = shiftIndex(endValue, spec)
  if (start === null && end === null) return null
  if (start === null) start = spec.start
  if (end === null) end = spec.start - 1
  if (end < start) return null
  const first =
    spec.axis === 'row'
      ? { row: start, column: bounds.startColumn }
      : { row: bounds.startRow, column: start }
  const second =
    spec.axis === 'row'
      ? { row: end, column: bounds.endColumn }
      : { row: bounds.endRow, column: end }
  if (first.row === second.row && first.column === second.column) return null
  return `${formatAddress(first.row, first.column)}:${formatAddress(second.row, second.column)}`
}

/** shifts records keyed by row number / column label (heights, widths) */
function shiftNumericKeyRecord(
  record: Readonly<Record<string, number>>,
  spec: ShiftSpec,
  toIndex: (key: string) => number,
  toKey: (index: number) => string,
): Record<string, number> {
  const next: Record<string, number> = {}
  for (const [key, value] of Object.entries(record)) {
    const shifted = shiftIndex(toIndex(key), spec)
    if (shifted !== null) next[toKey(shifted)] = value
  }
  return next
}

function rangesIntersect(a: string, b: string): boolean {
  const boundsA = parseRange(a)
  const boundsB = parseRange(b)
  return (
    boundsA.startRow <= boundsB.endRow &&
    boundsB.startRow <= boundsA.endRow &&
    boundsA.startColumn <= boundsB.endColumn &&
    boundsB.startColumn <= boundsA.endColumn
  )
}

function applyLayoutOp(snapshot: WorkbookSnapshot, op: LayoutOperation): void {
  if (op.op === 'subtotal')
    throw new Error('unsupported_operation: subtotal requires a canonical workbook session')
  if (op.op === 'edit_pivot')
    throw new Error('unsupported_operation: edit_pivot requires a canonical workbook session')
  if (op.op === 'edit_table')
    throw new Error('unsupported_operation: edit_table requires a canonical workbook session')
  if (op.op === 'relocate_range')
    throw new Error('unsupported_operation: range relocation requires a canonical session')
  if (op.op === 'paste_cells') {
    if (op.writeValues !== undefined || op.format !== undefined || op.skipCells !== undefined)
      throw new Error('unsupported_operation: selective paste requires a canonical session')
    const sheet = findSheet(snapshot, op.sheetId),
      origin = parseAddress(op.start),
      cells = { ...sheet.cells },
      styles = { ...sheet.styles }
    for (const [r, row] of op.cells.entries())
      for (const [c, cell] of row.entries()) {
        const address = formatAddress(origin.row + r, origin.column + c)
        cells[address] = {
          value: cell.formula ? null : cell.rawValue,
          ...(cell.formula ? { formula: cell.formula } : {}),
          ...(cell.kind === 'error' ? { error: true } : {}),
        }
        const {
          horizontalAlignment,
          verticalAlignment,
          textRotation,
          fontColor,
          fillColor,
          ...style
        } = cell.style ?? {}
        styles[address] = {
          ...style,
          ...(fontColor === undefined
            ? {}
            : {
                fontColor:
                  fontColor === null ? null : resolveStyleColor(normalizeStyleColor(fontColor)),
              }),
          ...(fillColor === undefined
            ? {}
            : {
                fillColor:
                  fillColor === null ? null : resolveStyleColor(normalizeStyleColor(fillColor)),
              }),
          ...(cell.numberFormat ? { numberFormat: cell.numberFormat } : {}),
          ...(horizontalAlignment ? { horizontalAlign: horizontalAlignment } : {}),
          ...(verticalAlignment ? { verticalAlign: verticalAlignment } : {}),
          ...(textRotation === undefined
            ? {}
            : {
                textRotation:
                  textRotation === 255
                    ? 'vertical'
                    : textRotation > 90
                      ? 90 - textRotation
                      : textRotation,
              }),
        }
      }
    const area = {
      startRow: origin.row,
      startColumn: origin.column,
      endRow: origin.row + op.cells.length - 1,
      endColumn: origin.column + op.cells[0]!.length - 1,
    }
    const intersects = (m: ReturnType<typeof parseRange>) =>
      m.startRow <= area.endRow &&
      m.endRow >= area.startRow &&
      m.startColumn <= area.endColumn &&
      m.endColumn >= area.startColumn
    for (const range of sheet.merges ?? []) {
      const m = parseRange(range)
      if (
        intersects(m) &&
        (op.merges === undefined ||
          m.startRow < area.startRow ||
          m.endRow > area.endRow ||
          m.startColumn < area.startColumn ||
          m.endColumn > area.endColumn)
      )
        throw new WorkbookConflictError('merge_partial_overlap')
    }
    replaceSheet(snapshot, {
      ...sheet,
      cells,
      styles,
      ...(op.merges === undefined
        ? {}
        : {
            merges: [
              ...(sheet.merges ?? []).filter((m) => !intersects(parseRange(m))),
              ...op.merges,
            ],
          }),
    })
    return
  }
  if (op.op === 'add_picture') {
    const sheet = findSheet(snapshot, op.sheetId)
    replaceSheet(snapshot, {
      ...sheet,
      visuals: [
        ...(sheet.visuals ?? []),
        {
          id: `added-image-${op.sheetId}-${(sheet.visuals?.length ?? 0) + 1}`,
          sheetId: sheet.id,
          kind: 'image',
          anchor: structuredClone(op.anchor),
          mediaType: op.mediaType,
          mediaDataUrl: `data:${op.mediaType};base64,${op.base64}`,
          name: op.name,
        },
      ],
    })
    return
  }
  if (op.op === 'add_shape') {
    const sheet = findSheet(snapshot, op.sheetId),
      cell = parseAddress(op.anchorCell)
    const anchor = op.anchor ?? {
      fromRow: cell.row,
      fromColumn: cell.column,
      fromRowOffset: 0,
      fromColumnOffset: 0,
      toRow: Math.min(1048575, cell.row + (op.shapeType === 'textbox' ? 4 : 6)),
      toColumn: Math.min(16383, cell.column + (op.shapeType === 'textbox' ? 4 : 3)),
      toRowOffset: 0,
      toColumnOffset: 0,
    }
    if (
      (anchor.toRow <= anchor.fromRow && anchor.toRowOffset <= anchor.fromRowOffset) ||
      (anchor.toColumn <= anchor.fromColumn && anchor.toColumnOffset <= anchor.fromColumnOffset)
    )
      throw new WorkbookConflictError('invalid_shape_anchor')
    replaceSheet(snapshot, {
      ...sheet,
      visuals: [
        ...(sheet.visuals ?? []),
        {
          id: (() => {
            let n = (sheet.visuals?.length ?? 0) + 1
            while (sheet.visuals?.some((v) => v.id === `added-shape-${op.sheetId}-${n}`)) n++
            return `added-shape-${op.sheetId}-${n}`
          })(),
          sheetId: sheet.id,
          kind: 'shape',
          shapeType: op.shapeType === 'textbox' ? 'rect' : op.shapeType,
          anchor: structuredClone(anchor),
          fillColor: op.fillColor ?? (op.shapeType === 'textbox' ? '#FFFFFF' : '#DDEBF7'),
          ...(op.text === undefined
            ? op.shapeType === 'textbox'
              ? { text: 'Text' }
              : {}
            : { text: op.text }),
          ...(op.shapeType === 'textbox' ? { isTextBox: true } : {}),
          editAs: 'twoCell',
        },
      ],
    })
    return
  }
  if (op.op === 'edit_shape') {
    const sheet = snapshot.sheets.find((s) =>
      s.visuals?.some((v) => v.id === op.visualId && v.kind === 'shape'),
    )
    const visual = sheet?.visuals?.find((v) => v.id === op.visualId)
    if (!sheet || !visual || visual.kind !== 'shape')
      throw new WorkbookConflictError('unknown_shape')
    const anchor = op.anchorCell
      ? (() => {
          const point = parseAddress(op.anchorCell)
          const rows = point.row - visual.anchor.fromRow,
            cols = point.column - visual.anchor.fromColumn
          return {
            ...visual.anchor,
            fromRow: point.row,
            toRow: visual.anchor.toRow + rows,
            fromColumn: point.column,
            toColumn: visual.anchor.toColumn + cols,
          }
        })()
      : visual.anchor
    if (anchor.toRow > 1048575 || anchor.toColumn > 16383)
      throw new WorkbookConflictError('invalid_shape_anchor')
    replaceSheet(snapshot, {
      ...sheet,
      visuals: sheet.visuals!.map((v) =>
        v.id === visual.id
          ? {
              ...v,
              anchor,
              ...(op.text === undefined ? {} : { text: op.text }),
              ...(op.fillColor === undefined ? {} : { fillColor: op.fillColor }),
            }
          : v,
      ),
    })
    return
  }
  if (op.op === 'edit_visual') {
    const sheet = snapshot.sheets.find((sheet) => sheet.visuals?.some((v) => v.id === op.visualId))
    if (!sheet?.visuals) throw new WorkbookConflictError(`Unknown visual: ${op.visualId}`)
    const visual = sheet.visuals.find((v) => v.id === op.visualId)!
    if (
      visual.kind === 'chart' &&
      ['fillColor', 'lineColor', 'rotation', 'flipH', 'flipV', 'frameSize'].some((key) =>
        Object.hasOwn(op, key),
      )
    )
      throw new WorkbookConflictError('unsupported_visual_properties')
    if (visual.kind === 'image' && (op.fillColor !== undefined || op.lineColor !== undefined))
      throw new WorkbookConflictError('unsupported_visual_properties')
    let visuals = sheet.visuals.map((v) =>
      v.id === op.visualId
        ? {
            ...v,
            ...(op.altText === undefined ? {} : { altText: op.altText }),
            ...(op.editAs === undefined ? {} : { editAs: op.editAs }),
            ...(v.kind !== 'chart'
              ? {
                  ...(op.fillColor === undefined ? {} : { fillColor: op.fillColor }),
                  ...(op.lineColor === undefined ? {} : { lineColor: op.lineColor }),
                  ...(op.rotation === undefined ? {} : { rotation: op.rotation }),
                  ...(op.flipH === undefined ? {} : { flipH: op.flipH }),
                  ...(op.flipV === undefined ? {} : { flipV: op.flipV }),
                  ...(op.frameSize
                    ? { frameWidth: op.frameSize.width, frameHeight: op.frameSize.height }
                    : {}),
                }
              : {}),
          }
        : v,
    )
    if (op.zIndex !== undefined) {
      if (op.zIndex >= visuals.length) throw new WorkbookConflictError('invalid_visual_z_index')
      const updated = visuals.find((v) => v.id === visual.id)!
      visuals = visuals.filter((v) => v.id !== visual.id)
      visuals.splice(op.zIndex, 0, updated)
    }
    replaceSheet(snapshot, { ...sheet, visuals })
    return
  }
  if (op.op === 'set_visual_anchor') {
    const sheet = findSheet(snapshot, op.sheetId)
    if (!sheet.visuals?.some((visual) => visual.id === op.visualId))
      throw new WorkbookConflictError(`Unknown visual: ${op.visualId}`)
    replaceSheet(snapshot, {
      ...sheet,
      visuals: sheet.visuals.map((visual) =>
        visual.id === op.visualId ? { ...visual, anchor: structuredClone(op.anchor) } : visual,
      ),
    })
    return
  }
  if (op.op === 'edit_chart') {
    // Demo charts have no chart part; the visual id doubles as chartPath.
    let sheet: WorksheetState | undefined
    let visual: import('./chart-visual').SheetChartVisual | undefined
    for (const candidate of snapshot.sheets) {
      const found = candidate.visuals?.find(
        (entry) => entry.id === op.chartPath && entry.kind === 'chart',
      )
      if (found) {
        sheet = candidate
        visual = found as import('./chart-visual').SheetChartVisual
        break
      }
    }
    if (!sheet || !visual) {
      throw new WorkbookConflictError(
        `Unknown chart: ${op.chartPath} (demo charts are addressed by their visual id).`,
      )
    }
    const chartVisual = visual
    if (
      op.axisTitles !== undefined &&
      chartVisual.chart.chartTypes.some((type) => /pie|doughnut/i.test(type))
    ) {
      throw new WorkbookConflictError(
        'Pie/doughnut charts have no axes — axisTitles does not apply.',
      )
    }
    if (op.grouping !== undefined) {
      const targetTypes =
        op.chartType !== undefined
          ? (CHART_EDIT_TYPES[op.chartType]?.chartTypes ?? [])
          : chartVisual.chart.chartTypes
      if (!targetTypes.some((type) => /^(barChart|lineChart|areaChart)$/.test(type))) {
        throw new WorkbookConflictError('Only bar, line, and area charts support grouping.')
      }
    }
    const series = chartSeriesRangeEdits(
      snapshot,
      chartVisual.sheetId,
      chartVisual.chart.series.length,
      op.seriesData ?? [],
    )
    const seriesSet = op.seriesSelection
      ? buildChartSeriesSelection(chartVisual.chart, op.seriesSelection, (reference) => {
          const resolved = resolveChartRange(reference, snapshot.sheets, chartVisual.sheetId)
          return {
            ref: resolved.ref,
            vector: gridFromRange(findSheet(snapshot, resolved.sheetId), resolved.range).flat(),
          }
        })
      : undefined
    const edit: ChartStateEdit = {
      ...(seriesSet ? { seriesSet } : {}),
      ...(op.title === undefined ? {} : { title: op.title }),
      ...(op.chartType === undefined ? {} : { chartType: op.chartType }),
      ...(op.seriesColors === undefined ? {} : { seriesColors: op.seriesColors }),
      ...(op.pointColors === undefined ? {} : { pointColors: op.pointColors }),
      ...(op.dataLabelPosition === undefined ? {} : { dataLabelPosition: op.dataLabelPosition }),
      ...(op.dataLabelFormat === undefined ? {} : { dataLabelFormat: op.dataLabelFormat }),
      ...(op.gridlines === undefined ? {} : { gridlines: op.gridlines }),
      ...(op.valueAxis === undefined ? {} : { valueAxis: op.valueAxis }),
      ...(op.gapWidthPct === undefined ? {} : { gapWidthPct: op.gapWidthPct }),
      ...(op.holeSizePct === undefined ? {} : { holeSizePct: op.holeSizePct }),
      ...(op.explosionPct === undefined ? {} : { explosionPct: op.explosionPct }),
      ...(op.pointExplosions === undefined ? {} : { pointExplosions: op.pointExplosions }),

      ...(op.legend === undefined ? {} : { legend: op.legend }),
      ...(op.dataLabels === undefined ? {} : { dataLabels: op.dataLabels }),
      ...(op.grouping === undefined ? {} : { grouping: op.grouping }),
      ...(op.axisTitles === undefined ? {} : { axisTitles: op.axisTitles }),
      ...(series.length > 0 ? { series } : {}),
    }
    assertChartFormat(chartVisual.chart, edit)
    const chart = applyChartStateEdit(chartVisual.chart, edit)
    replaceSheet(snapshot, {
      ...sheet,
      visuals: (sheet.visuals ?? []).map((entry) =>
        entry.id === chartVisual.id && entry.kind === 'chart' ? { ...entry, chart } : entry,
      ),
    })
    return
  }
  if (op.op === 'add_chart') {
    const sheet = findSheet(snapshot, op.sheetId)
    if (rangeCellCount(parseRange(op.dataRange)) > 2000) {
      throw new WorkbookConflictError('add_chart dataRange covers more than 2000 cells.')
    }
    try {
      const visual = buildChartVisual({
        id: `demo-chart-${op.sheetId}-${(sheet.visuals?.length ?? 0) + 1}`,
        sheetId: sheet.id,
        sheetName: sheet.name,
        chartType: op.chartType,
        dataRange: op.dataRange,
        title: op.title,
        anchorCell: op.anchorCell,
        values: gridFromRange(sheet, op.dataRange),
      })
      replaceSheet(snapshot, { ...sheet, visuals: [...(sheet.visuals ?? []), visual] })
    } catch (error: unknown) {
      throw new WorkbookConflictError(error instanceof Error ? error.message : String(error))
    }
    return
  }
  if (op.op === 'delete_visual') {
    const sheet = snapshot.sheets.find((candidate) =>
      candidate.visuals?.some((visual) => visual.id === op.visualId),
    )
    if (!sheet) {
      throw new WorkbookConflictError(
        `Unknown visual: ${op.visualId} (demo charts are addressed by their visual id).`,
      )
    }
    replaceSheet(snapshot, {
      ...sheet,
      visuals: (sheet.visuals ?? []).filter((visual) => visual.id !== op.visualId),
    })
    return
  }
  if (
    op.op !== 'merge_cells' &&
    op.op !== 'unmerge_cells' &&
    op.op !== 'set_row_height' &&
    op.op !== 'set_col_width'
  ) {
    // Shapes, images, links, filters, CF/DV, protection, and defined names
    // save into a real file; the demo snapshot has nowhere to keep them.
    throw new WorkbookConflictError(
      `${op.op} needs an imported xlsx file — the demo workbook does not support it.`,
    )
  }
  const sheet = findSheet(snapshot, op.sheetId)
  if (op.op === 'merge_cells') {
    if (rangeCellCount(parseRange(op.range)) < 2) {
      throw new WorkbookConflictError('merge_cells needs a range of at least two cells.')
    }
    const merges = sheet.merges ?? []
    if (merges.includes(op.range)) return
    const clash = merges.find((existing) => rangesIntersect(existing, op.range))
    if (clash) {
      throw new WorkbookConflictError(`${op.range} overlaps the existing merged range ${clash}.`)
    }
    replaceSheet(snapshot, { ...sheet, merges: [...merges, op.range] })
    return
  }
  if (op.op === 'unmerge_cells') {
    const merges = (sheet.merges ?? []).filter((existing) => !rangesIntersect(existing, op.range))
    replaceSheet(snapshot, { ...sheet, merges: merges.length > 0 ? merges : undefined })
    return
  }
  if (op.op === 'set_row_height') {
    const rowHeights: Record<string, number> = { ...sheet.rowHeights }
    for (let offset = 0; offset < op.count; offset += 1) {
      rowHeights[String(op.row + offset)] = op.heightPoints
    }
    replaceSheet(snapshot, { ...sheet, rowHeights })
    return
  }
  const colWidths: Record<string, number> = { ...sheet.colWidths }
  const first = columnIndex(op.column)
  for (let offset = 0; offset < op.count; offset += 1) {
    colWidths[columnLabel(first + offset)] = op.widthPx
  }
  replaceSheet(snapshot, { ...sheet, colWidths })
}

/// Merges a format patch into every cell of the range: null clears a
/// property, undefined leaves it alone; empty results drop the entry.
function applyFormatRange(snapshot: WorkbookSnapshot, op: FormatRangeOperation): void {
  const sheet = findSheet(snapshot, op.sheetId)
  const styles: Record<string, CellFormatState> = { ...sheet.styles }
  for (const address of rangeAddresses(parseRange(op.range))) {
    const merged = mergeFormat(op.reset ? undefined : styles[address], op.format)
    if (merged) styles[address] = merged
    else delete styles[address]
  }
  replaceSheet(snapshot, {
    ...sheet,
    styles: Object.keys(styles).length > 0 ? styles : undefined,
  })
}

function mergeFormat(
  existing: CellFormatState | undefined,
  patch: CellFormatPatch,
): CellFormatState | null {
  const merged: Record<string, unknown> = { ...existing }
  for (const [key, value] of Object.entries(patch)) {
    if (value === undefined) continue
    if (value === null) {
      if (
        [
          'fontColor',
          'fillColor',
          'fill',
          'vertAlign',
          'borderTop',
          'borderBottom',
          'borderLeft',
          'borderRight',
        ].includes(key)
      ) {
        merged[key] = null
        if (key === 'fill') merged.fillColor = null
        if (key === 'fillColor') delete merged.fill
      } else if (['bold', 'italic', 'underline', 'strikethrough'].includes(key)) merged[key] = false
      else delete merged[key]
    } else if (key === 'fontColor' || key === 'fillColor') {
      merged[key] = resolveStyleColor(normalizeStyleColor(value as StyleColorInput))
      if (key === 'fillColor') delete merged.fill
    } else if (key === 'fill') {
      merged.fill = value
      const display = fillDisplayColor(value as FillPatch as FillSpec)
      if (display) merged.fillColor = resolveStyleColor(normalizeStyleColor(display))
    } else if (key === 'border') {
      const border = value as BorderPatch
      // Keep the legacy preset echo while retaining independently composed edges.
      merged.border =
        border.color === undefined
          ? border
          : {
              ...border,
              color: resolveStyleColor(normalizeStyleColor(border.color)),
            }
      for (const edge of ['top', 'bottom', 'left', 'right'])
        if (border.type === 'all' || border.type === 'none' || border.type === edge)
          merged[`border${edge[0]!.toUpperCase()}${edge.slice(1)}`] =
            border.type === 'none'
              ? null
              : {
                  style: border.style ?? 'thin',
                  ...(border.color === undefined
                    ? {}
                    : { color: resolveStyleColor(normalizeStyleColor(border.color)) }),
                }
    } else merged[key] = value
  }
  return Object.keys(merged).length > 0 ? (merged as CellFormatState) : null
}

/// Cell values over a range as a row-major grid. Formula cells surface their
/// stored value, which the demo snapshot keeps as null (never computed) — a
/// chart over formula results is rejected by the numeric-column check.
function gridFromRange(sheet: WorksheetState, range: string): ChartGridValue[][] {
  const bounds = parseRange(range)
  const grid: ChartGridValue[][] = []
  for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
    const line: ChartGridValue[] = []
    for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
      line.push(sheet.cells[formatAddress(row, column)]?.value ?? null)
    }
    grid.push(line)
  }
  return grid
}

function findSheet(snapshot: WorkbookSnapshot, sheetId: string): WorksheetState {
  const sheet = snapshot.sheets.find((candidate) => candidate.id === sheetId)
  if (!sheet) {
    throw new WorkbookConflictError(`Unknown sheet: ${sheetId}`)
  }
  return sheet
}

function replaceSheet(snapshot: WorkbookSnapshot, replacement: WorksheetState): void {
  const sheets = snapshot.sheets as WorksheetState[]
  const index = sheets.findIndex((sheet) => sheet.id === replacement.id)
  sheets[index] = replacement
}

function replaceCell(
  snapshot: WorkbookSnapshot,
  sheetId: string,
  address: string,
  cell: CellState,
): void {
  const sheet = findSheet(snapshot, sheetId)
  // Every caller owns a structuredClone made by plan() or apply(). Copying the
  // complete dictionary per cell makes a 1,000-cell batch quadratic in workbook
  // size; mutating that private working copy preserves external snapshot isolation.
  ;(sheet.cells as Record<string, CellState>)[address] = cell
}
