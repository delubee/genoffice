/**
 * Workbook operation appliers for the sheets renderer.
 *
 * AI DSL operations (pivot / table / table-column adds) and pivot dialog
 * helpers applied against the live Univer runtime. Extracted from App.tsx;
 * every function receives its runtime and state explicitly.
 */
import { isDefaultFormat, numfmt } from '@univerjs/core'
import { DATE_1904_OFFSET, isCalendarDatePattern } from './numfmt-fix'
import type {
  AddPivotOperation,
  AddTableColumnOperation,
  AddTableOperation,
  AddTableRowOperation,
  DeleteTableColumnOperation,
  DeleteTableRowOperation,
} from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { columnLabel, parseAddress, parseRange } from '@genoffice/xlsx-gateway/domain/cell-address'
import { renameRefSheet } from '@genoffice/xlsx-gateway/domain/chart-visual'
import {
  areasOverlap,
  buildPivotLayout,
  PIVOT_SOURCE_COL_LIMIT,
  PIVOT_SOURCE_ROW_LIMIT,
  PivotLayoutError,
  pivotOutputArea,
  type PivotArea,
  type PivotLayout,
} from '@genoffice/xlsx-gateway/domain/pivot-layout'
import type {
  WorkbookPivotAdd,
  WorkbookTableAdd,
  WorkbookVisualObject,
} from '../shared/desktop-api'
import {
  recordPivotAdd,
  recordPivotCacheRefresh,
  recordPivotRefreshUpdate,
  recordTableAdd,
  updateTableAdd,
} from './edit-journal'
import { planFileTableRegistrations } from './file-tables'
import { areasIntersect, fileTableArea, liveFileTables, optionsOfTableAdd } from './table-design'
import { DEFAULT_TABLE_STYLE, builtinTablePalette, tableThemeJson } from './table-styles'
import { t } from './i18n/locale'
import type { PivotDefinition } from '@genoffice/xlsx-gateway/gateway/xlsx-pivot'
import { applyFormatPatchToRange, nextSessionPivotName, nextSessionTableName } from './univer-sync'
import type { LazyWorkbookState, UniverRuntime, UniverWorksheet } from './univer-state'

export function applyAiTableAdd(
  runtime: UniverRuntime,
  state: LazyWorkbookState,
  op: AddTableOperation,
): void {
  const worksheet = runtime.univerAPI.getActiveWorkbook()?.getSheetBySheetId(op.sheetId)
  if (!worksheet) throw new Error(`Unknown sheet: ${op.sheetId}`)
  const bounds = parseRange(op.range)
  if (bounds.endRow <= bounds.startRow) {
    throw new Error(t('appTableNeedsRows'))
  }
  const width = bounds.endColumn - bounds.startColumn + 1
  if (width > 1_000) throw new Error(t('appTableTooWide'))
  const fileTableNames = planFileTableRegistrations(state.file.sheets).map(
    (table) => table.tableName,
  )
  const name = op.name ?? nextSessionTableName(state.editJournal, fileTableNames)
  if (
    state.editJournal.tableAdds.some((table) => table.name.toLowerCase() === name.toLowerCase())
  ) {
    throw new Error(t('appTableNameUsed', { name }))
  }
  if (fileTableNames.some((table) => table.toLowerCase() === name.toLowerCase())) {
    throw new Error(t('appTableNameInFile', { name }))
  }
  for (const table of state.editJournal.tableAdds) {
    if (table.sheetId === op.sheetId && areasIntersect(bounds, table.area)) {
      throw new Error(t('appTableOverlapsSession', { name: table.name }))
    }
  }
  for (const { table, originalName } of liveFileTables(state, op.sheetId)) {
    if (areasIntersect(bounds, fileTableArea(table))) {
      throw new Error(t('appTableOverlapsFile', { name: table.name ?? originalName }))
    }
  }
  const headerValues =
    worksheet.getRange(bounds.startRow, bounds.startColumn, 1, width).getValues()[0] ?? []
  const columnNames: string[] = []
  const used = new Set<string>()
  for (let index = 0; index < width; index += 1) {
    const raw = String(headerValues[index] ?? '')
      .trim()
      .slice(0, 255)
    const base = raw.length === 0 ? `Column${index + 1}` : raw
    let candidate = base
    for (let suffix = 2; used.has(candidate.toLowerCase()); suffix += 1) {
      candidate = `${base}${suffix}`
    }
    used.add(candidate.toLowerCase())
    columnNames.push(candidate)
    if (candidate !== raw) {
      worksheet.getRange(bounds.startRow, bounds.startColumn + index).setValue(candidate)
    }
  }
  const entry: WorkbookTableAdd = {
    sheetId: op.sheetId,
    area: {
      startRow: bounds.startRow,
      startColumn: bounds.startColumn,
      endRow: bounds.endRow,
      endColumn: bounds.endColumn,
    },
    name,
    columnNames,
    ...(op.style === undefined ? {} : { style: op.style }),
    bandedRows: op.bandedRows ?? true,
  }
  // Rendering is best-effort (the facade call is async); the journal entry
  // is what the save writes, and the gateway re-checks conflicts.
  const tableId = `ai-table-${state.editJournal.tableAdds.length + 1}-${Date.now().toString(36)}`
  void Promise.resolve(worksheet.addTable(name, entry.area, tableId))
    .then(() => applySessionTableTheme(worksheet, tableId, entry, state.file.themeColors))
    .catch(() => undefined)
  recordTableAdd(state.editJournal, entry)
}

/// Paints the Excel built-in style over Univer's default table theme.
export function applySessionTableTheme(
  worksheet: UniverWorksheet,
  tableId: string,
  add: WorkbookTableAdd,
  themeColors: readonly string[] | undefined,
): void {
  const theme = tableThemeJson(
    `${tableId}-${Date.now().toString(36)}`,
    builtinTablePalette(add.style ?? DEFAULT_TABLE_STYLE, themeColors),
    optionsOfTableAdd(add),
  )
  const api = worksheet as unknown as {
    addTableTheme(id: string, theme: unknown): Promise<boolean>
  }
  void Promise.resolve(api.addTableTheme(tableId, theme)).catch(() => undefined)
}

/// AI add_table_row: inserts one or more rows into a session-added table.
/// The row index is 1-based within the data region (header = row 0).
export function applyAiTableRowAdd(
  runtime: UniverRuntime,
  state: LazyWorkbookState,
  op: AddTableRowOperation,
): void {
  const tableEntry = state.editJournal.tableAdds.find(
    (t) => t.sheetId === op.sheetId && t.name.toLowerCase() === op.tableName.toLowerCase(),
  )
  if (!tableEntry) throw new Error(`Table "${op.tableName}" not found in session journal.`)
  const worksheet = runtime.univerAPI.getActiveWorkbook()?.getSheetBySheetId(op.sheetId)
  if (!worksheet) throw new Error(`Unknown sheet: ${op.sheetId}`)
  const count = op.count ?? 1
  const dataRows = tableEntry.area.endRow - tableEntry.area.startRow
  // row is 1-based within data region; default = append at end
  const insertDataRow = op.row ?? dataRows + 1
  // Absolute sheet row: header is tableEntry.area.startRow (0-based),
  // data starts at startRow + 1.  Insert before this row in the sheet.
  const sheetRow = tableEntry.area.startRow + insertDataRow
  worksheet.insertRowsBefore(sheetRow, count)
  // Expand the journal entry's area
  updateTableAdd(state.editJournal, op.sheetId, op.tableName, {
    area: { ...tableEntry.area, endRow: tableEntry.area.endRow + count },
  })
}

/// AI delete_table_row: removes one or more rows from a session-added table.
export function applyAiTableRowDelete(
  runtime: UniverRuntime,
  state: LazyWorkbookState,
  op: DeleteTableRowOperation,
): void {
  const tableEntry = state.editJournal.tableAdds.find(
    (t) => t.sheetId === op.sheetId && t.name.toLowerCase() === op.tableName.toLowerCase(),
  )
  if (!tableEntry) throw new Error(`Table "${op.tableName}" not found in session journal.`)
  const worksheet = runtime.univerAPI.getActiveWorkbook()?.getSheetBySheetId(op.sheetId)
  if (!worksheet) throw new Error(`Unknown sheet: ${op.sheetId}`)
  const count = op.count ?? 1
  const sheetRow = tableEntry.area.startRow + op.row // 1-based data row → absolute
  worksheet.deleteRows(sheetRow, count)
  updateTableAdd(state.editJournal, op.sheetId, op.tableName, {
    area: { ...tableEntry.area, endRow: tableEntry.area.endRow - count },
  })
}

/// AI delete_table_column: removes one or more columns from a session-added table.
export function applyAiTableColumnDelete(
  runtime: UniverRuntime,
  state: LazyWorkbookState,
  op: DeleteTableColumnOperation,
): void {
  const tableEntry = state.editJournal.tableAdds.find(
    (t) => t.sheetId === op.sheetId && t.name.toLowerCase() === op.tableName.toLowerCase(),
  )
  if (!tableEntry) throw new Error(`Table "${op.tableName}" not found in session journal.`)
  const worksheet = runtime.univerAPI.getActiveWorkbook()?.getSheetBySheetId(op.sheetId)
  if (!worksheet) throw new Error(`Unknown sheet: ${op.sheetId}`)
  const count = op.count ?? 1
  const sheetCol = tableEntry.area.startColumn + op.column - 1
  worksheet.deleteColumns(sheetCol, count)
  const newCols = [...tableEntry.columnNames]
  newCols.splice(op.column - 1, count)
  updateTableAdd(state.editJournal, op.sheetId, op.tableName, {
    area: { ...tableEntry.area, endColumn: tableEntry.area.endColumn - count },
    columnNames: newCols,
  })
}

export function applyAiTableColumnAdd(
  runtime: UniverRuntime,
  state: LazyWorkbookState,
  op: AddTableColumnOperation,
): void {
  const tableEntry = state.editJournal.tableAdds.find(
    (t) => t.sheetId === op.sheetId && t.name.toLowerCase() === op.tableName.toLowerCase(),
  )
  if (!tableEntry) throw new Error(`Table "${op.tableName}" not found in session journal.`)
  const worksheet = runtime.univerAPI.getActiveWorkbook()?.getSheetBySheetId(op.sheetId)
  if (!worksheet) throw new Error(`Unknown sheet: ${op.sheetId}`)
  const count = op.count ?? 1
  const tableCols = tableEntry.area.endColumn - tableEntry.area.startColumn + 1
  const insertTableCol = op.column ?? tableCols + 1
  const sheetCol = tableEntry.area.startColumn + insertTableCol - 1
  worksheet.insertColumnsBefore(sheetCol, count)
  // Update column names list (insert placeholder(s) before insertTableCol-1)
  const newCols = [...tableEntry.columnNames]
  for (let i = 0; i < count; i += 1) {
    const colName = i === 0 ? op.columnName : `${op.columnName}${i + 1}`
    newCols.splice(insertTableCol - 1 + i, 0, colName)
  }
  // Write the header cell(s) so the worksheet reflects the new name(s)
  for (let i = 0; i < count; i += 1) {
    worksheet
      .getRange(tableEntry.area.startRow, sheetCol + i)
      .setValue(newCols[insertTableCol - 1 + i] ?? '')
  }
  updateTableAdd(state.editJournal, op.sheetId, op.tableName, {
    area: { ...tableEntry.area, endColumn: tableEntry.area.endColumn + count },
    columnNames: newCols,
  })
}

/// Width-independent pivot source read. getValues() goes through the view
/// interceptors, which clip numbers to the column width (#### fill,
/// budget-limited General) — lies that must not become pivot labels,
/// aggregates, saved sharedItems, or recompute keys (a date column displayed
/// as #### broke timeline binding and slicer recompute). Re-format raw
/// values through each cell's own pattern instead; General numbers stay
/// numeric.
export function readPivotSourceGrid(
  range: {
    getRawValues(): unknown[][]
    getNumberFormats(): string[][]
    getFormulas(): string[][]
  },
  date1904: boolean,
  valueFields: readonly string[] = [],
): (string | number | boolean | null)[][] {
  const rawValues = range.getRawValues() as (string | number | boolean | null)[][]
  const valueColumns = new Set(
    (rawValues[0] ?? []).flatMap((header, index) =>
      valueFields.includes(String(header ?? '').trim()) ? [index] : [],
    ),
  )
  const patterns = range.getNumberFormats()
  // 1904 workbooks: file-loaded static serials count from 1904-01-01, but
  // formula results come from the 1900-based engine — same rule as the
  // display interceptor's calendar-date shift.
  const formulas = date1904 ? range.getFormulas() : null
  return rawValues.map((row, rowOffset) =>
    row.map((value, columnOffset) => {
      if (typeof value !== 'number') return value
      // Measures must remain numeric; currency display strings aggregate as blanks.
      if (rowOffset > 0 && valueColumns.has(columnOffset)) return value
      const pattern = patterns[rowOffset]?.[columnOffset] ?? ''
      if (pattern === '' || isDefaultFormat(pattern)) return value
      const shift =
        formulas !== null &&
        isCalendarDatePattern(pattern) &&
        (formulas[rowOffset]?.[columnOffset] ?? '') === ''
          ? DATE_1904_OFFSET
          : 0
      try {
        return numfmt.format(pattern, value + shift, { nbsp: true, throws: false })
      } catch {
        return value
      }
    }),
  )
}

/// Rows the report-filter block occupies above the table body: one per page
/// field plus Excel's blank separator row.
export function pivotPageRows(pageFieldCount: number): number {
  return pageFieldCount > 0 ? pageFieldCount + 1 : 0
}

/// Relayout of an existing pivot: a file pivot (parts rewritten on save) or a
/// pivot created this session (its journal entry is replaced).
export interface PivotRelayoutTarget {
  readonly target:
    { readonly pivotPath: string; readonly cachePath: string } | { readonly sessionName: string }
  /// Current table body (location ref) and the report-filter rows above it.
  readonly oldBounds: {
    startRow: number
    startColumn: number
    endRow: number
    endColumn: number
  }
  readonly oldPageRows: number
}

const ALL_ITEMS_CAPTION = '(All)'

/// Returns the body (location) area of the written pivot.
export function applyAiPivotAdd(
  runtime: UniverRuntime,
  state: LazyWorkbookState,
  op: AddPivotOperation,
  relayout?: PivotRelayoutTarget,
): { readonly location: PivotArea; readonly pageRows: number } {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const sourceSheet = workbook?.getSheetBySheetId(op.sheetId)
  const targetSheetId = op.targetSheetId ?? op.sheetId
  const targetSheet = workbook?.getSheetBySheetId(targetSheetId)
  if (!sourceSheet) throw new Error(`Unknown sheet: ${op.sheetId}`)
  if (!targetSheet) throw new Error(`Unknown sheet: ${targetSheetId}`)
  if (
    state.editJournal.sheets.added.has(op.sheetId) ||
    state.editJournal.sheets.added.has(targetSheetId)
  ) {
    throw new Error(t('appPivotOnAddedSheet'))
  }
  const source = parseRange(op.sourceRange)
  const sourceRows = source.endRow - source.startRow + 1
  const sourceColumns = source.endColumn - source.startColumn + 1
  if (sourceRows < 2) throw new Error(t('appPivotSourceNeedsRows'))
  if (sourceRows > PIVOT_SOURCE_ROW_LIMIT) throw new Error(t('appPivotSourceRowLimit'))
  if (sourceColumns > PIVOT_SOURCE_COL_LIMIT) throw new Error(t('appPivotSourceColLimit'))

  const grid = readPivotSourceGrid(
    sourceSheet.getRange(source.startRow, source.startColumn, sourceRows, sourceColumns),
    state.file.date1904 === true,
    op.values.filter((value) => value.formula === undefined).map((value) => value.field),
  )
  let layout: PivotLayout
  try {
    layout = buildPivotLayout(grid, op, {
      subtotal: t('appPivotSubtotal'),
      grandTotal: 'Grand Total',
    })
  } catch (err) {
    if (err instanceof PivotLayoutError) {
      throw new Error(t(LAYOUT_ERROR_KEYS[err.code], err.params), { cause: err })
    }
    throw err
  }
  const { width, height } = layout
  // Report filters render above the body like Excel: "Field | (All)" rows and
  // a blank separator; the location ref stays the body only.
  const pageRows = pivotPageRows(layout.definition.pageFieldIndices?.length ?? 0)
  const pageMatrix: (string | number | null)[][] = (layout.definition.pageFieldIndices ?? []).map(
    (fieldIndex, page) => {
      const item = layout.definition.pageItems?.[page] ?? null
      const label =
        item === null ? ALL_ITEMS_CAPTION : (layout.definition.pageLevelItems?.[page]?.[item] ?? '')
      const line: (string | number | null)[] = new Array(width).fill(null)
      line[0] = layout.definition.fieldNames[fieldIndex] ?? ''
      line[1] = label
      return line
    },
  )
  if (pageRows > 0) pageMatrix.push(new Array(width).fill(null))
  const matrix = [...pageMatrix, ...layout.matrix]
  const blockHeight = pageRows + height

  const blockAnchor = parseAddress(op.targetCell)
  const anchor = { row: blockAnchor.row + pageRows, column: blockAnchor.column }
  const location = pivotOutputArea(anchor, layout)
  const block = { ...location, startRow: blockAnchor.row }
  if (targetSheetId === op.sheetId && areasOverlap(block, source)) {
    throw new Error(t('appPivotOverlapSource'))
  }
  const targetMeta = state.file.sheets.find((sheet) => sheet.id === targetSheetId)
  const sessionTarget =
    relayout !== undefined && 'sessionName' in relayout.target ? relayout.target.sessionName : null
  const fileTarget =
    relayout !== undefined && 'pivotPath' in relayout.target ? relayout.target : null
  for (const existing of targetMeta?.pivotRanges ?? []) {
    // When editing itself, overlapping the old output area is expected.
    if (
      relayout &&
      existing.startRow === relayout.oldBounds.startRow &&
      existing.startColumn === relayout.oldBounds.startColumn &&
      existing.endRow === relayout.oldBounds.endRow &&
      existing.endColumn === relayout.oldBounds.endColumn
    ) {
      continue
    }
    if (areasOverlap(block, existing)) throw new Error(t('appPivotOverlapExisting'))
  }
  for (const pivot of state.editJournal.pivotAdds) {
    if (pivot.sheetId !== targetSheetId || pivot.name === sessionTarget) continue
    if (areasOverlap(block, pivot.location)) {
      throw new Error(t('appPivotOverlapSession', { name: pivot.name }))
    }
  }

  // When editing a file pivot, name is just a placeholder — the saver keeps
  // the original name from the file.
  const name =
    sessionTarget ??
    (fileTarget ? 'PivotEdit' : (op.name ?? nextSessionPivotName(state.editJournal)))
  if (
    !relayout &&
    state.editJournal.pivotAdds.some((pivot) => pivot.name.toLowerCase() === name.toLowerCase())
  ) {
    throw new Error(t('appPivotNameUsed', { name }))
  }

  if (relayout) {
    // The old block = report-filter rows + body; cleared and rewritten as one
    // union so shrunken parts blank out.
    const old = {
      ...relayout.oldBounds,
      startRow: relayout.oldBounds.startRow - relayout.oldPageRows,
    }
    // Grown parts must be empty (same criteria as refresh growth); shrunken
    // parts are cleared via the union write.
    const assertEmpty = (row: number, column: number, rows: number, columns: number): void => {
      if (rows <= 0 || columns <= 0) return
      const values = targetSheet.getRange(row, column, rows, columns).getValues() as (
        string | number | boolean | null | undefined
      )[][]
      if (
        values.some((line) =>
          line.some((value) => value !== null && value !== undefined && value !== ''),
        )
      ) {
        throw new Error(t('appPivotRelayoutOverlap'))
      }
    }
    assertEmpty(old.endRow + 1, old.startColumn, block.endRow - old.endRow, width)
    assertEmpty(
      old.startRow,
      old.endColumn + 1,
      Math.min(old.endRow, block.endRow) - old.startRow + 1,
      block.endColumn - old.endColumn,
    )
    const totalRows = Math.max(blockHeight, old.endRow - old.startRow + 1)
    const totalColumns = Math.max(width, old.endColumn - old.startColumn + 1)
    const padded = Array.from({ length: totalRows }, (_, rowIndex) => {
      const row = matrix[rowIndex] ?? []
      return Array.from({ length: totalColumns }, (_, colIndex) => row[colIndex] ?? null)
    })
    targetSheet
      .getRange(blockAnchor.row, blockAnchor.column, totalRows, totalColumns)
      .setValues(
        padded as unknown as Parameters<ReturnType<UniverWorksheet['getRange']>['setValues']>[0],
      )
  } else {
    targetSheet
      .getRange(blockAnchor.row, blockAnchor.column, blockHeight, width)
      .setValues(
        matrix as unknown as Parameters<ReturnType<UniverWorksheet['getRange']>['setValues']>[0],
      )
  }

  // value columns' numFmt covers the data rows only (no header, no grand total)
  const dataHeight = height - 2
  for (const { columnOffset, format } of layout.numberFormats) {
    if (dataHeight <= 0) break
    const colIdx = anchor.column + columnOffset
    const startR = anchor.row + 1
    const rangeStr = `${columnLabel(colIdx)}${startR + 1}:${columnLabel(colIdx)}${startR + dataHeight}`
    applyFormatPatchToRange(targetSheet.getRange(rangeStr), { numberFormat: format })
  }

  const additionPayload: WorkbookPivotAdd = {
    sheetId: targetSheetId,
    sourceSheetId: op.sheetId,
    sourceArea: {
      startRow: source.startRow,
      startColumn: source.startColumn,
      endRow: source.endRow,
      endColumn: source.endColumn,
    },
    location,
    name,
    ...layout.definition,
  }
  if (sessionTarget !== null) {
    const at = state.editJournal.pivotAdds.findIndex((pivot) => pivot.name === sessionTarget)
    if (at < 0) throw new Error(t('appSlicerPivotStale'))
    state.editJournal.pivotAdds.splice(at, 1, additionPayload)
    return { location, pageRows }
  }
  if (fileTarget !== null && relayout) {
    const newOutputRef =
      `${columnLabel(location.startColumn)}${location.startRow + 1}` +
      `:${columnLabel(location.endColumn)}${location.endRow + 1}`
    recordPivotCacheRefresh(state.editJournal, fileTarget.cachePath)
    recordPivotRefreshUpdate(
      state.editJournal,
      fileTarget.cachePath,
      targetSheetId,
      newOutputRef,
      additionPayload,
    )
    // The old in-memory definition is stale: after deletion, refresh/slicers
    // are disabled for it until save-and-reopen; saving reopens the session and
    // re-parses the newly written definition. The pane keeps editing from the
    // journaled relayout payload.
    state.pivotDefinitions.delete(fileTarget.pivotPath)
    const staleRange = targetMeta?.pivotRanges.find(
      (range) =>
        range.startRow === relayout.oldBounds.startRow &&
        range.startColumn === relayout.oldBounds.startColumn &&
        range.endRow === relayout.oldBounds.endRow &&
        range.endColumn === relayout.oldBounds.endColumn,
    )
    if (staleRange) {
      staleRange.endRow = location.endRow
      staleRange.endColumn = location.endColumn
    }
    const pivotMeta = targetMeta?.pivotTables.find((entry) => entry.path === fileTarget.pivotPath)
    if (pivotMeta) pivotMeta.outputRef = newOutputRef
    return { location, pageRows }
  }
  recordPivotAdd(state.editJournal, additionPayload)
  return { location, pageRows }
}

export { pivotConfigToOpParts } from './pivot-op-parts'
import {
  PivotGrowError,
  grownPivotBlock,
  pivotColLineLabel as sharedPivotColLineLabel,
  pivotMemberCaption as sharedPivotMemberCaption,
} from '@genoffice/xlsx-gateway/domain/pivot-placement'
import { LAYOUT_ERROR_KEYS } from './pivot-op-parts'

export function renameChartRefsForSheet(
  state: LazyWorkbookState,
  oldName: string,
  newName: string,
): void {
  const renameSeries = <
    T extends {
      valuesRef?: string | undefined
      categoriesRef?: string | undefined
    },
  >(
    entry: T,
  ): T => ({
    ...entry,
    ...(entry.valuesRef === undefined
      ? {}
      : { valuesRef: renameRefSheet(entry.valuesRef, oldName, newName) }),
    ...(entry.categoriesRef === undefined
      ? {}
      : { categoriesRef: renameRefSheet(entry.categoriesRef, oldName, newName) }),
  })
  const renameVisual = (visual: WorkbookVisualObject): WorkbookVisualObject =>
    visual.chart
      ? { ...visual, chart: { ...visual.chart, series: visual.chart.series.map(renameSeries) } }
      : visual
  state.file.visuals.forEach((visual, at) => {
    state.file.visuals[at] = renameVisual(visual)
  })
  state.editJournal.visualAdds.forEach((visual, at) => {
    state.editJournal.visualAdds[at] = renameVisual(visual)
  })
  for (const [chartPath, edit] of state.editJournal.chartEdits) {
    if (edit.series === undefined && edit.seriesSet === undefined) continue
    state.editJournal.chartEdits.set(chartPath, {
      ...edit,
      ...(edit.series === undefined ? {} : { series: edit.series.map(renameSeries) }),
      ...(edit.seriesSet === undefined ? {} : { seriesSet: edit.seriesSet.map(renameSeries) }),
    })
  }
}

export function pivotMemberCaption(
  definition: PivotDefinition,
  field: number,
  member: number,
): string {
  return sharedPivotMemberCaption(definition, field, member)
}

export function pivotColLineLabel(
  definition: PivotDefinition,
  line: PivotDefinition['colLines'][number],
): string {
  return sharedPivotColLineLabel(definition, line, t('appPivotSubtotal'))
}

export function applyGrownPivotOutput(
  target: UniverWorksheet,
  grown: PivotDefinition,
  data: readonly (readonly (number | null)[])[],
  oldBounds: { startRow: number; startColumn: number; endRow: number; endColumn: number },
): string {
  let output: ReturnType<typeof grownPivotBlock>
  try {
    output = grownPivotBlock(grown, data, oldBounds, t('appPivotSubtotal'))
  } catch (error) {
    if (error instanceof PivotGrowError)
      throw new Error(t('appPivotGrowUnsupported'), { cause: error })
    throw error
  }
  // Newly occupied areas (the downward/rightward growth) must be empty to
  // avoid overwriting user content.
  for (const area of output.growth) {
    const values = target
      .getRange(
        area.startRow,
        area.startColumn,
        area.endRow - area.startRow + 1,
        area.endColumn - area.startColumn + 1,
      )
      .getValues() as (string | number | boolean | null | undefined)[][]
    if (
      values.some((line) =>
        line.some((value) => value !== null && value !== undefined && value !== ''),
      )
    ) {
      throw new Error(t('appPivotGrowConflict'))
    }
  }
  target
    .getRange(
      oldBounds.startRow,
      oldBounds.startColumn,
      output.block.length,
      output.block[0]!.length,
    )
    .setValues(
      output.block as unknown as Parameters<
        ReturnType<UniverWorksheet['getRange']>['setValues']
      >[0],
    )
  return output.outputRef
}
