import { isDateFormat } from 'numfmt'
import {
  formatAddress,
  parseAddress,
  parseRange,
  rangeAddresses,
} from '@genoffice/xlsx-gateway/domain/cell-address'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import { assertMergeOperation } from '@genoffice/xlsx-gateway/domain/session-merges'
import { stateChangeset } from '@genoffice/xlsx-gateway/shared/session-changeset'
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import {
  WorkbookError,
  assertNotCancelled,
  type TransactionInput,
  type WorkbookState,
} from '../domain'
import { prepareCandidateMaterial } from '../candidate-material'
import { refreshChartReferences } from '@genoffice/xlsx-gateway/domain/session-chart-references'
import type { WorkbookSnapshot, CellState } from '@genoffice/xlsx-gateway/domain/workbook.types'
import type { SheetVisual } from '@genoffice/xlsx-gateway/domain/chart-visual'
import type { WorkbookSource } from '../ports'
import { fromTypedCell } from './cells'
import { importBatchSchema, type ImportBatch } from './schemas'

/** A row-major batch is at most a leading row, full middle rows and trailing row. */
export function importValueOperations(batch: ImportBatch, sheetId: string): WorkbookOperation[] {
  const rows: { row: number; column: number; values: (string | number | boolean | null)[] }[] = []
  for (const item of batch.cells) {
    const point = parseAddress(item.address),
      previous = rows.at(-1),
      value = fromTypedCell(item.cell).rawValue
    if (
      previous &&
      point.row === previous.row &&
      point.column === previous.column + previous.values.length
    )
      previous.values.push(value)
    else {
      if (
        previous &&
        (point.row < previous.row ||
          (point.row === previous.row &&
            point.column <= previous.column + previous.values.length - 1))
      )
        throw new WorkbookError('artifact_incomplete', 'Import cells must be unique and ordered.')
      rows.push({ row: point.row, column: point.column, values: [value] })
    }
  }
  const groups: { row: number; column: number; values: (string | number | boolean | null)[][] }[] =
    []
  for (const row of rows) {
    const previous = groups.at(-1)
    if (
      previous &&
      row.row === previous.row + previous.values.length &&
      row.column === previous.column &&
      row.values.length === previous.values[0]!.length
    )
      previous.values.push(row.values)
    else groups.push({ row: row.row, column: row.column, values: [row.values] })
  }
  if (groups.length > 3)
    throw new WorkbookError(
      'type_mismatch',
      'A prepared import batch must be a contiguous slice of its output rectangle.',
    )
  return groups.map((group) => ({
    op: 'set_range',
    sheetId,
    start: formatAddress(group.row, group.column),
    values: group.values,
  }))
}

/** Both sub-stages are disposable. Only the combined root enters host commit. */
export async function prepareImportCandidate(input: {
  state: WorkbookState
  intent: TransactionInput
  batch: ImportBatch
  source: WorkbookSource
  sourceSessionId: string
  signal?: AbortSignal
}) {
  const batch = importBatchSchema.parse(input.batch)
  if (input.state.sheets.some((sheet) => sheet.readOnly))
    throw new WorkbookError(
      'readonly_workbook',
      'Save As and reopen a writable copy before importing changes.',
    )
  let current = input.state,
    sheetId: string
  const execute = async (operations: WorkbookOperation[]) => {
    const material = await prepareCandidateMaterial({
      state: current,
      operations,
      source: input.source,
      sourceSessionId: input.sourceSessionId,
      signal: input.signal,
    })
    const changes = await computeCandidateState(
      current,
      { ...input.intent, operations },
      material,
      { signal: input.signal },
    )
    current = changes.nextState
    return material
  }
  if ('newSheetName' in batch.target) {
    await execute([{ op: 'add_sheet', name: batch.target.newSheetName }])
    sheetId = current.sheets.find(
      (sheet) => !input.state.sheets.some((before) => before.id === sheet.id),
    )!.id
  } else sheetId = batch.target.sheetId
  // Exact typed cell primitives need one disposable root, not repeated legacy
  // grid snapshots. Protection and dependency reads still use the normal host path.
  const material = await prepareCandidateMaterial({
    state: current,
    operations: importValueOperations(batch, sheetId),
    source: input.source,
    sourceSessionId: input.sourceSessionId,
    signal: input.signal,
  })
  if (!material.writableSheetIds.includes(sheetId))
    throw new WorkbookError('protected_target', 'The prepared import target is not editable.')
  for (const operation of importValueOperations(batch, sheetId))
    assertMergeOperation(material.state, operation)
  const next = structuredClone(current),
    sheet = next.sheets.find((item) => item.id === sheetId)!
  next.stateId = material.allocation.stateId
  if (!sheet)
    throw new WorkbookError('target_invalid', 'The prepared import sheet no longer exists.')
  const hydrated = material.state.sheets.find((item) => item.id === sheetId)!
  if (hydrated.readOnly || hydrated.protected)
    throw new WorkbookError('protected_target', 'The import worksheet is not editable.')
  if (hydrated.features) sheet.features = structuredClone(hydrated.features)
  sheet.protected = hydrated.protected ?? false
  for (const { address, cell } of batch.cells) {
    assertNotCancelled(input.signal)
    if (cell.kind === 'formula' && batch.mode !== 'formulas')
      throw new WorkbookError('type_mismatch', 'Formula mode was not authorized.')
    const value = fromTypedCell(cell),
      date = cell.kind === 'formula' ? cell.cached : cell
    if (date?.kind === 'date' && date.dateSystem !== next.dateSystem)
      throw new WorkbookError('type_mismatch', 'Date systems must match.')
    const previous = sheet.cells[address] ?? hydrated.cells[address]
    const numberFormat =
      value.numberFormat ??
      (value.kind === 'number' && previous?.numberFormat && isDateFormat(previous.numberFormat)
        ? 'General'
        : previous?.numberFormat)
    const style = { ...previous?.style, ...value.style, ...(numberFormat ? { numberFormat } : {}) }
    sheet.cells[address] = {
      ...value,
      ...(Object.keys(style).length ? { style } : {}),
      ...(numberFormat ? { numberFormat } : {}),
    }
    const point = parseAddress(address)
    sheet.rows = Math.max(sheet.rows, point.row + 1)
    sheet.columns = Math.max(sheet.columns, point.column + 1)
  }
  if (
    next.sheets.reduce(
      (count, item) => count + Object.values(item.cells).filter((cell) => !!cell.formula).length,
      0,
    ) > 50_000
  )
    throw new WorkbookError(
      'formula_dependency_budget',
      'Imported formulas exceed the retained formula dependency budget.',
    )
  if (
    next.sheets.some(
      (owner) => Array.isArray(owner.features?.charts) && owner.features.charts.length,
    )
  ) {
    const dependencyBook: WorkbookSnapshot = {
      revision: input.intent.expectedRevision,
      sheets: next.sheets.map((owner) => {
        const cells: Record<string, CellState> = {},
          original = material.state.sheets.find((item) => item.id === owner.id)!
        for (const range of material.referenceRanges[owner.id] ?? [])
          for (const address of rangeAddresses(parseRange(range))) {
            const cell = Object.hasOwn(owner.cells, address)
              ? owner.cells[address]
              : original.cells[address]
            cells[address] = {
              value: cell?.rawValue ?? null,
              ...(cell?.formula ? { formula: cell.formula } : {}),
            }
          }
        return {
          id: owner.id,
          name: owner.name,
          cells,
          visuals: structuredClone((owner.features?.charts ?? []) as unknown as SheetVisual[]),
        }
      }),
    }
    refreshChartReferences(dependencyBook, material.completeSheetIds, material.referenceRanges)
    for (const owner of next.sheets)
      if (Array.isArray(owner.features?.charts))
        owner.features = {
          ...owner.features,
          charts: JSON.parse(
            JSON.stringify(dependencyBook.sheets.find((item) => item.id === owner.id)!.visuals),
          ),
        }
  }
  assertNotCancelled(input.signal)
  const changes = stateChangeset(
    input.state,
    next,
    batch.cells.map((cell) => `${sheetId}!${cell.address}`),
  )
  return {
    ...changes,
    calculationState:
      material.completeSheetIds.length < next.sheets.length
        ? ('pending' as const)
        : changes.calculationState,
  }
}
