import { composeCalculationDefinitions } from './definitions'
import type { CalculationNameDefinition } from './names'
import { calculationTablesFromFeatures, type CalculationTable } from './tables'
import { calculationContextSchema, type CalculationContext } from './context'
import { z } from 'zod'
import { currentAxisIndex } from '@genoffice/xlsx-gateway/domain/session-coordinates'
import {
  formatAddress,
  parseAddress,
  parseRange,
} from '@genoffice/xlsx-gateway/domain/cell-address'
import { MAX_GRID_COLUMNS, MAX_GRID_ROWS } from '@genoffice/xlsx-gateway/shared/grid-bounds'
import {
  workbookFindCellsResultSchema,
  type WorkbookFindCellsRequest,
} from '@genoffice/xlsx-gateway/shared/workbook-protocol'
import {
  assertNotCancelled,
  LIMITS,
  WorkbookError,
  type CellRecord,
  type WorkbookState,
} from '../domain'
import type { WorkbookSource } from '../ports'
import { readCurrentRange } from '../read'
import { indexedFeatures } from '../source-features'

const readCoordinates = z.object({
  cells: z.array(
    z.object({ row: z.number().int().nonnegative(), column: z.number().int().nonnegative() }),
  ),
})

/** Host-private logical data, not yet an engine snapshot or a published calculation cache.
 * External resources and array ownership still need resource binding; table
 * layouts beyond the installed formula engine's exact subset fail closed.
 * Source feature geometry remains in SOURCE coordinates; logical overlays are separate.
 */
export interface CalculationSourceData {
  stateId: string
  sourceContext: CalculationContext
  definedNames: CalculationNameDefinition[]
  tables: CalculationTable[]
  dateSystem: WorkbookState['dateSystem']
  workbookFeatures: WorkbookState['features']
  structuralLog: WorkbookState['structuralLog']
  sheets: {
    id: string
    name: string
    originalSheetId?: string
    rows: number
    columns: number
    cells: Record<string, CellRecord>
    sourceFeatures: Record<string, unknown>
    logicalFeatures: WorkbookState['sheets'][number]['features']
    sourceArrays: { sourceAddress: string; sourceRange: string; logicalAddress: string | null }[]
  }[]
}

/** Bounded rectangles over sparse points; never expand a distant pair into a huge range. */
function readBands(addresses: Iterable<string>, limit: number) {
  const points = [...addresses]
    .map(parseAddress)
    .sort((a, b) => a.row - b.row || a.column - b.column)
  const bands: { startRow: number; endRow: number; startColumn: number; endColumn: number }[] = []
  for (const point of points) {
    const last = bands.at(-1)
    const firstColumn = Math.min(last?.startColumn ?? point.column, point.column)
    const lastColumn = Math.max(last?.endColumn ?? point.column, point.column)
    if (
      last &&
      point.row - last.endRow <= 32 &&
      (point.row - last.startRow + 1) * (lastColumn - firstColumn + 1) <= limit
    ) {
      last.endRow = point.row
      last.startColumn = firstColumn
      last.endColumn = lastColumn
    } else
      bands.push({
        startRow: point.row,
        endRow: point.row,
        startColumn: point.column,
        endColumn: point.column,
      })
  }
  return bands.map(
    (band) =>
      `${formatAddress(band.startRow, band.startColumn)}:${formatAddress(band.endRow, band.endColumn)}`,
  )
}

export async function loadCalculationSource(input: {
  state: Readonly<WorkbookState>
  source: WorkbookSource
  sourceSessionId: string
  /** Required at publication too: the owning session must still have this exact root. */
  isCurrent(): boolean
  signal?: AbortSignal
  maxCells?: number
  maxBytes?: number
  readBatchCells?: number
}): Promise<CalculationSourceData> {
  const check = () => {
    assertNotCancelled(input.signal)
    if (!input.isCurrent())
      throw new WorkbookError(
        'revision_conflict',
        'The document changed while loading calculation data.',
        {},
        true,
      )
  }
  check()
  const state = structuredClone(input.state)
  const maxCells = input.maxCells ?? LIMITS.artifactCells
  const maxBytes = input.maxBytes ?? LIMITS.artifactBytes
  const batch = input.readBatchCells ?? LIMITS.readCells
  for (const value of [maxCells, maxBytes, batch])
    if (!Number.isSafeInteger(value) || value <= 0)
      throw new WorkbookError(
        'invalid_range',
        'Calculation budgets must be positive safe integers.',
      )
  if (batch > LIMITS.rangeCells)
    throw new WorkbookError('range_limit', 'Calculation read batch exceeds the source range limit.')
  let bytes = 0,
    count = 0
  const charge = (value: unknown) => {
    bytes += Buffer.byteLength(JSON.stringify(value) ?? '')
    if (bytes > maxBytes)
      throw new WorkbookError('range_limit', 'Calculation source exceeds its byte budget.')
  }
  if (!input.source.readCalculationContext)
    throw new WorkbookError(
      'calculation_unavailable',
      'Private calculation resources are unavailable.',
    )
  check()
  const sourceContext = calculationContextSchema.parse(
    await input.source.readCalculationContext(input.sourceSessionId),
  )
  check()
  if (sourceContext.sessionId !== input.sourceSessionId)
    throw new WorkbookError(
      'session_mismatch',
      'Calculation context belongs to another source session.',
    )
  const result: CalculationSourceData = {
    sourceContext,
    definedNames: composeCalculationDefinitions(state, sourceContext),
    tables: [],
    stateId: state.stateId,
    dateSystem: state.dateSystem,
    workbookFeatures: state.features,
    structuralLog: state.structuralLog,
    sheets: [],
  }
  charge({ ...result, sheets: state.sheets.map(({ cells: _cells, ...sheet }) => sheet) })
  for (const sheet of state.sheets) {
    check()
    const logicalFeatures = structuredClone(sheet.features)
    const addresses = new Set<string>()
    const requiredSource = new Map<string, string>()
    const observedSource = new Set<string>()
    const arrays: CalculationSourceData['sheets'][number]['sourceArrays'] = []
    let sourceFeatures: Record<string, unknown> = {}
    const keep = (address: string) => {
      const point = parseAddress(address)
      if (point.row >= MAX_GRID_ROWS || point.column >= MAX_GRID_COLUMNS)
        throw new WorkbookError('invalid_range', 'Calculation source exceeds the Excel grid.')
      if (!addresses.has(address)) {
        if (++count > maxCells)
          throw new WorkbookError('range_limit', 'Calculation source exceeds its cell budget.')
        addresses.add(address)
        charge(address)
      }
    }
    if (sheet.originalSheetId) {
      if (!input.source.findBaseCells)
        throw new WorkbookError(
          'calculation_unavailable',
          'Complete source cell enumeration is unavailable.',
        )
      sourceFeatures = structuredClone(
        await indexedFeatures(
          new Proxy(input.source, {
            get(target, key) {
              if (key === 'readFeatures')
                return async (sessionId: string, sheetId: string) => {
                  check()
                  const features = await target.readFeatures(sessionId, sheetId)
                  check()
                  return features
                }
              const value = Reflect.get(target, key, target)
              return typeof value === 'function' ? value.bind(target) : value
            },
          }),
          input.sourceSessionId,
          sheet.originalSheetId,
          input.signal,
        ),
      )
      check()
      charge(sourceFeatures)
      const tables = calculationTablesFromFeatures(sourceFeatures, sheet.id)
      charge(tables)
      result.tables.push(...tables)
      let cursor: WorkbookFindCellsRequest['resumeAt']
      let previous = -1
      let maxRow = 0,
        maxColumn = 0
      const deadline = Date.now() + 30_000
      for (;;) {
        check()
        const page = workbookFindCellsResultSchema.parse(
          await input.source.findBaseCells({
            sessionId: input.sourceSessionId,
            sheetId: sheet.originalSheetId,
            query: '*',
            lookIn: 'stored_cells',
            matchCase: false,
            matchEntireCell: false,
            wildcards: false,
            limit: 10_000,
            ...(cursor ? { resumeAt: cursor } : {}),
          }),
        )
        check()
        if (!page.indexingComplete) {
          if (Date.now() >= deadline)
            throw new WorkbookError(
              'index_pending',
              'Calculation needs complete source indexing.',
              {},
              true,
            )
          await new Promise((resolve) => setTimeout(resolve, 25))
          continue
        }
        for (const hit of page.matches) {
          const position = hit.row * MAX_GRID_COLUMNS + hit.column
          if (
            hit.sheetId !== sheet.originalSheetId ||
            hit.row >= MAX_GRID_ROWS ||
            hit.column >= MAX_GRID_COLUMNS ||
            position <= previous ||
            (cursor && position < cursor.row * MAX_GRID_COLUMNS + cursor.column)
          )
            throw new WorkbookError(
              'incomplete_source',
              'Source enumeration returned foreign, duplicate or unordered coordinates.',
            )
          previous = position
          maxRow = Math.max(maxRow, hit.row + 1)
          maxColumn = Math.max(maxColumn, hit.column + 1)
          const row = currentAxisIndex(state.structuralLog, sheet.id, hit.row, 'row')
          const column = currentAxisIndex(state.structuralLog, sheet.id, hit.column, 'column')
          if (row !== null && column !== null) {
            const address = formatAddress(row, column)
            keep(address)
            if (!Object.hasOwn(sheet.cells, address))
              requiredSource.set(address, formatAddress(hit.row, hit.column))
          }
          if (hit.arrayRef) {
            const bounds = parseRange(hit.arrayRef)
            if (
              bounds.endRow >= MAX_GRID_ROWS ||
              bounds.endColumn >= MAX_GRID_COLUMNS ||
              bounds.startRow !== hit.row ||
              bounds.startColumn !== hit.column
            )
              throw new WorkbookError('incomplete_source', 'Source array ownership is invalid.')
            const array = {
              sourceAddress: formatAddress(hit.row, hit.column),
              sourceRange: hit.arrayRef,
              logicalAddress: row === null || column === null ? null : formatAddress(row, column),
            }
            charge(array)
            arrays.push(array)
          }
        }
        if (page.complete) {
          if (page.nextCursor)
            throw new WorkbookError(
              'incomplete_source',
              'Completed enumeration unexpectedly returned a cursor.',
            )
          break
        }
        const next = page.nextCursor
        if (
          !next ||
          next.sheetId !== sheet.originalSheetId ||
          next.row >= MAX_GRID_ROWS ||
          next.column >= MAX_GRID_COLUMNS ||
          next.row * MAX_GRID_COLUMNS + next.column <=
            Math.max(previous, cursor ? cursor.row * MAX_GRID_COLUMNS + cursor.column : -1)
        )
          throw new WorkbookError(
            'incomplete_source',
            'Source enumeration did not advance to a valid cursor.',
          )
        cursor = next
      }
      // Some producers understate <dimension>. Repair only this private read root.
      const metadata = sheet.features?.originalMetadata as Record<string, unknown> | undefined
      sheet.features = {
        ...sheet.features,
        originalMetadata: {
          ...metadata,
          rowCount: Math.max(Number(metadata?.rowCount ?? 0), maxRow),
          columnCount: Math.max(Number(metadata?.columnCount ?? 0), maxColumn),
        },
      }
    }
    for (const address of Object.keys(sheet.cells)) keep(address)
    const cells: Record<string, CellRecord> = {}
    // Proxy keeps class receivers intact and reuses the fully indexed style table.
    const source = new Proxy(input.source, {
      get(target, key) {
        if (key === 'readFeatures')
          return async () => {
            check()
            return sourceFeatures
          }
        if (key === 'readBaseRange')
          return async (sessionId: string, sheetId: string, range: string) => {
            check()
            const response = await target.readBaseRange(sessionId, sheetId, range)
            check()
            for (const cell of readCoordinates.parse(response).cells)
              observedSource.add(formatAddress(cell.row, cell.column))
            return response
          }
        const value = Reflect.get(target, key, target)
        return typeof value === 'function' ? value.bind(target) : value
      },
    })
    let rows = sheet.rows,
      columns = sheet.columns
    for (const range of readBands(addresses, batch)) {
      check()
      const page = await readCurrentRange({
        state,
        source,
        sessionId: input.sourceSessionId,
        logicalSheetId: sheet.id,
        range,
        limit: batch,
        signal: input.signal,
      })
      check()
      for (const { address, displayValue: _display, ...cell } of page.cells) {
        if (!addresses.has(address)) continue
        cells[address] = cell
        charge(cell)
        const point = parseAddress(address)
        rows = Math.max(rows, point.row + 1)
        columns = Math.max(columns, point.column + 1)
      }
    }
    for (const sourceAddress of requiredSource.values())
      if (!observedSource.has(sourceAddress))
        throw new WorkbookError(
          'incomplete_source',
          'A stored source cell was absent from its range read.',
        )
    result.sheets.push({
      id: sheet.id,
      name: sheet.name,
      ...(sheet.originalSheetId ? { originalSheetId: sheet.originalSheetId } : {}),
      rows,
      columns,
      cells,
      sourceFeatures,
      logicalFeatures,
      sourceArrays: arrays,
    })
  }
  check()
  return result
}
