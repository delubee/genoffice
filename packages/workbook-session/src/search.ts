import {
  rewriteNameReferences,
  type NameReferenceChange,
} from '@genoffice/xlsx-gateway/domain/defined-name-references'
import { scalarToText, createSearchTextMatcher } from '@genoffice/xlsx-gateway/domain/search-text'
import { formatAddress, parseAddress } from '@genoffice/xlsx-gateway/domain/cell-address'
import { currentAxisIndex } from '@genoffice/xlsx-gateway/domain/session-coordinates'
import { shiftedSessionFormula } from '@genoffice/xlsx-gateway/domain/session-formula'
import { MAX_GRID_COLUMNS, MAX_GRID_ROWS } from '@genoffice/xlsx-gateway/shared/grid-bounds'
import { sessionFindQuerySchema } from '@genoffice/xlsx-gateway/shared/session-search'
import {
  workbookFindCellsResultSchema,
  type WorkbookFindCellsRequest,
} from '@genoffice/xlsx-gateway/shared/workbook-protocol'
import type { WorkbookSource } from './ports'
import { LIMITS, WorkbookError, type WorkbookState } from './domain'
import { logicalSheet } from './sheet-map'

export interface SessionFindHit {
  sheetId: string
  address: string
  value: string | number | boolean | null
  text: string
  formula?: string
}

/** Full immutable-source search with current overlays and coordinate transforms. */
export async function findCurrentCells(input: {
  state: Readonly<WorkbookState>
  query: unknown
  source: WorkbookSource
  sourceSessionId: string
  isCurrent?(): boolean
}) {
  const query = sessionFindQuerySchema.parse(input.query)
  const assertCurrent = () => {
    if (query.stateId !== input.state.stateId || input.isCurrent?.() === false)
      throw new WorkbookError(
        'revision_conflict',
        'The document changed during the search.',
        {},
        true,
      )
  }
  assertCurrent()
  const requested = new Set(query.sheetIds ?? input.state.sheets.map((sheet) => sheet.id))
  for (const id of requested) logicalSheet(input.state, id)
  const sheets = input.state.sheets.filter((sheet) => requested.has(sheet.id))
  const matches = createSearchTextMatcher(query)
  const scalar = scalarToText
  const found = new Map<string, SessionFindHit>()
  let bytes = 0
  const keep = (hit: SessionFindHit) => {
    const key = `${hit.sheetId}:${hit.address}`
    if (found.has(key)) return
    bytes += Buffer.byteLength(JSON.stringify(hit))
    if (found.size >= 100_000 || bytes > LIMITS.artifactBytes)
      throw new WorkbookError('range_limit', 'Too many search results; narrow the query.')
    found.set(key, hit)
  }
  let valuesPending =
    query.lookIn === 'values' &&
    (input.state.structuralLog.length > 0 ||
      !!input.state.features?.definedNamesEdits ||
      input.state.sheets.some((sheet) => Object.keys(sheet.cells).length > 0))
  for (const sheet of sheets) {
    assertCurrent()
    if (sheet.originalSheetId) {
      if (!input.source.findBaseCells)
        throw new WorkbookError('search_unavailable', 'Indexed source search is unavailable.')
      // Structural edits can change formula text so a native literal query is
      // insufficient. Enumerate formulas separately, rewrite then match them.
      const changes = (input.state.features?.definedNameReferences ??
        []) as unknown as NameReferenceChange[]
      const transformed =
        query.lookIn === 'formulas' && (input.state.structuralLog.length > 0 || changes.length > 0)
      const passes = transformed ? (['nonformulas', 'formulas'] as const) : (['direct'] as const)
      for (const pass of passes) {
        let cursor: WorkbookFindCellsRequest['resumeAt']
        const seenCursors = new Set<string>()
        const waitingSince = Date.now()
        for (;;) {
          assertCurrent()
          const page = workbookFindCellsResultSchema.parse(
            await input.source.findBaseCells({
              sessionId: input.sourceSessionId,
              sheetId: sheet.originalSheetId,
              query: pass === 'formulas' ? '?' : query.query,
              matchCase: query.matchCase,
              matchEntireCell: pass === 'formulas' ? false : query.matchEntireCell,
              lookIn:
                pass === 'formulas'
                  ? 'formulas_only'
                  : pass === 'nonformulas'
                    ? 'values'
                    : query.lookIn,
              wildcards: pass === 'formulas',
              limit: 10_000,
              ...(cursor ? { resumeAt: cursor } : {}),
            }),
          )
          assertCurrent()
          if (!page.indexingComplete) {
            if (Date.now() - waitingSince > 30_000)
              throw new WorkbookError(
                'index_pending',
                'Search is waiting for complete source indexing.',
                {},
                true,
              )
            await new Promise((resolve) => setTimeout(resolve, 50))
            continue
          }
          for (const hit of page.matches) {
            if (
              hit.sheetId !== sheet.originalSheetId ||
              hit.row >= MAX_GRID_ROWS ||
              hit.column >= MAX_GRID_COLUMNS
            )
              throw new WorkbookError(
                'session_mismatch',
                'Search returned an invalid source coordinate.',
              )
            if (pass === 'nonformulas' && hit.formulaText) continue
            const row = currentAxisIndex(input.state.structuralLog, sheet.id, hit.row, 'row')
            const column = currentAxisIndex(
              input.state.structuralLog,
              sheet.id,
              hit.column,
              'column',
            )
            if (row === null || column === null) continue
            if (row >= MAX_GRID_ROWS || column >= MAX_GRID_COLUMNS) continue
            const address = formatAddress(row, column)
            if (Object.hasOwn(sheet.cells, address)) continue
            const formula = hit.formulaText
              ? rewriteNameReferences(
                  shiftedSessionFormula(input.state, sheet.id, hit.formulaText),
                  sheet.id,
                  {
                    changes,
                    sheets: input.state.sheets.map((item) => ({ id: item.id, name: item.name })),
                  },
                )
              : undefined
            const text =
              query.lookIn === 'formulas' && formula
                ? formula
                : (hit.valueText ?? scalar(hit.value))
            if (matches(text))
              keep({
                sheetId: sheet.id,
                address,
                value: hit.value,
                text: text!,
                ...(formula ? { formula } : {}),
              })
          }
          if (page.complete) break
          if (!page.nextCursor || page.nextCursor.sheetId !== sheet.originalSheetId)
            throw new WorkbookError(
              'incomplete_search',
              'The source search did not supply a complete continuation.',
            )
          const key = JSON.stringify(page.nextCursor)
          if (seenCursors.has(key))
            throw new WorkbookError(
              'incomplete_search',
              'The source search cursor did not advance.',
            )
          seenCursors.add(key)
          cursor = page.nextCursor
        }
      }
    }
    for (const [address, cell] of Object.entries(sheet.cells)) {
      if (query.lookIn === 'values' && cell.formula && cell.calculationState !== 'ready')
        valuesPending = true
      const value =
        cell.formula && cell.calculationState === 'ready'
          ? (cell.cachedValue ?? cell.rawValue)
          : cell.rawValue
      const text = query.lookIn === 'formulas' && cell.formula ? cell.formula : scalar(value)
      if (matches(text))
        keep({
          sheetId: sheet.id,
          address,
          value,
          text: text!,
          ...(cell.formula ? { formula: cell.formula } : {}),
        })
    }
  }
  assertCurrent()
  const order = new Map(input.state.sheets.map((sheet, index) => [sheet.id, index]))
  const hits = [...found.values()].sort((a, b) => {
    const sheet = order.get(a.sheetId)! - order.get(b.sheetId)!
    if (sheet) return sheet
    const left = parseAddress(a.address),
      right = parseAddress(b.address)
    return query.direction === 'row'
      ? left.row - right.row || left.column - right.column
      : left.column - right.column || left.row - right.row
  })
  return { stateId: input.state.stateId, hits, valuesPending }
}
