import { parseAddress } from '@genoffice/xlsx-gateway/domain/cell-address'
import { formulaReadsCell } from '@genoffice/xlsx-gateway/domain/formula-audit'
import {
  workbookAuditRequestSchema,
  workbookAuditResultSchema,
  FORMULA_ERROR_VALUE,
} from '@genoffice/xlsx-gateway/shared/session-audit'
import { findCurrentCells } from './search'
import { logicalSheet } from './sheet-map'
import { WorkbookError, type WorkbookState } from './domain'
import type { WorkbookSource } from './ports'

/** Independent read-only query: never consumes a Find/Replace ticket or mutates history. */
export async function auditWorkbook(input: {
  request: unknown
  getState(): Readonly<WorkbookState>
  source: WorkbookSource
  sourceSessionId: string
}) {
  const request = workbookAuditRequestSchema.parse(input.request)
  const state = input.getState()
  const active = logicalSheet(state, request.sheetId)
  if (state.stateId !== request.stateId)
    throw new WorkbookError('revision_conflict', 'The audit state is no longer current.')
  const target = parseAddress(request.address)
  const result = await findCurrentCells({
    state,
    source: input.source,
    sourceSessionId: input.sourceSessionId,
    isCurrent: () => input.getState().stateId === state.stateId,
    query: {
      stateId: state.stateId,
      query: request.kind === 'errors' ? '#' : '=',
      lookIn: request.kind === 'errors' ? 'values' : 'formulas',
    },
  })
  const names = new Map(state.sheets.map((sheet) => [sheet.id, sheet.name]))
  const hits = result.hits
    .filter((hit) =>
      request.kind === 'errors'
        ? FORMULA_ERROR_VALUE.test(hit.text)
        : !!hit.formula &&
          !(hit.sheetId === active.id && hit.address === request.address) &&
          formulaReadsCell(
            hit.formula,
            names.get(hit.sheetId) === active.name,
            active.name,
            target,
          ),
    )
    .map((hit) => ({
      sheetId: hit.sheetId,
      address: hit.address,
      ...(request.kind === 'errors' ? { value: hit.text } : {}),
    }))
  return workbookAuditResultSchema.parse({
    stateId: state.stateId,
    hits,
    valuesPending: result.valuesPending,
  })
}
