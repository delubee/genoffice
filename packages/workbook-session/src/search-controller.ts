import { coerceReplaceValue, replaceSearchText } from '@genoffice/xlsx-gateway/domain/search-text'
import {
  workbookOperationSchema,
  type WorkbookOperation,
} from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { randomUUID } from 'node:crypto'
import {
  sessionSearchRequestSchema,
  sessionSearchPageSchema,
  type SessionFindQuery,
} from '@genoffice/xlsx-gateway/shared/session-search'
import type { WorkbookSource } from './ports'
import { WorkbookError, type WorkbookState } from './domain'
import { findCurrentCells } from './search'

/** One bounded search ticket per document view; no caller-supplied source handle. */
export class WorkbookSearchController {
  private sequence = 0
  private closed = false
  private active: { queryId: string; stateId: string } | undefined
  private cached:
    | { id: string; query: SessionFindQuery; result: Awaited<ReturnType<typeof findCurrentCells>> }
    | undefined
  constructor(
    private readonly input: {
      source: WorkbookSource
      sourceSessionId: string
      getState(): Readonly<WorkbookState>
      replace?(stateId: string, operations: WorkbookOperation[]): Promise<unknown>
    },
  ) {}
  async handle(raw: unknown) {
    const request = sessionSearchRequestSchema.parse(raw)
    if (this.closed) throw new WorkbookError('view_closed', 'The search view has closed.')
    if (request.action === 'release') {
      if (this.cached?.id !== request.searchId || this.cached.result.stateId !== request.stateId)
        throw new WorkbookError('search_not_found', 'The search session is no longer available.')
      this.cached = undefined
      this.active = undefined
      this.sequence++
      return { released: true }
    }
    if (request.action === 'cancel') {
      const cancelled =
        this.active?.queryId === request.queryId && this.active.stateId === request.stateId
      if (cancelled) {
        this.sequence++
        this.active = undefined
        this.cached = undefined
      }
      return { cancelled }
    }
    const state = this.input.getState()
    if (request.stateId !== state.stateId)
      throw new WorkbookError(
        'revision_conflict',
        'The requested search state is no longer current.',
        {},
        true,
      )
    if (request.action === 'start') {
      const sequence = ++this.sequence
      this.cached = undefined
      this.active = { queryId: request.queryId, stateId: request.stateId }
      const { action: _action, queryId: _queryId, ...query } = request
      const result = await findCurrentCells({
        state,
        query,
        source: this.input.source,
        sourceSessionId: this.input.sourceSessionId,
        isCurrent: () =>
          !this.closed &&
          sequence === this.sequence &&
          this.input.getState().stateId === state.stateId,
      })
      if (
        this.closed ||
        sequence !== this.sequence ||
        this.input.getState().stateId !== state.stateId
      )
        throw new WorkbookError(
          'revision_conflict',
          'The document changed before search publication.',
          {},
          true,
        )
      this.cached = { id: randomUUID(), query, result }
    } else if (
      this.cached?.id !== request.searchId ||
      this.cached.result.stateId !== state.stateId
    ) {
      throw new WorkbookError('search_not_found', 'The search session is no longer available.')
    }
    const cached = this.cached!
    if (request.action === 'replace') {
      if (!this.input.replace)
        throw new WorkbookError('unsupported_operation', 'Search replacement is unavailable.')
      const hits =
        request.mode === 'all' ? cached.result.hits : [cached.result.hits[request.index ?? -1]]
      if (hits.some((hit) => !hit))
        throw new WorkbookError('invalid_request', 'Choose a current search result to replace.')
      const metadata = state.features?.originalMetadata as { readOnly?: boolean } | undefined
      const selected = hits.filter((hit) => !(hit!.formula && cached.query.lookIn === 'values'))
      if (selected.length > 2000)
        throw new WorkbookError(
          'range_limit',
          'Replacement exceeds 2,000 cells; narrow the query. No cells were changed.',
        )
      const operations = selected.map((hit) => {
        const sheet = state.sheets.find((sheet) => sheet.id === hit!.sheetId)!
        if (metadata?.readOnly || state.sheets.some((sheet) => sheet.readOnly) || sheet.protected)
          throw new WorkbookError(
            'readonly_workbook',
            'The replacement includes a read-only or protected sheet.',
          )
        const text = replaceSearchText(hit!.text, cached.query, request.replacement)
        if (text.length > 32_767)
          throw new WorkbookError(
            'range_limit',
            'The replacement exceeds the cell text limit. No cells were changed.',
          )
        return workbookOperationSchema.parse(
          hit!.formula && text.startsWith('=')
            ? { op: 'set_formula', sheetId: hit!.sheetId, address: hit!.address, formula: text }
            : {
                op: 'set_cell',
                sheetId: hit!.sheetId,
                address: hit!.address,
                value: coerceReplaceValue(hit!.value, text),
              },
        )
      })
      // Consume the ticket before asynchronous submission. Retries cannot apply it twice.
      this.sequence++
      this.active = undefined
      this.cached = undefined
      if (operations.length) await this.input.replace(state.stateId, operations)
      return { replaced: operations.length, skipped: hits.length - operations.length }
    }
    const offset = request.action === 'page' ? request.offset : 0
    if (offset > cached.result.hits.length)
      throw new WorkbookError('invalid_request', 'The search page is outside the result set.')
    const hits = cached.result.hits.slice(offset, offset + 200).map((hit) => {
      const sheet = state.sheets.find((sheet) => sheet.id === hit.sheetId)!
      const metadata = state.features?.originalMetadata as { readOnly?: boolean } | undefined
      let text = hit.text.slice(0, 512)
      if (/[\uD800-\uDBFF]$/.test(text)) text = text.slice(0, -1)
      return {
        sheetId: hit.sheetId,
        address: hit.address,
        text,
        truncated: text.length < hit.text.length,
        isFormula: !!hit.formula,
        replaceable:
          !metadata?.readOnly &&
          !state.sheets.some((sheet) => sheet.readOnly) &&
          !sheet.protected &&
          !(hit.formula && cached.query.lookIn === 'values'),
      }
    })
    return sessionSearchPageSchema.parse({
      stateId: state.stateId,
      searchId: cached.id,
      offset,
      total: cached.result.hits.length,
      valuesPending: cached.result.valuesPending,
      hits,
      ...(offset + hits.length < cached.result.hits.length
        ? { nextOffset: offset + hits.length }
        : {}),
    })
  }
  dispose() {
    this.closed = true
    this.active = undefined
    this.sequence++
    this.cached = undefined
  }
}
