import { BehaviorSubject, Subject } from 'rxjs'
import {
  createInitFindReplaceState,
  FindBy,
  FindDirection,
  FindScope,
  type IFindMatch,
  type IFindReplaceState,
} from '@univerjs/find-replace'
import type { FindReplacePanelService } from '../../../apps/sheets/src/renderer/find-replace-session'
import { parseAddress, formatAddress } from '@genoffice/xlsx-gateway/domain/cell-address'
import {
  sessionSearchPageSchema,
  type SessionSearchRequest,
  type SessionSearchPage,
} from '@genoffice/xlsx-gateway/shared/session-search'
import { z } from 'zod'

type Context = { stateId: string; sheetId: string; unitId: string; address?: string }
const replacementResult = z.strictObject({
  replaced: z.number().int().nonnegative(),
  skipped: z.number().int().nonnegative(),
})

/** Desktop presentation over complete host queries; never searches the loaded viewport. */
export class HostFindReplaceService implements FindReplacePanelService {
  readonly state$ = new BehaviorSubject<IFindReplaceState>({
    ...createInitFindReplaceState(),
    revealed: false,
    replaceRevealed: false,
  })
  readonly focusSignal$ = new Subject<void>()
  readonly busy$ = new BehaviorSubject(false)
  readonly currentMatch$ = new BehaviorSubject<IFindMatch | null>(null)
  readonly replaceables$ = new BehaviorSubject<(IFindMatch & { replaceable: boolean })[]>([])
  private matches: IFindMatch[] = []
  private ticket: { stateId: string; searchId: string } | undefined
  private active: { stateId: string; queryId: string } | undefined
  private epoch = 0
  private timer: ReturnType<typeof setTimeout> | undefined
  private position = -1
  private context: Context | null = null
  private closed = false
  private replacing = false
  private composing = false
  constructor(
    private readonly input: {
      search(request: SessionSearchRequest): Promise<unknown>
      getContext(): Context | null
      flush(): Promise<boolean>
      jump(sheetId: string, address: string, isCurrent: () => boolean): Promise<void>
      report(error: unknown): void
    },
  ) {}
  get revealed() {
    return this.state$.value.revealed
  }
  get replaceRevealed() {
    return this.state$.value.replaceRevealed
  }
  getFindString() {
    return this.state$.value.findString
  }
  getMatches() {
    return this.matches
  }
  private patch(update: Partial<IFindReplaceState>) {
    this.state$.next({ ...this.state$.value, ...update })
  }
  private release() {
    if (this.timer) clearTimeout(this.timer)
    this.timer = undefined
    const active = this.active,
      ticket = this.ticket
    this.active = undefined
    this.ticket = undefined
    if (active) void this.input.search({ action: 'cancel', ...active }).catch(() => {})
    if (ticket) void this.input.search({ action: 'release', ...ticket }).catch(() => {})
    this.epoch++
    this.busy$.next(false)
    this.matches = []
    this.position = -1
    this.currentMatch$.next(null)
    this.replaceables$.next([])
    this.patch({ matchesCount: 0, matchesPosition: 0, findCompleted: false })
  }
  open(replace: boolean) {
    if (this.closed) return
    this.patch({
      revealed: true,
      replaceRevealed: replace,
      findString: this.state$.value.inputtingFindString,
    })
    this.find()
    this.focusSignal$.next()
  }
  terminate() {
    this.release()
    this.patch({ revealed: false, replaceRevealed: false })
  }
  revealReplace() {
    this.patch({ replaceRevealed: true })
    this.find()
  }
  hideReplace() {
    this.patch({ replaceRevealed: false })
  }
  changeInputtingFindString(value: string) {
    this.release()
    this.patch({ inputtingFindString: value })
    if (this.replaceRevealed && !this.composing)
      this.timer = setTimeout(() => this.changeFindString(value), 300)
  }
  setComposing(value: boolean) {
    this.composing = value
    if (value) this.release()
  }
  changeFindString(value: string) {
    this.patch({ findString: value, inputtingFindString: value })
    this.find()
  }
  changeReplaceString(value: string) {
    this.patch({ replaceString: value })
  }
  changeCaseSensitive(value: boolean) {
    this.patch({ caseSensitive: value })
    this.find()
  }
  changeMatchesTheWholeCell(value: boolean) {
    this.patch({ matchesTheWholeCell: value })
    this.find()
  }
  changeFindScope(value: FindScope) {
    this.patch({ findScope: value })
    this.find()
  }
  changeFindDirection(value: FindDirection) {
    this.patch({ findDirection: value })
    this.find()
  }
  changeFindBy(value: FindBy) {
    this.patch({ findBy: value })
    this.find()
  }
  contextChanged() {
    const next = this.input.getContext()
    if (
      this.context?.stateId === next?.stateId &&
      (this.state$.value.findScope === FindScope.UNIT || this.context?.sheetId === next?.sheetId)
    )
      return
    this.context = next
    if (this.revealed && !this.replacing) this.find()
  }
  find() {
    this.release()
    if (!this.revealed || this.closed || this.composing || !this.getFindString().trim()) return
    this.busy$.next(true)
    // Coalesce changeFindString + Find All's explicit find into one request.
    const epoch = this.epoch
    this.timer = setTimeout(() => {
      this.timer = undefined
      void this.query(epoch)
    }, 0)
  }
  private async query(epoch: number) {
    const current = () => !this.closed && epoch === this.epoch && this.revealed
    try {
      if (!(await this.input.flush())) throw new Error('Finish editing before searching.')
      if (!current()) return
      const context = this.input.getContext()
      if (!context) throw new Error('The document is still loading.')
      this.context = context
      const state = { ...this.state$.value }
      this.active = { stateId: context.stateId, queryId: crypto.randomUUID() }
      let page = sessionSearchPageSchema.parse(
        await this.input.search({
          action: 'start',
          ...this.active,
          query: state.findString,
          ...(state.findScope === FindScope.SUBUNIT ? { sheetIds: [context.sheetId] } : {}),
          matchCase: state.caseSensitive,
          matchEntireCell: state.matchesTheWholeCell,
          lookIn: state.findBy === FindBy.FORMULA ? 'formulas' : 'values',
          direction: state.findDirection === FindDirection.COLUMN ? 'column' : 'row',
        }),
      )
      if (!current()) return
      this.ticket = { stateId: page.stateId, searchId: page.searchId }
      const hits: SessionSearchPage['hits'] = []
      const total = page.total
      for (;;) {
        if (
          page.stateId !== context.stateId ||
          page.searchId !== this.ticket.searchId ||
          page.offset !== hits.length ||
          page.total !== total
        )
          throw new Error('The search returned an inconsistent result page.')
        hits.push(...page.hits)
        if (page.nextOffset === undefined) break
        if (page.nextOffset !== hits.length || page.hits.length === 0)
          throw new Error('The search result cursor did not advance.')
        page = sessionSearchPageSchema.parse(
          await this.input.search({ action: 'page', ...this.ticket, offset: page.nextOffset }),
        )
        if (!current()) return
      }
      if (hits.length !== total || this.input.getContext()?.stateId !== context.stateId)
        throw new Error('The document changed before search results were complete.')
      this.matches = hits.map((hit) => {
        const point = parseAddress(hit.address)
        return {
          provider: 'genoffice.session',
          unitId: context.unitId,
          range: {
            subUnitId: hit.sheetId,
            range: {
              startRow: point.row,
              endRow: point.row,
              startColumn: point.column,
              endColumn: point.column,
            },
          },
          replaceable: hit.replaceable,
          matchedText: hit.text,
        }
      })
      this.replaceables$.next(
        this.matches.filter(
          (hit): hit is IFindMatch & { replaceable: boolean } => hit.replaceable === true,
        ),
      )
      this.busy$.next(false)
      this.patch({ matchesCount: total, findCompleted: true })
      if (total) {
        // Desktop research uses stayIfOnMatch: retain a selected matching cell
        // rather than jumping back to the beginning of a new result set.
        const selection = this.input.getContext()
        const anchor = this.matches.findIndex(
          (match) =>
            (match.range as { subUnitId: string }).subUnitId === selection?.sheetId &&
            this.addressOf(match) === selection?.address,
        )
        this.position = anchor >= 0 ? anchor - 1 : -1
        this.move(1)
      }
      if (page.valuesPending)
        this.input.report(
          new Error('Some formula values may be stale. Search formulas for their current text.'),
        )
    } catch (error) {
      if (!current()) return
      this.release()
      this.input.report(error)
    }
  }
  private move(delta: number) {
    if (!this.matches.length || this.replacing) return
    this.position = (this.position + delta + this.matches.length) % this.matches.length
    const match = this.matches[this.position]!
    this.currentMatch$.next(match)
    this.patch({ matchesPosition: this.position + 1 })
    const range = match.range as {
      subUnitId: string
      range: { startRow: number; startColumn: number }
    }
    const hit = this.ticket
    const epoch = this.epoch
    // Host query positions are current only for the ticket's state.
    if (!hit || this.input.getContext()?.stateId !== hit.stateId) {
      this.contextChanged()
      return
    }
    const address = this.addressOf(match)
    void this.input
      .jump(range.subUnitId, address, () => !this.closed && epoch === this.epoch)
      .catch((error) => {
        if (epoch === this.epoch) this.input.report(error)
      })
  }
  private addressOf(match: IFindMatch): string {
    const range = match.range as { range: { startRow: number; startColumn: number } }
    return formatAddress(range.range.startRow, range.range.startColumn)
  }
  selectMatch(sheetId: string, address: string) {
    if (!this.ticket || this.input.getContext()?.stateId !== this.ticket.stateId || this.replacing)
      return
    const index = this.matches.findIndex(
      (match) =>
        (match.range as { subUnitId: string }).subUnitId === sheetId &&
        this.addressOf(match) === address,
    )
    if (index < 0) return
    this.position = index
    this.currentMatch$.next(this.matches[index]!)
    this.patch({ matchesPosition: index + 1 })
  }
  moveToNextMatch() {
    this.move(1)
  }
  moveToPreviousMatch() {
    this.move(-1)
  }
  async replace() {
    const result = await this.replaceMatches(false)
    return result.success > 0
  }
  async replaceAll() {
    return this.replaceMatches(true)
  }
  private async replaceMatches(all: boolean) {
    const ticket = this.ticket
    if (!ticket || this.replacing || this.state$.value.inputtingFindString !== this.getFindString())
      return { success: 0, failure: 0 }
    if (!all && this.currentMatch$.value?.replaceable !== true) return { success: 0, failure: 0 }
    this.replacing = true
    this.busy$.next(true)
    this.replaceables$.next([])
    this.currentMatch$.next(null)
    try {
      if (!(await this.input.flush()) || this.input.getContext()?.stateId !== ticket.stateId)
        throw new Error('The document changed before replacement. Search again.')
      const result = replacementResult.parse(
        await this.input.search({
          action: 'replace',
          ...ticket,
          mode: all ? 'all' : 'one',
          ...(!all ? { index: this.position } : {}),
          replacement: this.state$.value.replaceString ?? '',
        }),
      )
      return { success: result.replaced, failure: result.skipped }
    } catch (error) {
      this.input.report(error)
      return { success: 0, failure: all ? this.matches.length : 1 }
    } finally {
      this.replacing = false
      if (!this.closed && this.revealed) this.find()
    }
  }
  dispose() {
    this.terminate()
    this.closed = true
    this.focusSignal$.complete()
    this.busy$.complete()
    this.state$.complete()
    this.currentMatch$.complete()
    this.replaceables$.complete()
  }
}
