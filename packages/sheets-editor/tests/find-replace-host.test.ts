import { expect, it, vi } from 'vitest'
import { FindBy, FindDirection, FindScope } from '@univerjs/find-replace'
import { HostFindReplaceService } from '../src/find-replace-host'
import type {
  SessionSearchRequest,
  SessionSearchPage,
} from '@genoffice/xlsx-gateway/shared/session-search'

const ticket = crypto.randomUUID()
const pageOf = (offset: number, total = 201): SessionSearchPage => ({
  stateId: 'before',
  searchId: ticket,
  offset,
  total,
  valuesPending: false,
  hits: Array.from({ length: Math.min(200, total - offset) }, (_, index) => ({
    sheetId: 'sheet',
    address: `A${offset + index + 80}`,
    text: 'needle',
    truncated: false,
    isFormula: false,
    replaceable: true,
  })),
  ...(offset + 200 < total ? { nextOffset: offset + 200 } : {}),
})
function fixture(respond?: (request: SessionSearchRequest) => Promise<unknown>) {
  let context: { stateId: string; sheetId: string; unitId: string; address?: string } = {
    stateId: 'before',
    sheetId: 'sheet',
    unitId: 'unit',
  }
  const requests: SessionSearchRequest[] = [],
    errors: unknown[] = [],
    jumps: string[] = []
  const service = new HostFindReplaceService({
    async search(request) {
      requests.push(request)
      if (respond) return respond(request)
      if (request.action === 'start') return pageOf(0)
      if (request.action === 'page') return pageOf(request.offset)
      if (request.action === 'replace') {
        context = { ...context, stateId: 'after' }
        return { replaced: 201, skipped: 0 }
      }
      return {}
    },
    getContext: () => context,
    async flush() {
      return true
    },
    async jump(sheetId, address) {
      jumps.push(`${sheetId}:${address}`)
    },
    report(error) {
      errors.push(error)
    },
  })
  return {
    service,
    requests,
    errors,
    jumps,
    setContext(next: typeof context) {
      context = next
    },
  }
}

it('uses complete host pages and desktop options, navigates matches beyond a viewport and invalidates stale tickets', async () => {
  const f = fixture()
  try {
    f.service.open(false)
    f.service.changeFindString('needle')
    await vi.waitFor(() => expect(f.service.state$.value.findCompleted).toBe(true))
    expect(f.service.state$.value.matchesCount).toBe(201)
    expect(f.requests.filter((r) => r.action === 'start')).toHaveLength(1)
    expect(f.requests.some((r) => r.action === 'page' && r.offset === 200)).toBe(true)
    expect(f.jumps).toEqual(['sheet:A80'])
    f.service.moveToPreviousMatch()
    expect(f.jumps.at(-1)).toBe('sheet:A280')
    f.service.selectMatch('sheet', 'A90')
    expect(f.service.state$.value.matchesPosition).toBe(11)
    expect((f.service.currentMatch$.value!.range as any).range.startRow).toBe(89)
    f.service.changeFindScope(FindScope.UNIT)
    f.service.changeFindBy(FindBy.FORMULA)
    f.service.changeFindDirection(FindDirection.COLUMN)
    f.service.changeCaseSensitive(true)
    f.service.changeMatchesTheWholeCell(true)
    await vi.waitFor(() => expect(f.service.state$.value.findCompleted).toBe(true))
    expect(f.requests.filter((r) => r.action === 'start').at(-1)).toMatchObject({
      lookIn: 'formulas',
      direction: 'column',
      matchCase: true,
      matchEntireCell: true,
    })
    expect(f.requests.filter((r) => r.action === 'start').at(-1)).not.toHaveProperty('sheetIds')
    f.setContext({ stateId: 'after', sheetId: 'sheet', unitId: 'unit' })
    await f.service.replaceAll()
    expect(f.requests.filter((r) => r.action === 'replace')).toHaveLength(0)
    expect(String(f.errors.at(-1))).toContain('document changed')
  } finally {
    f.service.dispose()
  }
})

it('does not publish incomplete, failed, cancelled or out-of-order query results as zero matches', async () => {
  let release!: (value: unknown) => void
  const f = fixture(async (request) => {
    if (request.action === 'start')
      return new Promise((resolve) => {
        release = resolve
      })
    return {}
  })
  try {
    f.service.open(false)
    f.service.changeFindString('old')
    await vi.waitFor(() => expect(release).toBeTypeOf('function'))
    f.service.changeInputtingFindString('new')
    release(pageOf(0, 1))
    await new Promise((resolve) => setTimeout(resolve, 10))
    expect(f.service.getMatches()).toEqual([])
    expect(f.service.state$.value.findCompleted).toBe(false)
    expect(f.requests.some((r) => r.action === 'cancel')).toBe(true)
    expect(f.jumps).toEqual([])
  } finally {
    f.service.dispose()
  }
  const bad = fixture(async (request) =>
    request.action === 'start' ? { ...pageOf(0, 1), total: 2 } : {},
  )
  try {
    bad.service.open(false)
    bad.service.changeFindString('needle')
    await vi.waitFor(() => expect(bad.errors).toHaveLength(1))
    expect(bad.service.state$.value.findCompleted).toBe(false)
    expect(bad.jumps).toEqual([])
  } finally {
    bad.service.dispose()
  }
})

it('replaces from a complete ticket and prevents duplicate submission', async () => {
  let release!: () => void
  const f = fixture(async (request) => {
    if (request.action === 'start') return pageOf(0, 1)
    if (request.action === 'replace') {
      await new Promise<void>((resolve) => {
        release = resolve
      })
      return { replaced: 1, skipped: 0 }
    }
    return {}
  })
  try {
    f.service.open(true)
    f.service.changeFindString('needle')
    f.service.changeReplaceString('$&')
    await vi.waitFor(() => expect(f.service.state$.value.findCompleted).toBe(true))
    const pending = f.service.replaceAll()
    await vi.waitFor(() => expect(release).toBeTypeOf('function'))
    expect(await f.service.replaceAll()).toEqual({ success: 0, failure: 0 })
    release()
    expect(await pending).toEqual({ success: 1, failure: 0 })
    expect(f.requests.filter((r) => r.action === 'replace')).toEqual([
      {
        action: 'replace',
        stateId: 'before',
        searchId: ticket,
        mode: 'all',
        replacement: '$&',
      },
    ])
  } finally {
    f.service.dispose()
  }
})

it('defers replacement live search during IME composition and clears busy state after close', async () => {
  const f = fixture()
  try {
    f.service.open(true)
    f.service.setComposing(true)
    f.service.changeInputtingFindString('输入中')
    await new Promise((resolve) => setTimeout(resolve, 330))
    expect(f.requests.filter((request) => request.action === 'start')).toHaveLength(0)
    f.service.setComposing(false)
    f.service.changeInputtingFindString('needle')
    await vi.waitFor(() => expect(f.service.state$.value.findCompleted).toBe(true))
    expect(f.requests.filter((request) => request.action === 'start')).toHaveLength(1)
    expect(f.service.busy$.value).toBe(false)
    f.service.terminate()
    expect(f.service.revealed).toBe(false)
    expect(f.service.getMatches()).toEqual([])
    f.service.changeInputtingFindString('new uncommitted query')
    f.service.open(false)
    await vi.waitFor(() => expect(f.service.state$.value.findCompleted).toBe(true))
    expect(f.requests.filter((request) => request.action === 'start').at(-1)).toMatchObject({
      query: 'new uncommitted query',
    })
  } finally {
    f.service.dispose()
  }
})

it('keeps a matching selected cell as the initial match when a query is researched', async () => {
  const f = fixture()
  try {
    f.setContext({ stateId: 'before', sheetId: 'sheet', unitId: 'unit', address: 'A150' })
    f.service.open(false)
    f.service.changeFindString('needle')
    await vi.waitFor(() => expect(f.service.state$.value.findCompleted).toBe(true))
    expect(f.jumps).toEqual(['sheet:A150'])
    expect(f.service.state$.value.matchesPosition).toBe(71)
    f.service.moveToNextMatch()
    expect(f.jumps.at(-1)).toBe('sheet:A151')
    f.service.changeCaseSensitive(true)
    await vi.waitFor(() => expect(f.service.state$.value.findCompleted).toBe(true))
    expect(f.jumps.at(-1)).toBe('sheet:A150')
    expect(f.service.state$.value.matchesPosition).toBe(71)
  } finally {
    f.service.dispose()
  }
})
