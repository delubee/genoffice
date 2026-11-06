import { expect, it } from 'vitest'
import { initialState } from '../src/state'
import { findCurrentCells } from '../src/search'
import { WorkbookSearchController } from '../src/search-controller'
import type { WorkbookSource } from '../src/ports'
import type {
  WorkbookFindCellsRequest,
  WorkbookFindCellsResult,
} from '@genoffice/xlsx-gateway/shared/workbook-protocol'
import { sessionSearchPageSchema } from '@genoffice/xlsx-gateway/shared/session-search'

const stateOf = () =>
  structuredClone(
    initialState({
      sheets: [
        { id: 'base', name: 'Old', rowCount: 100, columnCount: 4 },
        { id: 'other', name: 'Other', rowCount: 3, columnCount: 2 },
      ],
    }),
  )
const page = (
  matches: WorkbookFindCellsResult['matches'],
  extra: Partial<WorkbookFindCellsResult> = {},
): WorkbookFindCellsResult => ({ matches, complete: true, indexingComplete: true, ...extra })
const hit = (row: number, column: number, value = 'needle', sheetId = 'base') => ({
  sheetId,
  row,
  column,
  value,
  valueText: value,
})
function sourceOf(find: (request: WorkbookFindCellsRequest) => Promise<unknown> | unknown) {
  const calls: WorkbookFindCellsRequest[] = []
  const source: WorkbookSource = {
    async openSnapshot() {
      throw new Error('unexpected open')
    },
    async readBaseRange() {
      throw new Error('search must use the complete index')
    },
    async readFeatures() {
      return {}
    },
    async archiveManifest() {
      return []
    },
    async close() {},
    async findBaseCells(request) {
      calls.push(request)
      return find(request)
    },
  }
  return { source, calls }
}

it('merges complete indexed pages with current overrides, clears and inserted/deleted coordinates without edits', async () => {
  const state = stateOf(),
    sheet = state.sheets[0]!
  state.structuralLog = [
    { op: 'insert_rows', sheetId: sheet.id, row: 2, count: 1 },
    { op: 'delete_rows', sheetId: sheet.id, row: 4, count: 1 },
  ]
  sheet.cells.A1 = { kind: 'blank', rawValue: null }
  sheet.cells.B1 = { kind: 'string', rawValue: 'needle overlay' }
  const before = structuredClone(state)
  const { source, calls } = sourceOf((request) =>
    request.resumeAt
      ? page([hit(2, 0), hit(79, 0)])
      : page([hit(0, 0), hit(1, 0)], {
          complete: false,
          nextCursor: { sheetId: 'base', row: 2, column: 0 },
        }),
  )
  const result = await findCurrentCells({
    state,
    source,
    sourceSessionId: crypto.randomUUID(),
    query: { stateId: state.stateId, query: 'needle', sheetIds: [sheet.id] },
  })
  expect(result.hits.map((hit) => hit.address)).toEqual(['B1', 'A3', 'A80'])
  expect(result.hits[0]!.value).toBe('needle overlay')
  expect(calls).toHaveLength(2)
  expect(state).toEqual(before)
  expect(result.valuesPending).toBe(true)
})

it('rewrites original formula text before matching after a sheet rename', async () => {
  const state = stateOf(),
    sheet = state.sheets[0]!
  sheet.name = 'New'
  state.structuralLog = [{ op: 'rename_sheet', sheetId: sheet.id, name: 'New' }]
  const { source, calls } = sourceOf(() => page([{ ...hit(0, 0), formulaText: '=Old!A1' }]))
  const result = await findCurrentCells({
    state,
    source,
    sourceSessionId: crypto.randomUUID(),
    query: { stateId: state.stateId, sheetIds: [sheet.id], query: 'New!', lookIn: 'formulas' },
  })
  expect(result.hits).toMatchObject([{ address: 'A1', formula: '=New!A1' }])
  expect(calls.map((call) => call.lookIn)).toEqual(['values', 'formulas_only'])
  expect(calls[1]).toMatchObject({ query: '?', wildcards: true })
})

it('supports case/whole-cell matching, direction and new-sheet overlays without an index', async () => {
  const state = stateOf(),
    sheet = state.sheets[0]!
  delete sheet.originalSheetId
  sheet.cells = {
    B1: { kind: 'string', rawValue: 'Needle' },
    A2: { kind: 'string', rawValue: ' Needle ' },
    C1: { kind: 'string', rawValue: 'needle' },
    A3: { kind: 'string', rawValue: 'Needle tail' },
  }
  const { source, calls } = sourceOf(() => {
    throw new Error('unexpected native query')
  })
  const result = await findCurrentCells({
    state,
    source,
    sourceSessionId: crypto.randomUUID(),
    query: {
      stateId: state.stateId,
      sheetIds: [sheet.id],
      query: 'Needle',
      matchCase: true,
      matchEntireCell: true,
      direction: 'column',
    },
  })
  expect(result.hits.map((hit) => hit.address)).toEqual(['A2', 'B1'])
  expect(calls).toHaveLength(0)
})

it('waits for complete indexing and rejects non-advancing/incomplete pages or foreign scope', async () => {
  const state = stateOf(),
    sheet = state.sheets[0]!
  let reads = 0
  const { source } = sourceOf(() =>
    ++reads === 1 ? page([], { indexingComplete: false, complete: false }) : page([hit(79, 0)]),
  )
  const input = {
    state,
    source,
    sourceSessionId: crypto.randomUUID(),
    query: { stateId: state.stateId, query: 'needle', sheetIds: [sheet.id] },
  }
  expect((await findCurrentCells(input)).hits[0]!.address).toBe('A80')
  const bad = sourceOf(() => page([hit(1, 0)], { complete: false }))
  await expect(findCurrentCells({ ...input, source: bad.source })).rejects.toMatchObject({
    code: 'incomplete_search',
  })
  const repeated = sourceOf(() =>
    page([hit(1, 0)], { complete: false, nextCursor: { sheetId: 'base', row: 1, column: 0 } }),
  )
  await expect(findCurrentCells({ ...input, source: repeated.source })).rejects.toMatchObject({
    code: 'incomplete_search',
  })
  await expect(
    findCurrentCells({ ...input, query: { ...input.query, sheetIds: ['foreign'] } }),
  ).rejects.toMatchObject({ code: 'sheet_not_found' })
  const foreign = sourceOf(() => page([hit(1, 0, 'needle', 'other')]))
  await expect(findCurrentCells({ ...input, source: foreign.source })).rejects.toMatchObject({
    code: 'session_mismatch',
  })
})

it('pages a bounded preview ticket, rejects extra source handles and invalidates it after revisions', async () => {
  let state = stateOf()
  const { source } = sourceOf((request) =>
    request.sheetId === 'base'
      ? page(Array.from({ length: 601 }, (_, row) => hit(row, 0, 'needle' + 'x'.repeat(600))))
      : page([]),
  )
  const controller = new WorkbookSearchController({
    source,
    sourceSessionId: crypto.randomUUID(),
    getState: () => state,
  })
  const request = {
    action: 'start',
    queryId: crypto.randomUUID(),
    stateId: state.stateId,
    query: 'needle',
    lookIn: 'formulas',
  }
  await expect(controller.handle({ ...request, sourcePath: '/private.xlsx' })).rejects.toThrow()
  const first = sessionSearchPageSchema.parse(await controller.handle(request))
  expect(first).toMatchObject({ total: 601, nextOffset: 200, offset: 0 })
  expect(first.hits).toHaveLength(200)
  expect(first.hits[0]).toMatchObject({ truncated: true, replaceable: true })
  expect(first.hits[0]!.text).toHaveLength(512)
  const last = sessionSearchPageSchema.parse(
    await controller.handle({
      action: 'page',
      searchId: first.searchId,
      stateId: first.stateId,
      offset: 600,
    }),
  )
  expect(last.hits).toHaveLength(1)
  expect(last.nextOffset).toBeUndefined()
  state = { ...state, stateId: crypto.randomUUID() }
  await expect(
    controller.handle({
      action: 'page',
      searchId: first.searchId,
      stateId: first.stateId,
      offset: 200,
    }),
  ).rejects.toMatchObject({ code: 'revision_conflict' })
  expect(
    await controller.handle({
      action: 'release',
      searchId: first.searchId,
      stateId: first.stateId,
    }),
  ).toEqual({ released: true })
  controller.dispose()
  await expect(controller.handle(request)).rejects.toMatchObject({ code: 'view_closed' })
})

it('cancels only the identified pending query and never publishes its late response', async () => {
  const state = stateOf()
  let finish!: (value: unknown) => void
  const { source } = sourceOf(
    () =>
      new Promise((resolve) => {
        finish = resolve
      }),
  )
  const controller = new WorkbookSearchController({
    source,
    sourceSessionId: crypto.randomUUID(),
    getState: () => state,
  })
  const queryId = crypto.randomUUID()
  const pending = controller.handle({
    action: 'start',
    queryId,
    stateId: state.stateId,
    query: 'needle',
    sheetIds: [state.sheets[0]!.id],
  })
  expect(
    await controller.handle({
      action: 'cancel',
      queryId: crypto.randomUUID(),
      stateId: state.stateId,
    }),
  ).toEqual({ cancelled: false })
  expect(await controller.handle({ action: 'cancel', queryId, stateId: state.stateId })).toEqual({
    cancelled: true,
  })
  finish(page([hit(0, 0)]))
  await expect(pending).rejects.toMatchObject({ code: 'revision_conflict' })
})

it('does not report a complete result when the bounded session match budget is exceeded', async () => {
  const state = stateOf()
  const { source } = sourceOf((request) => {
    const start = request.resumeAt?.row ?? 0
    const length = Math.min(10_000, 100_001 - start)
    return page(
      Array.from({ length }, (_, offset) => hit(start + offset, 0)),
      start + length < 100_001
        ? { complete: false, nextCursor: { sheetId: 'base', row: start + length, column: 0 } }
        : {},
    )
  })
  const controller = new WorkbookSearchController({
    source,
    sourceSessionId: crypto.randomUUID(),
    getState: () => state,
  })
  await expect(
    controller.handle({
      action: 'start',
      queryId: crypto.randomUUID(),
      stateId: state.stateId,
      query: 'needle',
      sheetIds: [state.sheets[0]!.id],
    }),
  ).rejects.toMatchObject({ code: 'range_limit' })
  controller.dispose()
})

it('rejects an indexed response overtaken by another document revision', async () => {
  let current = stateOf()
  let finish!: (value: unknown) => void
  const { source } = sourceOf(
    () =>
      new Promise((resolve) => {
        finish = resolve
      }),
  )
  const controller = new WorkbookSearchController({
    source,
    sourceSessionId: crypto.randomUUID(),
    getState: () => current,
  })
  const pending = controller.handle({
    action: 'start',
    queryId: crypto.randomUUID(),
    stateId: current.stateId,
    query: 'needle',
    sheetIds: [current.sheets[0]!.id],
  })
  current = { ...current, stateId: crypto.randomUUID() }
  finish(page([hit(79, 0)]))
  await expect(pending).rejects.toMatchObject({ code: 'revision_conflict' })
  controller.dispose()
})

it('replaces full untruncated source text with literal dollar signs in one consumed ticket', async () => {
  const state = stateOf(),
    sheetId = state.sheets[0]!.id
  const text = 'prefix'.repeat(120) + 'Needle suffix'
  const { source } = sourceOf(() => page([hit(79, 0, text)]))
  const commits: unknown[] = []
  const controller = new WorkbookSearchController({
    source,
    sourceSessionId: 'original',
    getState: () => state,
    async replace(stateId, operations) {
      commits.push({ stateId, operations })
    },
  })
  const result = sessionSearchPageSchema.parse(
    await controller.handle({
      action: 'start',
      queryId: crypto.randomUUID(),
      stateId: state.stateId,
      query: 'needle',
      sheetIds: [sheetId],
    }),
  )
  expect(result.hits[0]).toMatchObject({
    text: text.slice(0, 512),
    truncated: true,
    replaceable: true,
  })
  const request = {
    action: 'replace',
    stateId: state.stateId,
    searchId: result.searchId,
    mode: 'one',
    index: 0,
    replacement: '$&',
  }
  expect(await controller.handle(request)).toEqual({ replaced: 1, skipped: 0 })
  expect(commits).toEqual([
    {
      stateId: state.stateId,
      operations: [
        { op: 'set_cell', sheetId, address: 'A80', value: 'prefix'.repeat(120) + '$& suffix' },
      ],
    },
  ])
  await expect(controller.handle(request)).rejects.toMatchObject({ code: 'search_not_found' })
  expect(commits).toHaveLength(1)
  expect(state.sheets[0]!.cells).toEqual({})
})

it('preserves numeric and boolean replacement types, skips formula values and replaces formula text only in formula mode', async () => {
  const state = stateOf(),
    sheet = state.sheets[0]!
  sheet.cells.A1 = { kind: 'number', rawValue: 120 }
  sheet.cells.B1 = { kind: 'boolean', rawValue: true }
  sheet.cells.C1 = {
    kind: 'number',
    rawValue: 120,
    formula: '=120',
    calculationState: 'ready',
    cachedValue: 120,
  }
  const { source } = sourceOf(() => page([]))
  const commits: unknown[] = []
  const controller = new WorkbookSearchController({
    source,
    sourceSessionId: 'original',
    getState: () => state,
    async replace(_stateId, operations) {
      commits.push(operations)
    },
  })
  const start = async (query: string, lookIn = 'values') =>
    sessionSearchPageSchema.parse(
      await controller.handle({
        action: 'start',
        queryId: crypto.randomUUID(),
        stateId: state.stateId,
        query,
        lookIn,
        sheetIds: [sheet.id],
      }),
    )
  const replace = (searchId: string, replacement: string) =>
    controller.handle({
      action: 'replace',
      stateId: state.stateId,
      searchId,
      mode: 'all',
      replacement,
    })
  const numbers = await start('120')
  expect(numbers.valuesPending).toBe(true)
  expect(numbers.hits.map((hit) => hit.replaceable)).toEqual([true, false])
  expect(await replace(numbers.searchId, '240')).toEqual({ replaced: 1, skipped: 1 })
  expect(commits.at(-1)).toEqual([{ op: 'set_cell', sheetId: sheet.id, address: 'A1', value: 240 }])
  const boolean = await start('1')
  await replace(boolean.searchId, '0')
  expect(commits.at(-1) as unknown[]).toContainEqual({
    op: 'set_cell',
    sheetId: sheet.id,
    address: 'B1',
    value: false,
  })
  const formulas = await start('120', 'formulas')
  await replace(formulas.searchId, '360')
  expect(commits.at(-1)).toEqual([
    { op: 'set_cell', sheetId: sheet.id, address: 'A1', value: 360 },
    { op: 'set_formula', sheetId: sheet.id, address: 'C1', formula: '=360' },
  ])
})

it('rejects oversized, protected, stale and missing-index replacements before any submission', async () => {
  let state = stateOf()
  const { source } = sourceOf(() => page(Array.from({ length: 2001 }, (_, index) => hit(index, 0))))
  let commits = 0
  const controller = new WorkbookSearchController({
    source,
    sourceSessionId: 'original',
    getState: () => state,
    async replace() {
      commits++
    },
  })
  const start = async () =>
    sessionSearchPageSchema.parse(
      await controller.handle({
        action: 'start',
        queryId: crypto.randomUUID(),
        stateId: state.stateId,
        query: 'needle',
        sheetIds: [state.sheets[0]!.id],
      }),
    )
  const first = await start()
  const request = {
    action: 'replace',
    stateId: state.stateId,
    searchId: first.searchId,
    mode: 'all',
    replacement: 'done',
  }
  await expect(controller.handle(request)).rejects.toMatchObject({ code: 'range_limit' })
  await expect(controller.handle({ ...request, mode: 'one' })).rejects.toMatchObject({
    code: 'invalid_request',
  })
  state.sheets[0]!.protected = true
  await expect(controller.handle({ ...request, mode: 'one', index: 0 })).rejects.toMatchObject({
    code: 'readonly_workbook',
  })
  state = { ...state, stateId: crypto.randomUUID() }
  await expect(controller.handle({ ...request, mode: 'one', index: 0 })).rejects.toMatchObject({
    code: 'revision_conflict',
  })
  expect(commits).toBe(0)
})
