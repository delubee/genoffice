import { expect, it, vi } from 'vitest'
import { randomUUID } from 'node:crypto'
import { loadCalculationSource } from '../src/calculation/source'
import { initialState } from '../src/state'
import type { WorkbookSource } from '../src/ports'
import type { WorkbookFindCellsResult } from '@genoffice/xlsx-gateway/shared/workbook-protocol'
import { parseRange } from '@genoffice/xlsx-gateway/domain/cell-address'

const stateOf = () =>
  structuredClone(
    initialState({ sheets: [{ id: 'base', name: 'Data', rowCount: 1, columnCount: 1 }] }),
  )
const hit = (row: number, column: number, extra = {}) => ({
  sheetId: 'base',
  row,
  column,
  value: null,
  valueText: null,
  ...extra,
})
const page = (
  matches: WorkbookFindCellsResult['matches'],
  extra = {},
): WorkbookFindCellsResult => ({ matches, complete: true, indexingComplete: true, ...extra })
function fixture() {
  const cells = [
    { row: 0, column: 0, value: 10 },
    { row: 100000, column: 1, value: null, styleIndex: 1 },
    { row: 100001, column: 2, value: '#REF!', isError: true },
    { row: 100002, column: 3, value: 60, styleIndex: 2 },
    { row: 100003, column: 0, value: null, formula: '=A1' },
  ]
  const read = vi.fn(async (_session: string, _sheet: string, range: string) => {
    const b = parseRange(range)
    return {
      cells: cells.filter(
        (c) =>
          c.row >= b.startRow &&
          c.row <= b.endRow &&
          c.column >= b.startColumn &&
          c.column <= b.endColumn,
      ),
    }
  })
  const source: WorkbookSource = {
    async readCalculationContext(sessionId) {
      return { sessionId, names: [], calculationProperties: {}, externalRelationshipIds: [] }
    },
    async openSnapshot() {
      throw new Error('unused')
    },
    async close() {},
    async archiveManifest() {
      return {}
    },
    async readFeatures() {
      return {
        indexingComplete: true,
        styles: [
          {},
          { background: '#FF0000', numberFormat: '0.00' },
          { numberFormat: 'yyyy-mm-dd' },
        ],
      }
    },
    readBaseRange: read,
    findBaseCells: vi.fn(async (request) =>
      request.resumeAt
        ? page(
            cells.slice(2).map((c) => hit(c.row, c.column)),
            {},
          )
        : page(
            cells.slice(0, 2).map((c) => hit(c.row, c.column)),
            { complete: false, nextCursor: { sheetId: 'base', row: 100001, column: 2 } },
          ),
    ),
  }
  return { source, read, cells }
}
const request = (source: WorkbookSource, state = stateOf()) => ({
  source,
  state,
  sourceSessionId: randomUUID(),
  isCurrent: () => true,
})

it('loads far source cells, formatted blanks, typed errors/dates and formula caches without trusting dimensions or changing the root', async () => {
  const { source, read } = fixture(),
    input = request(source),
    before = structuredClone(input.state)
  const result = await loadCalculationSource(input)
  expect(input.state).toEqual(before)
  expect(result.sheets[0]!.cells).toMatchObject({
    A1: { rawValue: 10, kind: 'number' },
    B100001: { rawValue: null, kind: 'blank', numberFormat: '0.00' },
    C100002: { rawValue: '#REF!', kind: 'error' },
    D100003: { rawValue: 60, kind: 'date', dateSystem: '1900' },
    A100004: { formula: '=A1', calculationState: 'pending' },
  })
  expect(Object.keys(result.sheets[0]!.cells)).toHaveLength(5)
  expect(result.sheets[0]!.rows).toBe(100004)
  expect(read).toHaveBeenCalledTimes(2)
  expect(
    vi.mocked(source.findBaseCells!).mock.calls.every(([r]) => r.lookIn === 'stored_cells'),
  ).toBe(true)
})

it('maps insertion/deletion, applies explicit clears and styles, loads new sheets and excludes deleted sheets', async () => {
  const { source } = fixture(),
    state = stateOf(),
    sheet = state.sheets[0]!
  state.structuralLog = [
    { op: 'insert_rows', sheetId: sheet.id, row: 1, count: 1 },
    { op: 'delete_rows', sheetId: sheet.id, row: 100003, count: 1 },
  ]
  sheet.cells.A2 = { kind: 'blank', rawValue: null }
  sheet.cells.B100002 = { kind: 'number', rawValue: 42, style: { bold: true } }
  const added = structuredClone(sheet)
  added.id = randomUUID()
  added.name = 'Added'
  delete added.originalSheetId
  added.cells = { Z50: { kind: 'boolean', rawValue: false } }
  state.sheets.push(added)
  const result = await loadCalculationSource(request(source, state))
  expect(result.sheets[0]!.cells.A2).toMatchObject({ kind: 'blank', rawValue: null })
  expect(result.sheets[0]!.cells.B100002).toMatchObject({ rawValue: 42, style: { bold: true } })
  expect(result.sheets[0]!.cells.C100003).toBeUndefined()
  expect(result.sheets[1]!.cells).toEqual(added.cells)
  expect(vi.mocked(source.findBaseCells!).mock.calls.every(([r]) => r.sheetId === 'base')).toBe(
    true,
  )
})

it('retains source array ownership even when its anchor is overridden and records deleted anchors', async () => {
  const { source } = fixture(),
    state = stateOf(),
    sheet = state.sheets[0]!
  source.findBaseCells = async () => page([hit(0, 0, { arrayRef: 'A1:B2' })])
  sheet.cells.A1 = { kind: 'number', rawValue: 99 }
  const result = await loadCalculationSource(request(source, state))
  expect(result.sheets[0]!.cells.A1!.rawValue).toBe(99)
  expect(result.sheets[0]!.sourceArrays).toEqual([
    { sourceAddress: 'A1', sourceRange: 'A1:B2', logicalAddress: 'A1' },
  ])
  state.structuralLog = [{ op: 'delete_rows', sheetId: sheet.id, row: 1, count: 1 }]
  expect(
    (await loadCalculationSource(request(source, state))).sheets[0]!.sourceArrays[0]!
      .logicalAddress,
  ).toBeNull()
})

it('rejects stale roots and cancellation after asynchronous reads with no partial result', async () => {
  const { source } = fixture(),
    input = request(source)
  let current = true
  source.findBaseCells = async () => {
    current = false
    return page([hit(0, 0)])
  }
  await expect(loadCalculationSource({ ...input, isCurrent: () => current })).rejects.toMatchObject(
    { code: 'revision_conflict' },
  )
  const controller = new AbortController()
  source.findBaseCells = async () => {
    controller.abort()
    return page([hit(0, 0)])
  }
  await expect(
    loadCalculationSource({ ...input, signal: controller.signal }),
  ).rejects.toMatchObject({ code: 'operation_cancelled' })
})

it('enforces cell/byte budgets and bounded sparse source reads', async () => {
  const { source, read } = fixture(),
    input = request(source)
  await expect(loadCalculationSource({ ...input, maxCells: 4 })).rejects.toMatchObject({
    code: 'range_limit',
  })
  await expect(loadCalculationSource({ ...input, maxBytes: 100 })).rejects.toMatchObject({
    code: 'range_limit',
  })
  await loadCalculationSource({ ...input, readBatchCells: 1 })
  expect(
    read.mock.calls.every(([, , range]) => {
      const r = parseRange(range)
      return r.startRow === r.endRow && r.startColumn === r.endColumn
    }),
  ).toBe(true)
  await expect(loadCalculationSource({ ...input, readBatchCells: 0 })).rejects.toMatchObject({
    code: 'invalid_range',
  })
})

it.each([
  page([hit(0, 0)], { complete: false }),
  page([hit(0, 0)], { complete: false, nextCursor: { sheetId: 'base', row: 0, column: 0 } }),
  page([hit(0, 0)], { complete: false, nextCursor: { sheetId: 'foreign', row: 1, column: 0 } }),
  page([hit(0, 0), hit(0, 0)]),
  page([hit(1, 0), hit(0, 0)]),
  page([hit(0, 0, { sheetId: 'foreign' })]),
  page([hit(1048576, 0)]),
  page([hit(0, 16384)]),
  page([hit(0, 0, { arrayRef: 'B2:C3' })]),
])('rejects incomplete, non-advancing and invalid source enumeration %#', async (response) => {
  const { source } = fixture()
  source.findBaseCells = async () => response
  await expect(loadCalculationSource(request(source))).rejects.toMatchObject({
    code: 'incomplete_source',
  })
})

it('does not consume partially indexed pages and retains the same resume position when retrying', async () => {
  const { source } = fixture()
  let reads = 0
  source.findBaseCells = vi.fn(async () =>
    ++reads === 1
      ? page([hit(0, 0)], { complete: false, indexingComplete: false })
      : page([hit(0, 0)]),
  )
  expect((await loadCalculationSource(request(source))).sheets[0]!.cells.A1!.rawValue).toBe(10)
  expect(vi.mocked(source.findBaseCells).mock.calls.map(([r]) => r.resumeAt)).toEqual([
    undefined,
    undefined,
  ])
})

it('rejects missing source records rather than turning them into blank calculation inputs', async () => {
  const { source } = fixture()
  source.readBaseRange = async () => ({ cells: [] })
  await expect(loadCalculationSource(request(source))).rejects.toMatchObject({
    code: 'incomplete_source',
  })
})

it('checks root ownership after style and range reads, including source implementations with private receivers', async () => {
  for (const stage of ['features', 'range']) {
    const { source } = fixture()
    let current = true
    const readFeatures = source.readFeatures.bind(source),
      readBaseRange = source.readBaseRange.bind(source)
    class BoundSource implements WorkbookSource {
      #valid = true
      readCalculationContext = source.readCalculationContext
      openSnapshot = source.openSnapshot
      close = source.close
      archiveManifest = source.archiveManifest
      findBaseCells = source.findBaseCells
      async readFeatures(session: string, sheet: string) {
        expect(this.#valid).toBe(true)
        const result = await readFeatures(session, sheet)
        if (stage === 'features') current = false
        return result
      }
      async readBaseRange(session: string, sheet: string, range: string) {
        expect(this.#valid).toBe(true)
        const result = await readBaseRange(session, sheet, range)
        if (stage === 'range') current = false
        return result
      }
    }
    await expect(
      loadCalculationSource({ ...request(new BoundSource()), isCurrent: () => current }),
    ).rejects.toMatchObject({ code: 'revision_conflict' })
  }
})

it('retains host-private name/scope/settings resources, checks their owner, and counts them in the byte budget', async () => {
  const { source } = fixture(),
    input = request(source)
  const context = {
    sessionId: input.sourceSessionId,
    names: [{ name: 'Secret', formula: '42', hidden: true, sheetId: 'base' }],
    calculationProperties: { calcMode: 'manual', iterate: '1' },
    externalRelationshipIds: ['external'],
  }
  source.readCalculationContext = async () => context
  expect((await loadCalculationSource(input)).sourceContext).toEqual(context)
  await expect(loadCalculationSource({ ...input, maxBytes: 100 })).rejects.toMatchObject({
    code: 'range_limit',
  })
  source.readCalculationContext = async () => ({ ...context, sessionId: randomUUID() })
  await expect(loadCalculationSource(input)).rejects.toMatchObject({ code: 'session_mismatch' })
  source.readCalculationContext = undefined
  await expect(loadCalculationSource(input)).rejects.toMatchObject({
    code: 'calculation_unavailable',
  })
})

it('rejects stale ownership and abort after private context reads before enumerating cells', async () => {
  const { source } = fixture(),
    input = request(source)
  let current = true
  source.readCalculationContext = async (sessionId) => {
    current = false
    return { sessionId, names: [], calculationProperties: {}, externalRelationshipIds: [] }
  }
  await expect(loadCalculationSource({ ...input, isCurrent: () => current })).rejects.toMatchObject(
    { code: 'revision_conflict' },
  )
  expect(source.findBaseCells).not.toHaveBeenCalled()
  const controller = new AbortController()
  source.readCalculationContext = async (sessionId) => {
    controller.abort()
    return { sessionId, names: [], calculationProperties: {}, externalRelationshipIds: [] }
  }
  await expect(
    loadCalculationSource({ ...input, signal: controller.signal }),
  ).rejects.toMatchObject({ code: 'operation_cancelled' })
  expect(source.findBaseCells).not.toHaveBeenCalled()
})
