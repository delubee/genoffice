import { expect, it } from 'vitest'
import { initialState } from '../src/state'
import { auditWorkbook } from '../src/audit'
import { findCurrentCells } from '../src/search'
import type { WorkbookSource } from '../src/ports'
const fixture = () => {
  let state = structuredClone(
    initialState({
      sheets: [
        { id: 'base', name: 'Data', rowCount: 100001, columnCount: 5 },
        { id: 'other', name: 'Other', rowCount: 2, columnCount: 2 },
      ],
    }),
  )
  const source: WorkbookSource = {
    async openSnapshot() {
      throw new Error('unexpected')
    },
    async readBaseRange() {
      throw new Error('audit must query the complete index')
    },
    async readFeatures() {
      return {}
    },
    async close() {},
    async archiveManifest() {
      return []
    },
    async findBaseCells(request) {
      return {
        complete: true,
        indexingComplete: true,
        matches: (request.sheetId === 'base'
          ? [
              { row: 100000, column: 2, value: 121, formulaText: '=B2+1' },
              { row: 100000, column: 3, value: '#DIV/0!', formulaText: '=1/0' },
              { row: 100000, column: 4, value: '#ordinary-text' },
              { row: 1, column: 4, value: 'B2', formulaText: '="B2"' },
            ]
          : [
              { row: 0, column: 0, value: 120, formulaText: '=Data!B2' },
              { row: 1, column: 1, value: '#REF!' },
            ]
        ).map((hit) => ({ sheetId: request.sheetId, valueText: String(hit.value), ...hit })),
      }
    },
  }
  return {
    get state() {
      return state
    },
    source,
    replaceRoot() {
      state = { ...state, stateId: crypto.randomUUID() }
    },
    run(kind: 'dependents' | 'errors') {
      return auditWorkbook({
        request: { stateId: state.stateId, kind, sheetId: state.sheets[0]!.id, address: 'B2' },
        getState: () => state,
        source,
        sourceSessionId: crypto.randomUUID(),
      })
    },
  }
}
it('audits original off-page and cross-sheet formulas using the shared reference parser without authoring cells', async () => {
  const f = fixture(),
    before = structuredClone(f.state)
  const result = await f.run('dependents')
  expect(result.hits).toEqual([
    { sheetId: f.state.sheets[0]!.id, address: 'C100001' },
    { sheetId: f.state.sheets[1]!.id, address: 'A1' },
  ])
  expect(f.state).toEqual(before)
})
it('finds true error values across the full source and overlays while honoring clears and pending values', async () => {
  const f = fixture()
  expect((await f.run('errors')).hits.map((hit) => hit.address)).toEqual(['D100001', 'B2'])
  f.state.sheets[0]!.cells.D100001 = { kind: 'blank', rawValue: null }
  f.state.sheets[0]!.cells.C5 = { kind: 'error', rawValue: '#VALUE!' }
  const result = await f.run('errors')
  expect(result.hits.map((hit) => hit.address)).toEqual(['C5', 'B2'])
  expect(result.valuesPending).toBe(true)
})
it('rejects a source audit overtaken by another root before publishing results', async () => {
  const f = fixture(),
    original = f.source.findBaseCells!
  f.source.findBaseCells = async (request) => {
    const result = await original(request)
    f.replaceRoot()
    return result
  }
  await expect(f.run('dependents')).rejects.toThrow('changed')
})
it('searches rewritten original name references instead of their stale source spelling', async () => {
  const f = fixture()
  f.source.findBaseCells = async (request) => ({
    complete: true,
    indexingComplete: true,
    matches:
      request.sheetId === 'base'
        ? [
            {
              sheetId: 'base',
              row: 100000,
              column: 2,
              value: 1,
              formulaText: '=Revenue+1',
              valueText: '1',
            },
          ]
        : [],
  })
  f.state.features = {
    definedNamesEdits: { names: [], preserveNames: [] },
    definedNameReferences: [{ name: 'Revenue', replacement: 'Sales', shadowedSheetIds: [] }],
  }
  const result = await findCurrentCells({
    state: f.state,
    source: f.source,
    sourceSessionId: crypto.randomUUID(),
    query: { stateId: f.state.stateId, query: 'Sales', lookIn: 'formulas' },
  })
  expect(result.hits).toMatchObject([{ address: 'C100001', text: '=Sales+1' }])
  expect((await f.run('errors')).valuesPending).toBe(true)
})
