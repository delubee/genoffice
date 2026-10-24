import { expect, it } from 'vitest'
import { initialState } from '../src/state'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'

function fixture() {
  let state = structuredClone(
    initialState({ sheets: [{ id: 'base', name: 'Data', rowCount: 3, columnCount: 2 }] }),
  )
  const sheetId = state.sheets[0]!.id
  state.sheets[0]!.cells = {
    A1: { kind: 'string', rawValue: 'Region' },
    B1: { kind: 'string', rawValue: 'Sales' },
    A2: { kind: 'string', rawValue: 'North' },
    B2: { kind: 'number', rawValue: 12.5, numberFormat: '0.00', style: { numberFormat: '0.00' } },
    A3: { kind: 'string', rawValue: 'South' },
    B3: { kind: 'number', rawValue: 8 },
  }
  return {
    sheetId,
    get state() {
      return state
    },
    async apply(operations: WorkbookOperation[]) {
      const before = structuredClone(state)
      const result = await computeCandidateState(
        state,
        {
          schemaVersion: 1,
          sessionId: 'ribbon',
          transactionId: crypto.randomUUID(),
          expectedRevision: 0,
          origin: 'manual',
          summary: 'Ribbon action',
          operations,
        },
        {
          state: structuredClone(state),
          completeSheetIds: [sheetId],
          writableSheetIds: [sheetId],
          allocation: { stateId: crypto.randomUUID(), sheetIds: {} },
        },
      )
      expect(result.inverse.state).toEqual(before)
      state = result.nextState
      return result
    },
  }
}
it('commits page, view, freeze and note settings atomically without cell rewrites', async () => {
  const f = fixture(),
    before = structuredClone(f.state.sheets[0]!.cells)
  await f.apply([
    {
      op: 'set_page_setup',
      sheetId: f.sheetId,
      orientation: 'landscape',
      paperSize: 8,
      margins: 'narrow',
      showGridlines: false,
      zoomScale: 125,
    },
    { op: 'set_freeze', sheetId: f.sheetId, rows: 1, columns: 0 },
    { op: 'set_note', sheetId: f.sheetId, address: 'B2', text: 'Check & approve' },
  ])
  expect(f.state.sheets[0]!.features).toMatchObject({
    pageSetupEdits: {
      orientation: 'landscape',
      paperSize: 8,
      showGridlines: false,
      zoomScale: 125,
    },
    freezeEdits: { frozenRows: 1, frozenColumns: 0 },
    noteEdits: { B2: 'Check & approve' },
  })
  expect(f.state.sheets[0]!.cells).toEqual(before)
  await f.apply([{ op: 'set_note', sheetId: f.sheetId, address: 'B2', text: null }])
  expect(f.state.sheets[0]!.features?.noteEdits).toEqual({ B2: null })
})
it('filters by current formatted values, rejects filtered sorting and clears row visibility', async () => {
  const f = fixture()
  await f.apply([
    { op: 'set_filter', sheetId: f.sheetId, range: 'A1:B3' },
    { op: 'set_filter_criteria', sheetId: f.sheetId, column: 'B', values: ['12.50'] },
  ])
  expect(f.state.sheets[0]!.features?.filterEdits).toMatchObject({ hiddenRows: [2] })
  const before = JSON.stringify(f.state)
  await expect(
    f.apply([
      {
        op: 'sort_range',
        sheetId: f.sheetId,
        range: 'A1:B3',
        byColumn: 'B',
        order: 'asc',
        hasHeader: true,
      },
    ]),
  ).rejects.toThrow('active_filter_sort')
  expect(JSON.stringify(f.state)).toBe(before)
  await f.apply([{ op: 'clear_filter', sheetId: f.sheetId }])
  expect(f.state.sheets[0]!.features?.filterEdits).toMatchObject({ filter: null, hiddenRows: [] })
})
it('sorting moves typed dates and their formats instead of coercing them to ordinary numbers', async () => {
  const f = fixture()
  f.state.sheets[0]!.cells.A3 = {
    kind: 'date',
    rawValue: 60,
    dateSystem: '1900',
    style: { numberFormat: 'yyyy-mm-dd', bold: true },
    numberFormat: 'yyyy-mm-dd',
  }
  await f.apply([
    {
      op: 'sort_range',
      sheetId: f.sheetId,
      range: 'A1:B3',
      byColumn: 'B',
      order: 'asc',
      hasHeader: true,
    },
  ])
  expect(f.state.sheets[0]!.cells.A2).toMatchObject({
    kind: 'date',
    rawValue: 60,
    dateSystem: '1900',
    numberFormat: 'yyyy-mm-dd',
    style: { bold: true },
  })
  expect(f.state.sheets[0]!.cells.B2!.rawValue).toBe(8)
})
it('keeps image notes and formula sorts unchanged when unsupported', async () => {
  const f = fixture()
  f.state.sheets[0]!.features = { originalMetadata: { comments: [{ image: { id: 'picture' } }] } }
  await expect(
    f.apply([{ op: 'set_note', sheetId: f.sheetId, address: 'B2', text: 'new' }]),
  ).rejects.toThrow('unsupported_note_image')
  f.state.sheets[0]!.cells.B2 = {
    kind: 'number',
    rawValue: 2,
    formula: '=1+1',
    calculationState: 'ready',
  }
  await expect(
    f.apply([
      {
        op: 'sort_range',
        sheetId: f.sheetId,
        range: 'A1:B3',
        byColumn: 'B',
        order: 'asc',
        hasHeader: true,
      },
    ]),
  ).rejects.toThrow()
})

it('uses the workbook date system when matching displayed filter values', async () => {
  const f = fixture()
  f.state.dateSystem = '1904'
  f.state.sheets[0]!.cells.B2 = {
    kind: 'date',
    rawValue: 1,
    dateSystem: '1904',
    numberFormat: 'yyyy-mm-dd',
    style: { numberFormat: 'yyyy-mm-dd' },
  }
  await f.apply([
    { op: 'set_filter', sheetId: f.sheetId, range: 'B1:B3' },
    { op: 'set_filter_criteria', sheetId: f.sheetId, column: 'B', values: ['1904-01-02'] },
  ])
  expect(f.state.sheets[0]!.features?.filterEdits).toMatchObject({ hiddenRows: [2] })
})

it('sorts Han text by an explicit pinyin collation independently of the host UI language', async () => {
  const f = fixture()
  f.state.sheets[0]!.cells.A2 = { kind: 'string', rawValue: '中' }
  f.state.sheets[0]!.cells.A3 = { kind: 'string', rawValue: '阿' }
  await f.apply([
    {
      op: 'sort_range',
      sheetId: f.sheetId,
      range: 'A1:B3',
      byColumn: 'A',
      order: 'asc',
      hasHeader: true,
    },
  ])
  expect(f.state.sheets[0]!.cells.A2!.rawValue).toBe('阿')
  expect(f.state.sheets[0]!.cells.A3!.rawValue).toBe('中')
  expect(f.state.sheets[0]!.cells.B2!.rawValue).toBe(8)
})
