import { expect, it, vi } from 'vitest'
import { initialState } from '../src/state'
import { readCurrentOutline } from '../src/read-outline'
import type { WorkbookSource } from '../src/ports'
function fixture() {
  const state = structuredClone(
    initialState({ sheets: [{ id: 'base', name: 'Data', rowCount: 60000, columnCount: 4 }] }),
  )
  const sheet = state.sheets[0]!
  sheet.features!.originalMetadata = {
    columnWidths: [{ startColumn: 1, endColumn: 2, outlineLevel: 2, hidden: true }],
    outlineSummaryBelow: false,
  }
  const readRowOutline = vi.fn(async () => ({
    indexingComplete: true,
    rows: [
      { row: 50000, level: 2, hidden: true },
      { row: 50001, level: 0, collapsed: true },
    ],
  }))
  const source = { readRowOutline } as unknown as WorkbookSource
  return { state, sheet, source, readRowOutline }
}
it('reads off-viewport source outlines and replays absolute levels and settings without changing the root', async () => {
  const f = fixture()
  f.sheet.features!.outlineEdits = [
    { op: 'set_outline', sheetId: f.sheet.id, axis: 'rows', start: 50001, count: 1, level: 3 },
    {
      op: 'set_outline',
      sheetId: f.sheet.id,
      axis: 'rows',
      start: 50002,
      count: 1,
      level: 0,
      collapsed: false,
    },
    { op: 'set_outline_settings', sheetId: f.sheet.id, summaryBelow: true, summaryRight: false },
  ]
  const before = structuredClone(f.state)
  const result = await readCurrentOutline({
    state: f.state,
    source: f.source,
    sessionId: 'source',
    sheetId: f.sheet.id,
  })
  expect(result).toMatchObject({ summaryBelow: true, summaryRight: false })
  expect(result.rows).toEqual([
    { index: 50000, level: 3, hidden: true, collapsed: false },
    { index: 50001, level: 0, hidden: false, collapsed: false },
  ])
  expect(result.cols).toEqual(
    [1, 2].map((index) => ({ index, level: 2, hidden: true, collapsed: false })),
  )
  expect(f.state).toEqual(before)
})
it('maps original outline coordinates after an earlier insertion', async () => {
  const f = fixture()
  f.state.structuralLog = [{ op: 'insert_rows', sheetId: f.sheet.id, row: 3, count: 2 }]
  const result = await readCurrentOutline({
    state: f.state,
    source: f.source,
    sessionId: 'source',
    sheetId: f.sheet.id,
  })
  expect(result.rows.map((row) => row.index)).toEqual([50002, 50003])
})
it('waits for complete indexing rather than treating partial rows as a complete model', async () => {
  const f = fixture()
  f.readRowOutline.mockResolvedValueOnce({ indexingComplete: false, rows: [] })
  const result = await readCurrentOutline({
    state: f.state,
    source: f.source,
    sessionId: 'source',
    sheetId: f.sheet.id,
  })
  expect(f.readRowOutline).toHaveBeenCalledTimes(2)
  expect(result.rows).toHaveLength(2)
})
it('rejects a foreign journal and cancellation', async () => {
  const f = fixture()
  f.sheet.features!.outlineEdits = [
    { op: 'set_outline', sheetId: 'foreign', axis: 'rows', start: 1, count: 1, level: 1 },
  ]
  await expect(
    readCurrentOutline({
      state: f.state,
      source: f.source,
      sessionId: 'source',
      sheetId: f.sheet.id,
    }),
  ).rejects.toMatchObject({ code: 'invalid_document' })
  await expect(
    readCurrentOutline({
      state: f.state,
      source: f.source,
      sessionId: 'source',
      sheetId: f.sheet.id,
      signal: AbortSignal.abort(),
    }),
  ).rejects.toBeDefined()
})

it('uses the latest authored axis visibility for source and newly grouped detail', async () => {
  const f = fixture()
  f.sheet.features!.outlineEdits = [
    { op: 'set_outline', sheetId: f.sheet.id, axis: 'rows', start: 4, count: 2, level: 1 },
  ]
  f.sheet.features!.visibilityEdits = [
    { axis: 'rows', start: 50000, count: 1, hidden: false },
    { axis: 'rows', start: 3, count: 2, hidden: true },
    { axis: 'rows', start: 3, count: 1, hidden: false },
    { axis: 'cols', start: 1, count: 1, hidden: false },
  ]
  const result = await readCurrentOutline({
    state: f.state,
    source: f.source,
    sessionId: 'source',
    sheetId: f.sheet.id,
  })
  expect(result.rows.find((row) => row.index === 3)?.hidden).toBe(false)
  expect(result.rows.find((row) => row.index === 4)?.hidden).toBe(true)
  expect(result.rows.find((row) => row.index === 50000)?.hidden).toBe(false)
  expect(result.cols.map((column) => column.hidden)).toEqual([false, true])
})

it('composes current filter hiding with outline detail and later explicit visibility', async () => {
  const f = fixture()
  f.sheet.features!.filterEdits = {
    sheetId: f.sheet.id,
    filter: {
      range: { startRow: 49999, endRow: 50001, startColumn: 0, endColumn: 2 },
      columns: [],
    },
    visibilityRange: { startRow: 49999, endRow: 50001, startColumn: 0, endColumn: 2 },
    hiddenRows: [50001],
  }
  let result = await readCurrentOutline({
    state: f.state,
    source: f.source,
    sessionId: 'source',
    sheetId: f.sheet.id,
  })
  expect(result.rows.map((row) => row.hidden)).toEqual([false, true])
  f.sheet.features!.visibilityEdits = [{ axis: 'rows', start: 50001, count: 1, hidden: false }]
  result = await readCurrentOutline({
    state: f.state,
    source: f.source,
    sessionId: 'source',
    sheetId: f.sheet.id,
  })
  expect(result.rows.map((row) => row.hidden)).toEqual([false, false])
})
