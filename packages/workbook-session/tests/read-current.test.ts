import { expect, it, vi } from 'vitest'
import { readCurrentRange } from '../src/read'
import { immutable, initialState } from '../src/state'
import type { WorkbookSource } from '../src/ports'

const state = initialState({
  sheets: [{ id: 'source-sheet', name: 'Data', rowCount: 10, columnCount: 3 }],
})
const sheetId = state.sheets[0]!.id
const source: WorkbookSource = {
  async openSnapshot() {
    throw new Error('Unused')
  },
  async archiveManifest() {
    return {}
  },
  async close() {},
  async readFeatures() {
    return { styles: [{ numberFormat: 'General' }, { numberFormat: 'yyyy-mm-dd' }] }
  },
  async readBaseRange() {
    return {
      cells: [
        { row: 0, column: 0, value: 10 },
        { row: 1, column: 0, value: '', styleIndex: 0 },
        { row: 0, column: 1, value: 60, styleIndex: 1 },
        { row: 0, column: 2, value: '#REF!', isError: true },
      ],
    }
  },
}
it('clips source reads to original dimensions while preserving editable blank grid space', async () => {
  const current = structuredClone(state)
  current.sheets[0]!.cells.E20 = { kind: 'number', rawValue: 42 }
  const readBaseRange = vi.fn(source.readBaseRange)
  const boundedSource = { ...source, readBaseRange }
  const view = await readCurrentRange({
    state: current,
    sessionId: 'source',
    source: boundedSource,
    logicalSheetId: sheetId,
    range: 'A1:T50',
  })
  expect(readBaseRange).toHaveBeenCalledExactlyOnceWith('source', 'source-sheet', 'A1:C10')
  expect(view.cells.find((cell) => cell.address === 'E20')?.rawValue).toBe(42)
  expect(view.cells.find((cell) => cell.address === 'T50')?.kind).toBe('blank')
  readBaseRange.mockClear()
  await readCurrentRange({
    state: current,
    sessionId: 'source',
    source: boundedSource,
    logicalSheetId: sheetId,
    range: 'D11:T50',
  })
  expect(readBaseRange).not.toHaveBeenCalled()
})
it('reads an immutable current overlay, typed empty strings, dates, errors and blank cells', async () => {
  const current = structuredClone(state)
  current.sheets[0]!.cells.A1 = { kind: 'string', rawValue: '00123' }
  const view = await readCurrentRange({
    state: immutable(current),
    sessionId: 'source',
    source,
    logicalSheetId: sheetId,
    range: 'A1:C2',
  })
  expect(view.cells.map((cell) => cell.rawValue)).toEqual(['00123', 60, '#REF!', '', null, null])
  expect(view.cells[1]).toMatchObject({ kind: 'date', dateSystem: '1900', rawValue: 60 })
  expect(view.cells[2]!.kind).toBe('error')
  expect(view.cells[3]!.kind).toBe('string')
  expect(view.cells[4]!.kind).toBe('blank')
})
it('maps coordinates through structure and refuses unreadable sources instead of returning empty data', async () => {
  const current = structuredClone(state)
  current.structuralLog.push({ op: 'insert_rows', sheetId, row: 1, count: 2 })
  expect(
    (
      await readCurrentRange({
        state: current,
        sessionId: 'source',
        source,
        logicalSheetId: sheetId,
        range: 'A1:A3',
      })
    ).cells.map((cell) => cell.rawValue),
  ).toEqual([null, null, 10])
  const broken = {
    ...source,
    async readBaseRange() {
      throw new Error('Index unavailable')
    },
  }
  await expect(
    readCurrentRange({
      state,
      sessionId: 'source',
      source: broken,
      logicalSheetId: sheetId,
      range: 'A1',
    }),
  ).rejects.toThrow('Index unavailable')
  await expect(
    readCurrentRange({
      state,
      sessionId: 'source',
      source,
      logicalSheetId: sheetId,
      range: 'A1:C1000',
    }),
  ).rejects.toThrow('2000')
})
