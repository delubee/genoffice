import { expect, it } from 'vitest'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { initialState } from '../src/state'

async function copy(operations: (sheetId: string) => WorkbookOperation[]) {
  const before = initialState({
    sheets: [{ id: 'base', name: 'Data', rowCount: 20, columnCount: 5 }],
  })
  const material = structuredClone(before),
    sheetId = before.sheets[0]!.id
  material.sheets[0]!.cells = {
    A1: {
      kind: 'string',
      rawValue: '00123',
      style: { bold: true, borderTop: { style: 'double' } },
    },
    A2: {
      kind: 'date',
      rawValue: 60,
      dateSystem: '1900',
      numberFormat: 'yyyy-mm-dd',
      style: { numberFormat: 'yyyy-mm-dd' },
    },
    A3: { kind: 'error', rawValue: '#REF!' },
    A4: { kind: 'string', rawValue: '' },
    A5: { kind: 'number', rawValue: 4, formula: '=B1+$C$1', calculationState: 'ready' },
    C1: {
      kind: 'number',
      rawValue: 999,
      style: { italic: true, borderBottom: { style: 'double' } },
    },
    C6: { kind: 'number', rawValue: 999 },
  }
  const result = await computeCandidateState(
    before,
    {
      schemaVersion: 1,
      sessionId: 'session',
      transactionId: 'copy',
      expectedRevision: 0,
      origin: 'manual',
      summary: 'Copy current source',
      operations: operations(sheetId),
    },
    {
      state: material,
      completeSheetIds: [sheetId],
      writableSheetIds: [sheetId],
      allocation: { stateId: 'after', sheetIds: {} },
    },
  )
  return { result, before }
}

it('copies typed values, shifted formulas and blanks while retaining destination formatting', async () => {
  const { result, before } = await copy((sheetId) => [
    { op: 'copy_range', sheetId, source: 'A1:A6', target: 'C1' },
  ])
  const cells = result.nextState.sheets[0]!.cells
  expect(cells.C1).toMatchObject({
    kind: 'string',
    rawValue: '00123',
    style: { italic: true, borderBottom: { style: 'double' } },
  })
  expect(cells.C1!.style?.bold).toBeUndefined()
  expect(cells.C2).toMatchObject({ kind: 'date', rawValue: 60, dateSystem: '1900' })
  expect(cells.C3).toMatchObject({ kind: 'error', rawValue: '#REF!' })
  expect(cells.C4).toMatchObject({ kind: 'string', rawValue: '' })
  expect(cells.C5).toMatchObject({ formula: '=D1+$C$1', calculationState: 'pending' })
  expect(cells.C6).toMatchObject({ kind: 'blank', rawValue: null })
  expect(cells.A1).toBeUndefined()
  expect(result.inverse.state).toEqual(before)
})

it('tiles a range larger than the ordinary expanded cell budget and guards overlaps', async () => {
  const { result } = await copy((sheetId) => [
    { op: 'fill_range', sheetId, source: 'A1', target: 'B1:B2001' },
  ])
  expect(result.nextState.sheets[0]!.cells.B2001).toMatchObject({
    kind: 'string',
    rawValue: '00123',
  })
  expect(result.nextState.sheets[0]!.rows).toBe(2001)
  await expect(
    copy((sheetId) => [{ op: 'copy_range', sheetId, source: 'A1:A3', target: 'A2' }]),
  ).rejects.toThrow('overlap')
})
