import { expect, it } from 'vitest'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import { initialState } from '../src/state'
it('rejects adding a worksheet to a read-only workbook before allocating a candidate', async () => {
  const state = initialState({
    readOnly: true,
    sheets: [{ id: 'base', name: 'Data', rowCount: 1, columnCount: 1 }],
  })
  await expect(
    computeCandidateState(
      state,
      {
        schemaVersion: 1,
        sessionId: 'readonly',
        transactionId: 'readonly-add',
        expectedRevision: 0,
        origin: 'agent',
        summary: 'Must reject new sheet',
        operations: [{ op: 'add_sheet', name: 'Bypass' }],
      },
      {
        state,
        completeSheetIds: [],
        writableSheetIds: [],
        allocation: { stateId: 'changed', sheetIds: { '0': 'new' } },
      },
    ),
  ).rejects.toThrow('readonly_workbook')
  expect(state.sheets).toHaveLength(1)
})
it('does not persist hydrated original cells as document edits', async () => {
  const state = initialState({
    sheets: [{ id: 'base', name: 'Data', rowCount: 10, columnCount: 3 }],
  })
  const material = structuredClone(state),
    sheetId = state.sheets[0]!.id
  material.sheets[0]!.cells = {
    A1: { kind: 'number', rawValue: 1 },
    B1: { kind: 'string', rawValue: '00123', style: { bold: true } },
    C1: { kind: 'error', rawValue: '#REF!' },
  }
  const result = await computeCandidateState(
    state,
    {
      schemaVersion: 1,
      sessionId: 'session',
      transactionId: 'overlay-test',
      expectedRevision: 0,
      origin: 'agent',
      summary: 'Only A1 changes',
      operations: [{ op: 'set_cell', sheetId, address: 'A1', value: 9 }],
    },
    {
      state: material,
      completeSheetIds: [sheetId],
      writableSheetIds: [sheetId],
      allocation: { stateId: 'after', sheetIds: {} },
    },
  )
  expect(Object.keys(result.nextState.sheets[0]!.cells)).toEqual(['A1'])
  expect(result.nextState.sheets[0]!.cells.A1!.rawValue).toBe(9)
  expect(result.inverse.state).toEqual(state)
})
