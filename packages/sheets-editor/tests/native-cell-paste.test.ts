import { CellValueType, BorderStyleTypes, type IMutationInfo } from '@univerjs/core'
import { expect, it } from 'vitest'
import { nativePasteOperations } from '../src/native-cell-paste'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import { workbookStateSchema } from '@genoffice/xlsx-gateway/shared/session-state'

it('preserves boolean types and replaces original formatting through one reversible canonical paste', async () => {
  const state = workbookStateSchema.parse({
    schemaVersion: 1,
    stateId: 'before',
    structuralLog: [],
    sheets: [
      {
        id: 'sheet',
        name: 'Data',
        rows: 10,
        columns: 10,
        cells: {
          A1: { kind: 'string', rawValue: 'Before', style: { italic: true, fillColor: '#FF0000' } },
        },
      },
    ],
  })
  const operations = nativePasteOperations(
    [
      {
        id: 'sheet.mutation.set-range-values',
        params: { subUnitId: 'sheet', cellValue: { 0: { 0: { s: null } } } },
      },
      {
        id: 'sheet.mutation.set-range-values',
        params: { subUnitId: 'sheet', cellValue: { 0: { 0: { v: 1, t: CellValueType.BOOLEAN } } } },
      },
      {
        id: 'sheet.mutation.set-range-values',
        params: { subUnitId: 'sheet', cellValue: { 0: { 0: { s: { bl: 1, it: 0, bg: null } } } } },
      },
    ],
    'sheet',
  )
  const result = await computeCandidateState(
    state,
    {
      schemaVersion: 1,
      sessionId: 'paste',
      transactionId: 'paste',
      expectedRevision: 0,
      origin: 'manual',
      summary: 'Paste',
      operations,
    },
    {
      state: structuredClone(state),
      writableSheetIds: ['sheet'],
      completeSheetIds: ['sheet'],
      allocation: { stateId: 'after', sheetIds: {} },
    },
  )
  expect(result.nextState.sheets[0]!.cells.A1).toMatchObject({
    kind: 'boolean',
    rawValue: true,
    styleReset: true,
    style: { bold: true, italic: false, fillColor: null },
  })
  expect(result.inverse.state).toEqual(state)
})

it('rejects unsupported clipboard mutations, foreign sheets and mixed rich text before publishing a batch', () => {
  const value: IMutationInfo = {
    id: 'sheet.mutation.set-range-values',
    params: { subUnitId: 'sheet', cellValue: { 0: { 0: { v: 'Prefix' } } } },
  }
  expect(() =>
    nativePasteOperations(
      [value, { id: 'sheet.mutation.delete-worksheet', params: { subUnitId: 'sheet' } }],
      'sheet',
    ),
  ).toThrow(/clipboard|剪贴板/)
  expect(() => nativePasteOperations([value], 'foreign')).toThrow('invalid_manual_target')
  expect(() =>
    nativePasteOperations(
      [
        {
          id: value.id,
          params: {
            subUnitId: 'sheet',
            cellValue: {
              0: {
                0: {
                  p: {
                    body: {
                      dataStream: 'Mixed text\r\n',
                      textRuns: [{ st: 0, ed: 5, ts: { bl: 1 } }],
                    },
                  },
                },
              },
            },
          },
        },
      ],
      'sheet',
    ),
  ).toThrow(/clipboard|剪贴板/)
})

it('keeps independent clipboard border edges and clears source borders without resurrecting an old preset', async () => {
  const state = workbookStateSchema.parse({
    schemaVersion: 1,
    stateId: 'border-before',
    structuralLog: [],
    dateSystem: '1900',
    sheets: [
      {
        id: 'sheet',
        name: 'Data',
        rows: 10,
        columns: 10,
        cells: {
          A1: {
            kind: 'string',
            rawValue: 'Before',
            style: {
              borderTop: { style: 'thick', color: '#FF0000' },
              borderBottom: { style: 'double', color: '#000000' },
            },
          },
        },
      },
    ],
  })
  const operations = nativePasteOperations(
    [
      {
        id: 'sheet.mutation.set-range-values',
        params: {
          subUnitId: 'sheet',
          cellValue: {
            0: {
              0: {
                v: 'Pasted',
                s: {
                  bd: {
                    t: null,
                    b: { s: BorderStyleTypes.DOTTED, cl: { rgb: '#123456' } },
                    l: { s: 11, cl: { rgb: 'rgb(1, 2, 3)' } },
                  },
                },
              },
            },
          },
        },
      },
    ],
    'sheet',
  )
  const result = await computeCandidateState(
    state,
    {
      schemaVersion: 1,
      sessionId: 'paste',
      transactionId: 'borders',
      expectedRevision: 0,
      origin: 'manual',
      summary: 'Paste borders',
      operations,
    },
    {
      state: structuredClone(state),
      writableSheetIds: ['sheet'],
      completeSheetIds: ['sheet'],
      allocation: { stateId: 'after', sheetIds: {} },
    },
  )
  expect(result.nextState.sheets[0]!.cells.A1!.style).toMatchObject({
    borderTop: null,
    borderBottom: { style: 'dotted', color: '#123456' },
    borderLeft: { style: 'mediumDashDotDot', color: '#010203' },
  })
  expect(result.inverse.state).toEqual(state)
})
