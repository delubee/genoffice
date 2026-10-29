import { expect, it } from 'vitest'
import { initialState } from '../src/state'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import { ribbonIntents, type RibbonTarget } from '@genoffice/sheets-editor/ribbon-intents'
import { toSelectionFormat } from '../../../apps/sheets/src/renderer/selection-format'

const target = (address: string): RibbonTarget => ({
  sheetId: 'sheet',
  address,
  format: toSelectionFormat({}, 'General'),
  doubleUnderline: false,
  dimensions: { rows: 10, columns: 10 },
})

async function apply(command: string, selection: RibbonTarget) {
  const state = structuredClone(
    initialState({
      sheets: [{ id: 'sheet', name: 'Data', rowCount: 10, columnCount: 10 }],
    }),
  )
  const sheetId = state.sheets[0]!.id
  selection.sheetId = sheetId
  state.sheets[0]!.cells = {
    B2: { kind: 'number', rawValue: 11, style: { indent: 0 } },
    C2: { kind: 'blank', rawValue: null, formula: '=$A2+B$1', style: { indent: 2 } },
    B3: { kind: 'number', rawValue: 22, style: { indent: 249 } },
    C3: { kind: 'number', rawValue: 33, style: { indent: 250 } },
    J10: { kind: 'string', rawValue: 'untouched' },
  }
  const result = await computeCandidateState(
    state,
    {
      schemaVersion: 1,
      sessionId: 'relative',
      transactionId: crypto.randomUUID(),
      expectedRevision: 0,
      origin: 'manual',
      summary: command,
      operations: ribbonIntents(command, selection),
    },
    {
      state: structuredClone(state),
      completeSheetIds: [sheetId],
      writableSheetIds: [sheetId],
      allocation: { stateId: crypto.randomUUID(), sheetIds: {} },
    },
  )
  expect(result.inverse.state).toEqual(state)
  expect(result.nextState.sheets[0]!.cells.J10).toEqual(state.sheets[0]!.cells.J10)
  return result.nextState.sheets[0]!.cells
}

it('steps mixed cell indents independently and clamps at both desktop limits', async () => {
  const selection = target('B2:C3')
  selection.cellIndents = [
    { address: 'B2', indent: 0 },
    { address: 'C2', indent: 2 },
    { address: 'B3', indent: 249 },
    { address: 'C3', indent: 250 },
  ]
  const increased = await apply('indent-step:1', selection)
  expect(['B2', 'C2', 'B3', 'C3'].map((a) => increased[a]?.style?.indent)).toEqual([1, 3, 250, 250])
  const decreased = await apply('indent-step:-1', selection)
  expect(['B2', 'C2', 'B3', 'C3'].map((a) => decreased[a]?.style?.indent)).toEqual([0, 1, 248, 249])
  expect(decreased.C2?.formula).toBe('=$A2+B$1')
})

it.each([
  ['fill-down', 'B3:C3', 'B2:C2', 'B3:C3'],
  ['fill-up', 'B1:C1', 'B2:C2', 'B1:C1'],
  ['fill-right', 'D2:D3', 'C2:C3', 'D2:D3'],
  ['fill-left', 'A2:A3', 'B2:B3', 'A2:A3'],
  ['fill-down', 'B2:C4', 'B2:C2', 'B3:C4'],
  ['fill-up', 'B1:C3', 'B3:C3', 'B1:C2'],
  ['fill-right', 'B2:D3', 'B2:B3', 'C2:D3'],
  ['fill-left', 'A2:C3', 'C2:C3', 'A2:B3'],
])('matches desktop %s source and destination for %s', (command, address, source, destination) => {
  expect(ribbonIntents(command!, target(address!))).toMatchObject([
    { op: 'fill_range', source, target: destination },
  ])
})

it('copies single-row/column formulas in every direction without overwriting the donor', async () => {
  const down = await apply('fill-down', target('B3:C3'))
  expect(down.B3?.rawValue).toBe(11)
  expect(down.C3?.formula).toBe('=$A3+B$1')
  expect(down.C2?.formula).toBe('=$A2+B$1')
  const up = await apply('fill-up', target('B1:C1'))
  expect(up.C1?.formula).toBe('=$A1+B$1')
  const right = await apply('fill-right', target('D2:D3'))
  expect(right.D2?.formula).toBe('=$A2+C$1')
  const left = await apply('fill-left', target('A2:A3'))
  expect(left.A2?.rawValue).toBe(11)
})

it('does nothing at sheet edges and rejects incomplete/foreign indent snapshots before submitting', () => {
  for (const [command, address] of [
    ['fill-down', 'A1:B1'],
    ['fill-right', 'A1:A2'],
    ['fill-up', 'I10:J10'],
    ['fill-left', 'J9:J10'],
  ])
    expect(ribbonIntents(command!, target(address!))).toEqual([])
  const selection = target('B2:C2')
  expect(() => ribbonIntents('indent-step:1', selection)).toThrow('snapshot')
  selection.cellIndents = [{ address: 'B2', indent: 0 }]
  expect(() => ribbonIntents('indent-step:1', selection)).toThrow()
  selection.cellIndents = [
    { address: 'B2', indent: 0 },
    { address: 'B2', indent: 1 },
  ]
  expect(() => ribbonIntents('indent-step:1', selection)).toThrow()
  selection.cellIndents = [
    { address: 'B2', indent: 0 },
    { address: 'D2', indent: 1 },
  ]
  expect(() => ribbonIntents('indent-step:1', selection)).toThrow()
  selection.cellIndents = [
    { address: 'B2', indent: 0 },
    { address: 'C2', indent: 1 },
  ]
  expect(() => ribbonIntents('indent-step:2', selection)).toThrow()
})

it('steps General and explicit number formats as one validated batch without changing values/formulas', async () => {
  const selection = target('B2:C3')
  selection.decimalCells = [
    { row: 1, col: 1, pattern: 'General', raw: { v: 11, t: 2 } },
    { row: 1, col: 2, pattern: '0.00%', raw: { v: 0.12, t: 2 } },
    { row: 2, col: 1, raw: { v: 22, t: 2 } },
    { row: 2, col: 2, pattern: '@', raw: { v: 'text', t: 1 } },
  ]
  const up = await apply('decimal-inc', selection)
  expect(up.B2?.style?.numberFormat).toBe('0.000')
  expect(up.C2?.style?.numberFormat).toBe('0.000%')
  expect(up.B3?.style?.numberFormat).toBe('0.000')
  expect(up.C3?.style?.numberFormat).toBeUndefined()
  expect(up.B2?.rawValue).toBe(11)
  expect(up.C2?.formula).toBe('=$A2+B$1')
  const down = await apply('decimal-dec', selection)
  expect(down.B2?.style?.numberFormat).toBe('0.0')
  expect(down.C2?.style?.numberFormat).toBe('0.0%')
  expect(down.C3?.style?.numberFormat).toBe('@')
})

it('handles explicit General decimals and rejects missing/duplicate number format snapshots', () => {
  const selection = target('A1')
  selection.decimalCells = [{ row: 0, col: 0, pattern: 'General', raw: { v: 341504.77, t: 2 } }]
  expect(ribbonIntents('decimal-inc', selection)).toMatchObject([
    { range: 'A1', format: { numberFormat: '0.000' } },
  ])
  expect(ribbonIntents('decimal-dec', selection)).toMatchObject([
    { range: 'A1', format: { numberFormat: '0.0' } },
  ])
  delete selection.decimalCells
  expect(() => ribbonIntents('decimal-inc', selection)).toThrow('snapshot')
  selection.address = 'A1:B1'
  selection.decimalCells = [
    { row: 0, col: 0 },
    { row: 0, col: 0 },
  ]
  expect(() => ribbonIntents('decimal-inc', selection)).toThrow('snapshot')
})
