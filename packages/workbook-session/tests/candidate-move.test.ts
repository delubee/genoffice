import { expect, it } from 'vitest'
import { initialState } from '../src/state'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'

function fixture() {
  const state = structuredClone(
    initialState({
      sheets: [
        { id: 'data', name: 'Data', rowCount: 100, columnCount: 10 },
        { id: 'other', name: 'Other', rowCount: 100, columnCount: 10 },
      ],
    }),
  )
  const source = state.sheets[0]!,
    target = state.sheets[1]!
  source.cells = {
    A2: {
      kind: 'string',
      rawValue: '=literal',
      style: { bold: true, borderLeft: { style: 'double', color: '#123456' } },
    },
    B2: { kind: 'error', rawValue: '#DIV/0!' },
    A3: { kind: 'date', rawValue: 45000, dateSystem: '1900', numberFormat: 'yyyy-mm-dd' },
    B3: {
      kind: 'number',
      rawValue: 12,
      formula: '=A3+$C$1',
      calculationState: 'ready',
      cachedValue: 12,
    },
    C1: { kind: 'number', rawValue: 7 },
    J80: {
      kind: 'number',
      rawValue: 9,
      formula: '=$A$3+A3+SUM(A1:B3)',
      calculationState: 'ready',
      cachedValue: 9,
    },
  }
  target.cells = {
    A80: {
      kind: 'number',
      rawValue: 9,
      formula: '=Data!$A$3+Data!A2',
      calculationState: 'ready',
      cachedValue: 9,
    },
  }
  const move: WorkbookOperation = {
    op: 'relocate_range',
    sourceSheetId: source.id,
    sheetId: target.id,
    source: 'A2:B3',
    start: 'D5',
  }
  const run = (operation = move, complete = state.sheets.map((s) => s.id)) =>
    computeCandidateState(
      state,
      {
        schemaVersion: 1,
        sessionId: 'move',
        transactionId: crypto.randomUUID(),
        expectedRevision: 0,
        origin: 'manual',
        summary: 'Cut block',
        operations: [operation],
      },
      {
        state: structuredClone(state),
        completeSheetIds: complete,
        writableSheetIds: state.sheets.filter((s) => !s.protected).map((s) => s.id),
        allocation: { stateId: crypto.randomUUID(), sheetIds: {} },
      },
    )
  return { state, source, target, run, move }
}
it('moves typed cells and replacement formats across sheets while updating distant absolute and relative dependents', async () => {
  const f = fixture(),
    before = structuredClone(f.state),
    result = await f.run()
  const source = result.nextState.sheets[0]!.cells,
    target = result.nextState.sheets[1]!.cells
  expect(source.A2).toMatchObject({ rawValue: null, styleReset: true, style: {} })
  expect(source.B3!.formula).toBeUndefined()
  expect(target.D5).toMatchObject({
    kind: 'string',
    rawValue: '=literal',
    styleReset: true,
    style: { bold: true, borderLeft: { style: 'double', color: '#123456' } },
  })
  expect(target.E5).toMatchObject({ kind: 'error', rawValue: '#DIV/0!' })
  expect(target.D6).toMatchObject({
    kind: 'date',
    rawValue: 45000,
    dateSystem: '1900',
    numberFormat: 'yyyy-mm-dd',
  })
  expect(target.E6).toMatchObject({
    formula: "=D6+'Data'!$C$1",
    calculationState: 'pending',
    rawValue: null,
  })
  expect(target.E6!.cachedValue).toBeUndefined()
  expect(source.J80!.formula).toBe("='Other'!$D$6+'Other'!D6+SUM(A1:B3)")
  expect(target.A80!.formula).toBe("='Other'!$D$6+'Other'!D5")
  expect(result.inverse.state).toEqual(before)
  expect(f.state).toEqual(before)
})
it('moves overlapping blocks from an immutable source and retains formulas without copy offsets', async () => {
  const f = fixture()
  const result = await f.run({ ...f.move, sheetId: f.source.id, start: 'B3' })
  const cells = result.nextState.sheets[0]!.cells
  expect(cells.A2!.rawValue).toBeNull()
  expect(cells.B3!.rawValue).toBe('=literal')
  expect(cells.C3!.kind).toBe('error')
  expect(cells.B4!.rawValue).toBe(45000)
  expect(cells.C4!.formula).toBe('=B4+$C$1')
})
it('moves contained merges, notes and links and retains outside merges', async () => {
  const f = fixture()
  f.source.cells.B2 = { kind: 'blank', rawValue: null }
  f.source.features = {
    merges: [{ startRow: 1, endRow: 1, startColumn: 0, endColumn: 1 }],
    noteEdits: { A2: 'Note' },
    hyperlinkEdits: { A3: 'https://example.com' },
  }
  f.target.features = { merges: [{ startRow: 8, endRow: 8, startColumn: 0, endColumn: 1 }] }
  const result = await f.run()
  expect(result.nextState.sheets[0]!.features!.mergeEdits).toEqual([])
  expect(result.nextState.sheets[1]!.features!.mergeEdits).toEqual([
    { startRow: 8, endRow: 8, startColumn: 0, endColumn: 1 },
    { startRow: 4, endRow: 4, startColumn: 3, endColumn: 4 },
  ])
  expect(result.nextState.sheets[0]!.features!.noteEdits).toEqual({ A2: null })
  expect(result.nextState.sheets[1]!.features!.noteEdits).toEqual({ D5: 'Note' })
  expect(result.nextState.sheets[0]!.features!.hyperlinkEdits).toEqual({ A3: null })
  expect(result.nextState.sheets[1]!.features!.hyperlinkEdits).toEqual({
    D6: 'https://example.com',
  })
})
it('rejects incomplete dependency material and protected source, destination or changed dependent sheets atomically', async () => {
  const f = fixture(),
    before = structuredClone(f.state)
  await expect(f.run(f.move, [f.source.id])).rejects.toThrow('incomplete_move_material')
  expect(f.state).toEqual(before)
  f.source.protected = true
  await expect(f.run()).rejects.toThrow(/protected_target|unverified_writable_target/)
  f.source.protected = false
  f.target.protected = true
  await expect(f.run()).rejects.toThrow(/protected_target|unverified_writable_target/)
  f.target.protected = false
  const third = structuredClone(f.target)
  third.id = 'dependent'
  third.name = 'Dependent'
  third.protected = true
  third.cells = {
    A1: { kind: 'blank', rawValue: null, formula: '=Data!A2', calculationState: 'pending' },
  }
  f.state.sheets.push(third)
  await expect(f.run()).rejects.toThrow('protected_formula_dependency')
})

it('rejects partial merges and combined expanded-write overflow without changing either root', async () => {
  const f = fixture()
  f.source.features = { merges: [{ startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 }] }
  const before = structuredClone(f.state)
  await expect(f.run()).rejects.toThrow('merge_partial_overlap')
  expect(f.state).toEqual(before)
  f.source.features = {}
  await expect(f.run({ ...f.move, source: 'A1:A1001' })).rejects.toThrow('expanded_write_limit')
})
