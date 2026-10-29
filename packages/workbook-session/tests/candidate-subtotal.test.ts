import { expect, it } from 'vitest'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import { initialState } from '../src/state'

function fixture() {
  const state = structuredClone(
    initialState({
      sheets: [
        { id: 'data', name: 'Data', rowCount: 8, columnCount: 3 },
        { id: 'other', name: 'Other', rowCount: 1, columnCount: 1 },
      ],
    }),
  )
  const sheet = state.sheets[0]!
  sheet.cells = {
    B3: { kind: 'string', rawValue: 'Group' },
    C3: { kind: 'string', rawValue: 'Amount' },
    B4: { kind: 'string', rawValue: 'North' },
    C4: { kind: 'number', rawValue: 12.5, style: { fillColor: '#CCEEFF' } },
    B5: { kind: 'string', rawValue: 'North' },
    C5: { kind: 'number', rawValue: 7 },
    B6: { kind: 'string', rawValue: 'South' },
    C6: { kind: 'number', rawValue: 3 },
    A8: { kind: 'number', rawValue: 42 },
  }
  state.sheets[1]!.cells.A1 = {
    kind: 'number',
    rawValue: 42,
    formula: '=Data!A8',
    cachedValue: 42,
    calculationState: 'ready',
  }
  const intent = {
    schemaVersion: 1,
    sessionId: 'subtotal',
    transactionId: 'test',
    expectedRevision: 0,
    origin: 'manual',
    summary: 'Subtotal',
    operations: [
      {
        op: 'subtotal',
        sheetId: sheet.id,
        range: 'B3:C6',
        groupColumn: 'B',
        valueColumn: 'C',
        agg: 'sum',
      },
    ],
  }
  const material = {
    state: structuredClone(state),
    completeSheetIds: state.sheets.map((s) => s.id),
    writableSheetIds: state.sheets.map((s) => s.id),
    allocation: { stateId: 'subtotal-result', sheetIds: {} },
  }
  return { state, intent, material }
}
it('builds offset subtotals, grand total and outline in one changeset and shifts unrelated formulas exactly once', async () => {
  const f = fixture(),
    before = structuredClone(f.state)
  const result = await computeCandidateState(f.state, f.intent, f.material)
  expect(f.state).toEqual(before)
  expect(result.inverse.state).toEqual(before)
  const [sheet, other] = result.nextState.sheets
  expect(result.beforeStateId).toBe(before.stateId)
  expect(result.nextState.stateId).toBe('subtotal-result')
  expect(sheet!.cells.C6!.formula).toBe('=SUBTOTAL(9,$C$4:$C$5)')
  expect(sheet!.cells.C8!.formula).toBe('=SUBTOTAL(9,$C$7:$C$7)')
  expect(sheet!.cells.C9!.formula).toBe('=SUBTOTAL(9,$C$4:$C$8)')
  expect(sheet!.cells.B6!.rawValue).toBe('North Total')
  expect(sheet!.cells.B8!.rawValue).toBe('South Total')
  expect(sheet!.cells.B9!.rawValue).toBe('Grand Total')
  expect(sheet!.cells.C4!.style?.fillColor).toBe('#CCEEFF')
  expect(sheet!.cells.A11!.rawValue).toBe(42)
  expect(other!.cells.A1!.formula).toBe('=Data!A11')
  expect(sheet!.features?.outlineEdits).toMatchObject([
    { op: 'set_outline_settings', summaryBelow: true },
    { op: 'set_outline', start: 4, count: 5, level: 1 },
    { op: 'set_outline', start: 4, count: 2, level: 2 },
    { op: 'set_outline', start: 7, count: 1, level: 2 },
  ])
})
it('rejects unsupported structure, missing material and invalid columns without publishing any state', async () => {
  for (const variant of ['incomplete', 'same-column', 'source-outline', 'protected'] as const) {
    const f = fixture()
    if (variant === 'incomplete') f.material.completeSheetIds = []
    if (variant === 'same-column') f.intent.operations[0]!.valueColumn = 'B'
    if (variant === 'source-outline')
      f.material.state.sheets[0]!.features!.originalMetadata = { outlineLevelRow: 1 }
    if (variant === 'protected') f.material.writableSheetIds = []
    const before = structuredClone(f.state)
    await expect(computeCandidateState(f.state, f.intent, f.material)).rejects.toThrow()
    expect(f.state).toEqual(before)
  }
})

it.each([
  ['sum', 9],
  ['count', 3],
  ['average', 1],
] as const)('uses desktop %s SUBTOTAL numbers', async (agg, number) => {
  const f = fixture()
  f.intent.operations[0]!.agg = agg
  const result = await computeCandidateState(f.state, f.intent, f.material)
  expect(result.nextState.sheets[0]!.cells.C9!.formula).toBe(`=SUBTOTAL(${number},$C$4:$C$8)`)
})
it('never mutates the original root when a private structural or content stage fails or is cancelled', async () => {
  for (const failAt of [1, 4, 7]) {
    const f = fixture(),
      before = structuredClone(f.state)
    let steps = 0
    await expect(
      computeCandidateState(f.state, f.intent, f.material, {
        fault() {
          if (++steps === failAt) throw new Error('private stage failure')
        },
      }),
    ).rejects.toThrow('private stage failure')
    expect(f.state).toEqual(before)
    expect(f.material.state).toEqual(before)
  }
  const f = fixture(),
    controller = new AbortController()
  controller.abort()
  await expect(
    computeCandidateState(f.state, f.intent, f.material, { signal: controller.signal }),
  ).rejects.toThrow('operation_cancelled')
})
