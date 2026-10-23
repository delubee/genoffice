import { expect, it } from 'vitest'
import { assertSerializable, stateChangeset, type CandidateState } from '../src/changeset'
const state: CandidateState = {
  schemaVersion: 1,
  stateId: 'before',
  dateSystem: '1900',
  structuralLog: [],
  sheetMap: {},
  sheets: [{ id: 'logical', name: 'Data', rows: 10, columns: 2, cells: {} }],
}
it('rejects closures, runtime handles, accessors and cyclic history records', () => {
  for (const value of [
    { undo() {} },
    { handle: new Map() },
    { value: Infinity },
    {
      get secret() {
        throw new Error('getter ran')
      },
    },
  ])
    expect(() => assertSerializable(value)).toThrow()
  const cycle: Record<string, unknown> = {}
  cycle.self = cycle
  expect(() => assertSerializable(cycle)).toThrow('cyclic')
})
it('keeps complete prior and next feature states in JSON history', () => {
  const after = structuredClone(state)
  after.stateId = 'after'
  after.sheets[0]!.features = { charts: [{ id: 'chart-1', values: [1, 2] }] }
  const changes = JSON.parse(JSON.stringify(stateChangeset(state, after, ['logical', 'logical'])))
  expect(changes.inverse.state).toEqual(state)
  expect(changes.forward.state).toEqual(after)
  expect(changes.affectedTargets).toEqual(['logical'])
})
