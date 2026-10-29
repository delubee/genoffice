import { expect, it } from 'vitest'
import { initialState } from '../src/state'
import { composeCalculationDefinitions } from '../src/calculation/definitions'
import { sessionDefinedNames } from '@genoffice/xlsx-gateway/domain/session-defined-names'
import type { CalculationContext } from '../src/calculation/context'
const stateOf = () =>
  structuredClone(
    initialState({
      sheets: [
        { id: 'base', name: 'Data', rowCount: 3, columnCount: 2 },
        { id: 'second', name: 'Other', rowCount: 3, columnCount: 2 },
      ],
      definedNames: [{ name: 'Local', formula: 'Other!$A$1', sheetIndex: 1 }],
    }),
  )
const contextOf = (): CalculationContext => ({
  sessionId: crypto.randomUUID(),
  names: [],
  calculationProperties: {},
  externalRelationshipIds: [],
})

it('keeps source/local scope stable when tabs move or source scopes are removed', () => {
  const state = stateOf(),
    context = contextOf(),
    other = state.sheets[1]!.id
  context.names = [{ name: 'Local', formula: 'Other!$A$1', sheetId: 'second', hidden: false }]
  state.sheets.reverse()
  expect(sessionDefinedNames(state).names[0]!.scopeSheetId).toBe(other)
  expect(composeCalculationDefinitions(state, context)).toEqual([
    { name: 'Local', formula: 'Other!$A$1', sheetId: other, hidden: false },
  ])
  state.sheets = state.sheets.filter((sheet) => sheet.id !== other)
  expect(sessionDefinedNames(state).names).toEqual([])
  expect(composeCalculationDefinitions(state, context)).toEqual([])
})

it('composes private source dependencies with action-time current names without replaying history over later definitions', () => {
  const state = stateOf(),
    context = contextOf(),
    other = state.sheets[1]!.id
  context.names = [
    { name: 'Revenue', formula: 'Data!$A$1', hidden: false },
    { name: 'Private', formula: 'Revenue*2', hidden: true },
    { name: 'Revenue', formula: 'Other!$A$1', hidden: true, sheetId: 'second' },
    { name: 'PrivateLocal', formula: 'Revenue+Data!Revenue', hidden: true, sheetId: 'second' },
    { name: 'Empty', formula: '', hidden: true },
    { name: '_xlnm.Print_Area', formula: 'Data!$A$1:$B$2', hidden: false, sheetId: 'base' },
    { name: 'Preserved', formula: 'Revenue+1', hidden: false },
  ]
  state.features!.definedNameReferences = [
    { name: 'Revenue', replacement: 'Net', shadowedSheetIds: [other] },
  ]
  state.features!.definedNamesEdits = {
    names: [
      { name: 'Net', formula: 'Data!$A$1' },
      { name: 'Revenue', formula: '1' },
      { name: 'Later', formula: 'Revenue*5' },
    ],
    preserveNames: ['Preserved'],
  }
  const before = structuredClone({ state, context })
  const result = composeCalculationDefinitions(state, context)
  expect(result.find((name) => name.name === 'Private')!.formula).toBe('Net*2')
  expect(result.find((name) => name.name === 'PrivateLocal')!.formula).toBe('Revenue+Data!Net')
  expect(result.find((name) => name.name === 'Preserved')!.formula).toBe('Net+1')
  expect(result.find((name) => name.name === 'Later')!.formula).toBe('Revenue*5')
  expect(result.filter((name) => name.name === 'Revenue' && !name.sheetId)).toEqual([
    { name: 'Revenue', formula: '1' },
  ])
  expect(result.find((name) => name.name === 'Empty')!.formula).toBe('')
  expect(result.find((name) => name.name === '_xlnm.Print_Area')!.sheetId).toBe(state.sheets[0]!.id)
  expect({ state, context }).toEqual(before)
})

it('maps source reference geometry while leaving a later journal definition in its already current coordinates', () => {
  const state = stateOf(),
    context = contextOf(),
    sheet = state.sheets[0]!
  context.names = [{ name: 'Source', formula: 'Data!$A$2', hidden: true }]
  state.structuralLog = [
    { op: 'insert_rows', sheetId: sheet.id, row: 1, count: 2 },
    { op: 'rename_sheet', sheetId: sheet.id, name: 'Renamed' },
  ]
  sheet.name = 'Renamed'
  state.features!.definedNamesEdits = {
    names: [{ name: 'Later', formula: 'Renamed!$A$2' }],
    preserveNames: [],
  }
  expect(composeCalculationDefinitions(state, context).map((name) => name.formula)).toEqual([
    'Renamed!$A$4',
    'Renamed!$A$2',
  ])
})
