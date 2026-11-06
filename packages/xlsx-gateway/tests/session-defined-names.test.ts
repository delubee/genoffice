import { expect, it } from 'vitest'
import { workbookStateSchema } from '../src/shared/session-state'
function initialState(input: { definedNames?: object[]; sheets: unknown[] }) {
  return workbookStateSchema.parse({
    schemaVersion: 1,
    stateId: 'names-root',
    structuralLog: [],
    features: { originalMetadata: { definedNames: input.definedNames ?? [] } },
    sheets: [{ id: 'logical', name: 'Data', rows: 1, columns: 1, cells: {}, features: {} }],
  })
}
import {
  editSessionDefinedName,
  manageSessionDefinedName,
  sessionDefinedNames,
} from '../src/domain/session-defined-names'
it('preserves unmodeled source names and case-ambiguous entries without modifying metadata', () => {
  const state = structuredClone(
    initialState({
      sheets: [{ id: 'base', name: 'Data', rowCount: 1, columnCount: 1 }],
      definedNames: [
        { name: 'Good', formula: 'Data!$A$1' },
        { name: 'Local', formula: '1', sheetIndex: 0 },
        { name: 'Invalid name', formula: '2' },
        { name: 'MissingScope', formula: '3', sheetIndex: 7 },
        { name: 'Duplicate', formula: '4' },
        { name: 'duplicate', formula: '5' },
      ],
    }),
  )
  const before = structuredClone(state)
  const current = sessionDefinedNames(state)
  expect(current.names).toEqual([
    { name: 'Good', formula: 'Data!$A$1' },
    { name: 'Local', formula: '1', scopeSheetId: state.sheets[0]!.id },
  ])
  expect(current.preserveNames).toEqual(['Invalid name', 'MissingScope', 'Duplicate', 'duplicate'])
  expect(() =>
    editSessionDefinedName(state, { op: 'add_defined_name', name: 'DUPLICATE', ref: '6' }),
  ).toThrow('duplicate_defined_name')
  expect(state).toEqual(before)
})
it('rejects reserved names and blank formulas and keeps global/local namespaces independent', () => {
  const state = structuredClone(
    initialState({ sheets: [{ id: 'base', name: 'Data', rowCount: 1, columnCount: 1 }] }),
  )
  for (const name of ['A1', 'R1C1', '_XLNM.Private']) {
    expect(() => editSessionDefinedName(state, { op: 'add_defined_name', name, ref: '1' })).toThrow(
      'invalid_defined_name',
    )
  }
  expect(() =>
    editSessionDefinedName(state, { op: 'add_defined_name', name: 'Blank', ref: '=' }),
  ).toThrow('empty_defined_name_reference')
  const first = editSessionDefinedName(state, { op: 'add_defined_name', name: 'Total', ref: '1' })
  state.features!.definedNamesEdits = JSON.parse(JSON.stringify(first))
  expect(
    editSessionDefinedName(state, {
      op: 'add_defined_name',
      name: 'total',
      ref: '2',
      scopeSheetId: state.sheets[0]!.id,
    }).names,
  ).toHaveLength(2)
  expect(() =>
    editSessionDefinedName(state, { op: 'delete_defined_name', name: 'missing' }),
  ).toThrow('defined_name_missing_or_ambiguous')
})

it('checks protected source namespaces before accepting names and preserves logical scope after sheet deletion', () => {
  const state = initialState({ sheets: [] })
  state.sheets[0]!.originalSheetId = 'original'
  state.features!.originalMetadata = {
    nameBindingsComplete: true,
    reservedDefinedNames: [
      { name: 'Hidden', sheetIndex: 0, sheetId: 'original' },
      { name: 'GlobalHidden' },
    ],
  }
  expect(() =>
    editSessionDefinedName(state, {
      op: 'add_defined_name',
      name: 'Hidden',
      ref: '1',
      scopeSheetId: state.sheets[0]!.id,
    }),
  ).toThrow('duplicate_defined_name')
  expect(
    editSessionDefinedName(state, { op: 'add_defined_name', name: 'Hidden', ref: '1' }).names,
  ).toHaveLength(1)
  expect(() =>
    editSessionDefinedName(state, { op: 'add_defined_name', name: 'GlobalHidden', ref: '1' }),
  ).toThrow('duplicate_defined_name')
})

it('rejects incomplete bindings and destination shadowing atomically, including hidden local scopes', () => {
  const state = initialState({ sheets: [], definedNames: [{ name: 'Revenue', formula: '1' }] })
  const operation = {
    op: 'manage_defined_name' as const,
    name: 'Revenue',
    replacement: { name: 'Sales', ref: '1' },
  }
  expect(() => manageSessionDefinedName(state, operation)).toThrow(
    'defined_name_bindings_incomplete',
  )
  state.sheets[0]!.originalSheetId = 'original'
  state.features!.originalMetadata = {
    definedNames: [{ name: 'Revenue', formula: '1' }],
    nameBindingsComplete: true,
    reservedDefinedNames: [{ name: 'Sales', sheetIndex: 0, sheetId: 'original' }],
  }
  const before = structuredClone(state)
  expect(() => manageSessionDefinedName(state, operation)).toThrow(
    'defined_name_destination_shadowed',
  )
  expect(state).toEqual(before)
})
