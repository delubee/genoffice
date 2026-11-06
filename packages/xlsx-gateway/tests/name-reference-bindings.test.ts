import { expect, it } from 'vitest'
import { rewriteNameReferences } from '../src/domain/defined-name-references'
import {
  rewriteSourceDefinedNameReferences,
  decodeDefinedNameFormula,
} from '../src/gateway/xlsx-defined-names'
import { manageSessionDefinedName } from '../src/domain/session-defined-names'
import { workbookStateSchema } from '../src/shared/session-state'
const plan = {
  changes: [{ name: 'Revenue', replacement: 'Sales', shadowedSheetIds: ['other'] }],
  sheets: [
    { id: 'data', name: 'Data' },
    { id: 'other', name: 'Other' },
  ],
}

it('rewrites document names while preserving LET/LAMBDA declarations, initializers and nested scope', () => {
  for (const [before, after] of [
    ['=LET(Revenue,Revenue+1,Revenue)+Revenue', '=LET(Revenue,Sales+1,Revenue)+Sales'],
    ['=LAMBDA(Revenue,Revenue+OtherName)(Revenue)', '=LAMBDA(Revenue,Revenue+OtherName)(Sales)'],
    ['=LET(Revenue,{1,2},SUM(Revenue))+Revenue', '=LET(Revenue,{1,2},SUM(Revenue))+Sales'],
    ['=LET(OtherName,1,Revenue+OtherName)', '=LET(OtherName,1,Sales+OtherName)'],
    [
      '=LAMBDA(Revenue,LET(Revenue,Revenue+1,Revenue))',
      '=LAMBDA(Revenue,LET(Revenue,Revenue+1,Revenue))',
    ],
    ['=LET(Revenue,1,Data!Revenue+Revenue)', '=LET(Revenue,1,Data!Sales+Revenue)'],
    ['="Revenue"&Table[Revenue]&Revenue', '="Revenue"&Table[Revenue]&Sales'],
  ])
    expect(rewriteNameReferences(before!, 'data', plan)).toBe(after)
})

it('keeps built-in calls while renaming bare names and callable named LAMBDAs', () => {
  const change = (name: string) => ({
    ...plan,
    changes: [{ name, replacement: 'NewName', shadowedSheetIds: [] }],
  })
  expect(rewriteNameReferences('=SUM(1,2)+SUM', 'data', change('SUM'))).toBe('=SUM(1,2)+NewName')
  expect(rewriteNameReferences('=Bump(1)+Other!Bump(2)', 'data', change('Bump'))).toBe(
    '=NewName(1)+Other!NewName(2)',
  )
  expect(rewriteNameReferences('=LAMBDA(Bump,Bump(1))(Bump)', 'data', change('Bump'))).toBe(
    '=LAMBDA(Bump,Bump(1))(NewName)',
  )
})

it('updates source definition dependencies with global lexical scope and preserves unchanged hidden XML and literal CDATA entities', () => {
  const hidden = '<definedName name="Unchanged" hidden="1"><![CDATA["Revenue"]]></definedName>'
  const xml =
    '<definedNames><definedName name="Global">Revenue*2</definedName><definedName name="Local" localSheetId="1">Revenue+Data!Revenue</definedName><definedName name="Secret" hidden="1"><![CDATA[IF(Revenue>0,"&amp; Revenue",0)]]></definedName>' +
    hidden +
    '</definedNames>'
  const result = rewriteSourceDefinedNameReferences(xml, plan)
  expect(result).toContain('<definedName name="Global">Sales*2</definedName>')
  expect(result).toContain(
    '<definedName name="Local" localSheetId="1">Revenue+Data!Sales</definedName>',
  )
  expect(result).toContain('IF(Sales&gt;0,"&amp;amp; Revenue",0)')
  expect(result).toContain(hidden)
  expect(decodeDefinedNameFormula('&quot;&#26679;&#x672c;&amp;&quot;')).toBe('"样本&"')
  expect(() =>
    rewriteSourceDefinedNameReferences(
      '<definedName name="Bad" localSheetId="99">Revenue</definedName>',
      plan,
    ),
  ).toThrow('Unknown defined-name reference scope')
})

it('rewrites current modeled dependencies at action time and leaves a newly supplied replacement reference untouched', () => {
  const state = workbookStateSchema.parse({
    schemaVersion: 1,
    stateId: 'names',
    sheets: plan.sheets.map((sheet) => ({ ...sheet, rows: 1, columns: 1, cells: {} })),
    structuralLog: [],
    features: {
      originalMetadata: { nameBindingsComplete: true },
      definedNamesEdits: {
        preserveNames: [],
        names: [
          { name: 'Revenue', formula: '10' },
          { name: 'Revenue', formula: '20', scopeSheetId: 'other' },
          { name: 'Total', formula: 'Revenue*2' },
          { name: 'Total', formula: 'Revenue+Data!Revenue', scopeSheetId: 'other' },
          { name: 'Lambda', formula: 'LAMBDA(Revenue,Revenue+1)' },
        ],
      },
    },
  })
  const result = manageSessionDefinedName(state, {
    name: 'Revenue',
    replacement: { name: 'Sales', ref: 'Revenue+1' },
  })
  expect(result.names.names.find((name) => name.name === 'Sales')!.formula).toBe('Revenue+1')
  expect(
    result.names.names.find((name) => name.name === 'Total' && !name.scopeSheetId)!.formula,
  ).toBe('Sales*2')
  expect(
    result.names.names.find((name) => name.name === 'Total' && name.scopeSheetId)!.formula,
  ).toBe('Revenue+Data!Sales')
  expect(result.names.names.find((name) => name.name === 'Lambda')!.formula).toBe(
    'LAMBDA(Revenue,Revenue+1)',
  )
})
