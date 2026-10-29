import { expect, it } from 'vitest'
import {
  rewriteNameReferences,
  type NameReferencePlan,
} from '../src/domain/defined-name-references'
const sheets = [
  { id: 'data', name: 'Data' },
  { id: 'other', name: "Bob's Data" },
]
it('rewrites complete identifiers and qualified names while preserving literals, tables, errors and external books', () => {
  const plan: NameReferencePlan = {
    sheets,
    changes: [{ name: 'Revenue', replacement: 'Sales', shadowedSheetIds: [] }],
  }
  expect(
    rewriteNameReferences("=Revenue+REVENUE+Data!Revenue+'Bob''s Data'!Revenue", 'data', plan),
  ).toBe("=Sales+Sales+Data!Sales+'Bob''s Data'!Sales")
  for (const formula of [
    '="Revenue"',
    '="a ""Revenue"""',
    '=Revenue2',
    '=Table1[Revenue]',
    '=Revenue[Column]',
    "='[Other.xlsx]Data'!Revenue",
    '=[Book.xlsx]Data!Revenue',
    '=#REF!+#N/A',
  ])
    expect(rewriteNameReferences(formula, 'data', plan)).toBe(formula)
})
it('captures scoped shadowing at action time and preserves later name reuse through ordered changes', () => {
  const global: NameReferencePlan = {
    sheets,
    changes: [{ name: 'Revenue', replacement: 'Sales', shadowedSheetIds: ['other'] }],
  }
  expect(rewriteNameReferences('=Revenue+Data!Revenue', 'other', global)).toBe(
    '=Revenue+Data!Sales',
  )
  const local: NameReferencePlan = {
    sheets,
    changes: [{ name: 'Revenue', replacement: null, scopeSheetId: 'other', shadowedSheetIds: [] }],
  }
  expect(rewriteNameReferences("=Revenue+'Bob''s Data'!Revenue", 'data', local)).toBe(
    '=Revenue+#REF!',
  )
  expect(rewriteNameReferences('=Revenue', 'other', local)).toBe('=#REF!')
  const ordered: NameReferencePlan = {
    sheets,
    changes: [
      ...global.changes,
      { name: 'Revenue', replacement: 'NewSource', shadowedSheetIds: [] },
    ],
  }
  expect(rewriteNameReferences('=Revenue', 'data', ordered)).toBe('=Sales')
  expect(rewriteNameReferences('=Revenue', 'other', ordered)).toBe('=NewSource')
})
