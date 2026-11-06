import { expect, it } from 'vitest'
import { customFilterPredicate } from '../src/domain/custom-filter'
import { workbookOperationSchema } from '../src/domain/workbook-dsl'

it('validates exactly one legacy value-list or one/two custom conditions', () => {
  const base = { op: 'set_filter_criteria', sheetId: 's', column: 'B' }
  expect(workbookOperationSchema.safeParse({ ...base, values: null }).success).toBe(true)
  expect(workbookOperationSchema.safeParse({ ...base, values: ['East'] }).success).toBe(true)
  expect(
    workbookOperationSchema.safeParse({
      ...base,
      customs: { filters: [{ val: '10', operator: 'greaterThan' }] },
    }).success,
  ).toBe(true)
  for (const patch of [
    {},
    { values: [], customs: { filters: [{ val: 'a' }] } },
    { values: null, customs: { filters: [{ val: 'a' }] } },
    { customs: { filters: [] } },
    { customs: { filters: [{ val: 'a' }, { val: 'b' }, { val: 'c' }] } },
    { customs: { filters: [{ val: Infinity }] } },
    { customs: { filters: [{ val: 'a', operator: 'contains' }] } },
  ])
    expect(workbookOperationSchema.safeParse({ ...base, ...patch }).success).toBe(false)
})

it('uses raw numeric bounds with AND/OR while preserving text-number distinctions', () => {
  const filters = [
    { operator: 'greaterThanOrEqual', val: '10' },
    { operator: 'lessThan', val: '20' },
  ]
  const between = customFilterPredicate({ and: true, filters })
  expect([null, 9, 10, 12.5, 20, '12.5', false].map(between)).toEqual([
    false,
    false,
    true,
    true,
    false,
    false,
    false,
  ])
  const outside = customFilterPredicate({
    and: false,
    filters: [
      { operator: 'lessThanOrEqual', val: 10 },
      { operator: 'greaterThan', val: 20 },
    ],
  })
  expect([10, 15, 20, 21].map(outside)).toEqual([true, false, false, true])
  const unequal = customFilterPredicate({ filters: [{ operator: 'notEqual', val: '10' }] })
  expect([10, 11, '10', null].map(unequal)).toEqual([false, true, true, true])
})

it('matches the desktop wildcard, boolean, blank and literal not-equal rules', () => {
  const starts = customFilterPredicate({ filters: [{ val: 'Ea*', operator: 'equal' }] })
  expect(['East', 'east', 'West', '', null].map(starts)).toEqual([true, true, false, false, false])
  expect(customFilterPredicate({ filters: [{ val: '*' }] })(null)).toBe(false)
  expect(customFilterPredicate({ filters: [{ val: '' }] })(null)).toBe(true)
  expect(customFilterPredicate({ filters: [{ val: 'TRUE' }] })(true)).toBe(true)
  expect(customFilterPredicate({ filters: [{ val: 'A[1].?' }] })('a[1].x')).toBe(true)
  const notStarts = customFilterPredicate({ filters: [{ val: 'Ea*', operator: 'notEqual' }] })
  expect(['East', 'east', 'West', null].map(notStarts)).toEqual([false, false, true, true])
  const literal = customFilterPredicate({ filters: [{ val: 'East', operator: 'notEqual' }] })
  expect(['East', 'east', null].map(literal)).toEqual([false, true, true])
})
