import { expect, it } from 'vitest'
import { buildSubtotalPlan } from '../src/domain/subtotal-plan'
import { parseRange } from '../src/domain/cell-address'

it('groups adjacent equal text only, retaining blank and repeated later groups at absolute offsets', () => {
  const plan = buildSubtotalPlan(parseRange('C7:D12'), ['North', 'North', null, 'North', 'South'], {
    groupCol: 2,
    valueCol: 3,
    agg: 'count',
  })
  expect(plan.totals.map((total) => total.label)).toEqual([
    'North Total',
    ' Total',
    'North Total',
    'South Total',
  ])
  expect(plan.insertBefore).toEqual([12, 11, 10, 9, 16])
  expect(plan.totals.map((total) => total.formula)).toEqual([
    '=SUBTOTAL(3,$D$8:$D$9)',
    '=SUBTOTAL(3,$D$11:$D$11)',
    '=SUBTOTAL(3,$D$13:$D$13)',
    '=SUBTOTAL(3,$D$15:$D$15)',
  ])
  expect(plan.grand.formula).toBe('=SUBTOTAL(3,$D$8:$D$16)')
})
it('refuses mismatched labels, invalid columns, too many groups and grid overflow before generating writes', () => {
  const range = parseRange('A1:B3'),
    config = { groupCol: 0, valueCol: 1, agg: 'sum' as const }
  expect(() => buildSubtotalPlan(range, ['North'], config)).toThrow('Incomplete')
  expect(() => buildSubtotalPlan(range, ['North', 'South'], { ...config, valueCol: 0 })).toThrow(
    'appSubtotalSameColumns',
  )
  expect(() =>
    buildSubtotalPlan(
      parseRange('A1:B202'),
      Array.from({ length: 201 }, (_, index) => String(index)),
      config,
    ),
  ).toThrow('appSubtotalTooManyGroups')
  expect(() => buildSubtotalPlan(parseRange('A1048575:B1048576'), ['North'], config)).toThrow(
    'row limit',
  )
})
