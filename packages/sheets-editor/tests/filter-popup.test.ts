import { expect, it } from 'vitest'
import { filterColumnToUniver, filterPopupOperation } from '../src/filter-popup'

const range = { startRow: 10, endRow: 20, startColumn: 3, endColumn: 5 }
it('restores imported column offsets, blanks and default equal conditions without recalculating', () => {
  expect(filterColumnToUniver(range, { colId: 1, values: ['00123'], blank: true })).toEqual({
    colId: 4,
    filters: { filters: ['00123'], blank: true },
  })
  expect(
    filterColumnToUniver(range, {
      colId: 0,
      customs: {
        and: true,
        filters: [
          { val: 'N*', operator: 'equal' },
          { val: '*West', operator: 'notEqual' },
        ],
      },
    }),
  ).toEqual({
    colId: 3,
    customFilters: {
      and: 1,
      customFilters: [{ val: 'N*' }, { val: '*West', operator: 'notEqual' }],
    },
  })
  expect(filterColumnToUniver(range, { colId: 3, values: ['outside'] })).toBeNull()
})
it('turns native popup values, blank-only, clear and custom criteria into canonical transactions', () => {
  expect(
    filterPopupOperation('s', range, 4, { colId: 4, filters: { filters: ['00123'], blank: true } }),
  ).toEqual({ op: 'set_filter_criteria', sheetId: 's', column: 'E', values: ['00123', ''] })
  expect(filterPopupOperation('s', range, 3, { colId: 3, filters: { blank: true } })).toMatchObject(
    { column: 'D', values: [''] },
  )
  expect(filterPopupOperation('s', range, 5, null)).toMatchObject({ column: 'F', values: null })
  expect(
    filterPopupOperation('s', range, 4, {
      colId: 4,
      customFilters: {
        and: 1,
        customFilters: [
          { val: 10, operator: 'greaterThan' },
          { val: 20, operator: 'lessThan' },
        ],
      },
    }),
  ).toMatchObject({
    column: 'E',
    customs: {
      and: true,
      filters: [
        { val: 10, operator: 'greaterThan' },
        { val: 20, operator: 'lessThan' },
      ],
    },
  })
})
it('rejects unsupported or malformed drafts instead of mutating the local view', () => {
  expect(() => filterPopupOperation('s', range, 2, null)).toThrow('outside_range')
  expect(() =>
    filterPopupOperation('s', range, 3, { colId: 3, filters: { filters: [] } }),
  ).toThrow()
})

it('round-trips native fill and font color criteria at offset columns', () => {
  expect(
    filterPopupOperation('s', range, 4, {
      colId: 4,
      colorFilters: { cellFillColors: ['rgb(255,0,0)', null] },
    }),
  ).toEqual({
    op: 'set_filter_criteria',
    sheetId: 's',
    column: 'E',
    colors: { fill: ['#FF0000', null] },
  })
  expect(filterColumnToUniver(range, { colId: 1, colors: { font: ['#000000'] } })).toEqual({
    colId: 4,
    colorFilters: { cellTextColors: ['rgb(0,0,0)'] },
  })
})
