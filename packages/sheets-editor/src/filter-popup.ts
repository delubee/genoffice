import {
  fromUniverColorFilter,
  toUniverColorFilter,
} from '@genoffice/xlsx-gateway/domain/univer-color-filter'
import { BooleanNumber } from '@univerjs/core'
import type { IFilterColumn } from '@univerjs/preset-sheets-filter'
import { workbookOperationSchema } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { columnLabel } from '@genoffice/xlsx-gateway/domain/cell-address'
import type { WorkbookRangeResult } from '@genoffice/xlsx-gateway/shared/workbook-protocol'
import type { GridRange } from './ports'

/** OOXML offsets become absolute model columns; restoring never recalculates. */
export function filterColumnToUniver(
  range: GridRange,
  column: WorkbookRangeResult['autoFilterColumns'][number],
): IFilterColumn | null {
  const colId = range.startColumn + column.colId
  if (colId > range.endColumn) return null
  const criteria: IFilterColumn = { colId }
  if (column.values !== undefined || column.blank)
    criteria.filters = {
      ...(column.values === undefined ? {} : { filters: [...column.values] }),
      ...(column.blank ? { blank: true } : {}),
    }
  if (column.customs) {
    const filters = column.customs.filters.map(({ val, operator }) => ({
      val,
      ...(operator === undefined || operator === 'equal' ? {} : { operator }),
    }))
    criteria.customFilters = {
      ...(filters.length > 1 && column.customs.and ? { and: BooleanNumber.TRUE } : {}),
      customFilters: filters as NonNullable<IFilterColumn['customFilters']>['customFilters'],
    }
  }
  if (column.colors) criteria.colorFilters = toUniverColorFilter(column.colors)
  return criteria.filters || criteria.customFilters || criteria.colorFilters ? criteria : null
}

/** Validate the UI command at the canonical boundary, before any view mutation. */
export function filterPopupOperation(
  sheetId: string,
  range: GridRange,
  column: number,
  criteria: IFilterColumn | null | undefined,
) {
  if (!Number.isInteger(column) || column < range.startColumn || column > range.endColumn)
    throw new Error('filter_column_outside_range')
  const target = { op: 'set_filter_criteria', sheetId, column: columnLabel(column) }
  if (!criteria) return workbookOperationSchema.parse({ ...target, values: null })
  if (criteria.colorFilters) {
    if (criteria.filters || criteria.customFilters)
      throw new Error('Invalid mixed filter criteria.')
    return workbookOperationSchema.parse({
      ...target,
      colors: fromUniverColorFilter(criteria.colorFilters),
    })
  }
  if (criteria.customFilters && criteria.filters) throw new Error('Invalid mixed filter criteria.')
  if (criteria.customFilters)
    return workbookOperationSchema.parse({
      ...target,
      customs: {
        and: criteria.customFilters.and === BooleanNumber.TRUE,
        filters: criteria.customFilters.customFilters,
      },
    })
  if (criteria.filters)
    return workbookOperationSchema.parse({
      ...target,
      values: [...(criteria.filters.filters ?? []), ...(criteria.filters.blank ? [''] : [])],
    })
  throw new Error('Unsupported filter criteria.')
}
