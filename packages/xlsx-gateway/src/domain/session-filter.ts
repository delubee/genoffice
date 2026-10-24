import { format as formatNumber, isDateFormat } from 'numfmt'
import { columnIndex, formatAddress, parseRange, rangeCellCount } from './cell-address'
import type { WorkbookOperation } from './workbook-dsl'
import type { WorkbookSnapshot } from './workbook.types'
import { workbookFilterStateSchema, workbookRangeResultSchema } from '../shared/workbook-protocol'

export function updateSessionFilter(
  features: Record<string, unknown>,
  op: Extract<WorkbookOperation, { op: 'set_filter' | 'clear_filter' | 'set_filter_criteria' }>,
  book: WorkbookSnapshot,
  dateSystem: '1900' | '1904' = '1900',
) {
  const source = features.filterEdits
    ? workbookFilterStateSchema.parse(features.filterEdits)
    : {
        sheetId: op.sheetId,
        filter: features.autoFilter
          ? {
              range: workbookRangeResultSchema.shape.autoFilter.unwrap().parse(features.autoFilter),
              columns: workbookRangeResultSchema.shape.autoFilterColumns.parse(
                features.autoFilterColumns ?? [],
              ),
            }
          : null,
        hiddenRows: [],
        visibilityRange: features.autoFilter
          ? workbookRangeResultSchema.shape.autoFilter.unwrap().parse(features.autoFilter)
          : undefined,
      }
  let filter = source.filter,
    visibilityRange = source.visibilityRange
  if (op.op === 'set_filter') {
    const range = parseRange(op.range)
    filter = { range, columns: [] }
    visibilityRange = visibilityRange
      ? {
          startRow: Math.min(visibilityRange.startRow, range.startRow),
          endRow: Math.max(visibilityRange.endRow, range.endRow),
          startColumn: Math.min(visibilityRange.startColumn, range.startColumn),
          endColumn: Math.max(visibilityRange.endColumn, range.endColumn),
        }
      : range
  } else if (op.op === 'clear_filter') filter = null
  else {
    if (!filter) throw new Error('filter_missing: create a filter on the selected table first')
    const colId = columnIndex(op.column) - filter.range.startColumn
    if (colId < 0 || colId > filter.range.endColumn - filter.range.startColumn)
      throw new Error('filter_column_outside_range')
    filter = {
      ...filter,
      columns: [
        ...filter.columns.filter((column) => column.colId !== colId),
        ...(op.values === null
          ? []
          : [
              {
                colId,
                values: op.values.filter((value) => value !== ''),
                blank: op.values.includes(''),
              },
            ]),
      ],
    }
  }
  if (!visibilityRange) throw new Error('filter_missing')
  if (rangeCellCount(visibilityRange) > 2000)
    throw new Error('filter_range_limit: filter at most 2000 cells per table in this preview')
  if (filter?.columns.some((column) => column.customs))
    throw new Error(
      'unsupported_filter_criteria: custom file filters must be cleared before editing',
    )
  const sheet = book.sheets.find((sheet) => sheet.id === op.sheetId)!,
    hiddenRows: number[] = []
  if (filter)
    for (let row = filter.range.startRow + 1; row <= filter.range.endRow; row++) {
      const matches = filter.columns.every((column) => {
        const address = formatAddress(row, filter!.range.startColumn + column.colId),
          value = sheet.cells[address]?.value ?? null
        const numberFormat = sheet.styles?.[address]?.numberFormat
        const label =
          typeof value === 'number' && numberFormat
            ? formatNumber(
                numberFormat,
                value + (dateSystem === '1904' && isDateFormat(numberFormat) ? 1462 : 0),
                { throws: false },
              )
            : typeof value === 'boolean'
              ? String(value).toUpperCase()
              : String(value ?? '')
        return (
          (column.blank && (value === null || value === '')) ||
          column.values?.some((choice) => choice.toLowerCase() === label.toLowerCase())
        )
      })
      if (!matches) hiddenRows.push(row)
    }
  return workbookFilterStateSchema.parse({
    sheetId: op.sheetId,
    filter,
    visibilityRange,
    hiddenRows,
  })
}
