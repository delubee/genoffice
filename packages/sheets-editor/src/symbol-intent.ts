import { covertCellValue, type ICellData } from '@univerjs/core'
import { appendSymbol } from '../../../apps/sheets/src/renderer/SymbolDialog'
import {
  workbookOperationSchema,
  type WorkbookOperation,
} from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { parseAddress } from '@genoffice/xlsx-gateway/domain/cell-address'
import { manualValueIntents } from './manual-intents'
import type { ViewCell } from './grid-cell'

/** Same append + facade conversion as desktop, applied only to the selected anchor. */
export function symbolIntents(
  sheetId: string,
  address: string,
  cell: ViewCell,
  symbol: string,
): WorkbookOperation[] {
  if (cell?.formula && cell.calculationState !== 'ready' && cell.calculationState !== 'error')
    throw new Error('The formula result is pending. No symbols were inserted.')
  const value = cell?.formula
    ? Object.hasOwn(cell, 'cachedValue')
      ? cell.cachedValue
      : cell.rawValue
    : cell?.rawValue
  const text = appendSymbol(value, symbol)
  const converted = covertCellValue(text) as ICellData
  if (typeof converted.v === 'string' && converted.v.length > 32_767)
    throw new Error('The symbol would exceed the cell text limit. No cells were changed.')
  const position = parseAddress(address)
  const operations = manualValueIntents(sheetId, {
    value: [[converted]],
    range: {
      startRow: position.row,
      endRow: position.row,
      startColumn: position.column,
      endColumn: position.column,
    },
  })
  if (converted.s && typeof converted.s !== 'string' && converted.s.n?.pattern)
    operations.push({
      op: 'format_range',
      sheetId,
      range: address,
      format: { numberFormat: converted.s.n.pattern },
    })
  return operations.map((operation) => workbookOperationSchema.parse(operation))
}
