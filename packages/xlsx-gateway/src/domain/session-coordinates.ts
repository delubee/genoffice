import { columnIndex } from './cell-address'
import type { WorkbookOperation } from './workbook-dsl'

/** Resolve one current axis index into the immutable original workbook. */
export function originalAxisIndex(
  log: readonly WorkbookOperation[],
  sheetId: string,
  index: number,
  axis: 'row' | 'column',
): number | null {
  for (const op of [...log].reverse()) {
    if (!('sheetId' in op) || op.sheetId !== sheetId) continue
    if (
      op.op !== 'insert_rows' &&
      op.op !== 'delete_rows' &&
      op.op !== 'insert_cols' &&
      op.op !== 'delete_cols'
    )
      continue
    if (
      axis === 'row'
        ? op.op !== 'insert_rows' && op.op !== 'delete_rows'
        : op.op !== 'insert_cols' && op.op !== 'delete_cols'
    )
      continue
    const at =
      op.op === 'insert_rows' || op.op === 'delete_rows' ? op.row - 1 : columnIndex(op.column)
    const count = 'count' in op ? op.count : 0
    if (op.op.startsWith('insert_')) {
      if (index >= at && index < at + count) return null
      if (index >= at + count) index -= count
    } else if (index >= at) index += count
  }
  return index
}
