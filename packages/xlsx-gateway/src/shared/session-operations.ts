export const SESSION_OPERATION_NAMES = [
  'set_cell',
  'set_formula',
  'clear_cell',
  'set_range',
  'clear_range',
  'format_range',
  'merge_cells',
  'unmerge_cells',
  'set_row_height',
  'set_col_width',
  'set_page_setup',
  'set_freeze',
  'set_note',
  'sort_range',
  'find_replace',
  'set_filter',
  'clear_filter',
  'set_filter_criteria',
  'copy_range',
  'fill_range',
  'insert_rows',
  'delete_rows',
  'insert_cols',
  'delete_cols',
  'add_sheet',
  'delete_sheet',
  'rename_sheet',
  'add_chart',
  'edit_chart',
  'delete_visual',
] as const

const structural = new Set([
  'insert_rows',
  'delete_rows',
  'insert_cols',
  'delete_cols',
  'add_sheet',
  'delete_sheet',
  'rename_sheet',
])
export function hasMixedSessionOperations(operations: readonly { op: string }[]) {
  return (
    operations.some((operation) => structural.has(operation.op)) &&
    operations.some((operation) => !structural.has(operation.op))
  )
}
