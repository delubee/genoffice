export const SESSION_OPERATION_NAMES = [
  'set_cell',
  'set_formula',
  'clear_cell',
  'set_range',
  'clear_range',
  'format_range',
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
