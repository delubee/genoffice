import type { WorkbookOperation } from './workbook-dsl'

const viewFields = new Set(['showFormulas', 'showGridlines', 'showHeadings', 'zoomScale'])
/** Sheet protection restricts content, not these desktop worksheet-view controls.
 * A mixed page/print patch must retain normal write permission checks. */
export function isWorksheetViewOperation(operation: WorkbookOperation): boolean {
  return (
    operation.op === 'set_page_setup' &&
    Object.keys(operation).every(
      (key) => key === 'op' || key === 'sheetId' || viewFields.has(key),
    ) &&
    Object.keys(operation).some((key) => viewFields.has(key))
  )
}
