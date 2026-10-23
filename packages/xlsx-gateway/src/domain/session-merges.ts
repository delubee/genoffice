import { formatAddress, parseRange } from './cell-address'
import { copyTargetBounds, type WorkbookOperation } from './workbook-dsl'
import type { CandidateState } from '../shared/session-changeset'
import { workbookMergeRangesResultSchema } from '../shared/workbook-protocol'

type Area = ReturnType<typeof parseRange>
const intersects = (a: Area, b: Area) =>
  a.startRow <= b.endRow &&
  a.endRow >= b.startRow &&
  a.startColumn <= b.endColumn &&
  a.endColumn >= b.startColumn
const contains = (a: Area, b: Area) =>
  a.startRow <= b.startRow &&
  a.endRow >= b.endRow &&
  a.startColumn <= b.startColumn &&
  a.endColumn >= b.endColumn

/** Both manual and agent candidates obey the same original merge geometry. */
export function assertMergeOperation(state: CandidateState, op: WorkbookOperation) {
  if (!('sheetId' in op)) return
  const mergesFor = (id: string) =>
    workbookMergeRangesResultSchema.shape.merges.parse(
      state.sheets.find((sheet) => sheet.id === id)?.features?.merges ?? [],
    )
  const merges = mergesFor(op.sheetId)
  let bounds: Area | undefined
  if ('address' in op) bounds = parseRange(op.address)
  if (op.op === 'set_range') {
    const start = parseRange(op.start ?? op.range ?? 'A1')
    bounds = {
      ...start,
      endRow: start.startRow + op.values.length - 1,
      endColumn: start.startColumn + op.values[0]!.length - 1,
    }
  }
  if (op.op === 'clear_range') bounds = parseRange(op.range)
  if (bounds)
    for (const merge of merges) {
      if (!intersects(bounds, merge)) continue
      const anchorOnly =
        bounds.startRow === merge.startRow &&
        bounds.endRow === merge.startRow &&
        bounds.startColumn === merge.startColumn &&
        bounds.endColumn === merge.startColumn
      if (anchorOnly || (op.op === 'clear_range' && contains(bounds, merge))) continue
      throw new Error(
        `merged_cell_target: edit ${formatAddress(merge.startRow, merge.startColumn)}, the top-left anchor of this merged cell`,
      )
    }
  if (op.op === 'copy_range' || op.op === 'fill_range') {
    const source = parseRange(op.source)
    const target = op.op === 'copy_range' ? copyTargetBounds(op) : parseRange(op.target)
    const sourceId = 'sourceSheetId' in op && op.sourceSheetId ? op.sourceSheetId : op.sheetId
    if (
      merges.some((merge) => intersects(target, merge)) ||
      mergesFor(sourceId).some((merge) => intersects(source, merge))
    )
      throw new Error('merged_range_copy: copying or filling merged ranges is not supported')
  }
}
