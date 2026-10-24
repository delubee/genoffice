import { formatAddress, parseAddress, parseRange, rangeCellCount } from './cell-address'
import { copyTargetBounds, type WorkbookOperation } from './workbook-dsl'
import { workbookMergeRangesResultSchema } from '../shared/workbook-protocol'

type Area = ReturnType<typeof parseRange>
export function sessionMerges(features?: Record<string, unknown>) {
  return workbookMergeRangesResultSchema.shape.merges.parse(
    features?.mergeEdits ?? features?.merges ?? [],
  )
}

/** Never discard covered data: callers must explicitly clear it before merging. */
export function updateSessionMerges(
  features: Record<string, unknown>,
  op: Extract<WorkbookOperation, { op: 'merge_cells' | 'unmerge_cells' }>,
  cells: Record<string, { rawValue: unknown; formula?: string }>,
) {
  const range = parseRange(op.range),
    merges = sessionMerges(features)
  if (rangeCellCount(range) > 200_000) throw new Error('range_limit')
  if (op.op === 'unmerge_cells') return merges.filter((merge) => !intersects(range, merge))
  if (rangeCellCount(range) === 1) throw new Error('merge_select_range')
  const metadata = features.originalMetadata as
    { tables?: { range: Area }[]; pivotRanges?: Area[] } | undefined
  const filter = features.filterEdits as { filter?: { range: Area } | null } | undefined
  const filteredArea = filter ? filter.filter?.range : (features.autoFilter as Area | undefined)
  if (
    metadata?.tables?.some((table) => intersects(range, table.range)) ||
    metadata?.pivotRanges?.some((pivot) => intersects(range, pivot)) ||
    (filteredArea && intersects(range, filteredArea))
  )
    throw new Error('merged_range_operation: select a range without tables, pivots or filters')
  if (merges.some((merge) => intersects(range, merge) && !contains(range, merge)))
    throw new Error('merge_partial_overlap')
  for (const [address, cell] of Object.entries(cells)) {
    const point = parseAddress(address)
    if (
      !contains(range, {
        startRow: point.row,
        endRow: point.row,
        startColumn: point.column,
        endColumn: point.column,
      })
    )
      continue
    if (point.row === range.startRow && point.column === range.startColumn) continue
    if (cell.formula || cell.rawValue !== null) throw new Error('merge_nonempty_cells')
  }
  return [...merges.filter((merge) => !intersects(range, merge)), range]
}
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

/** Manual, agent and import candidates obey the same current merge geometry. */
export function assertMergeOperation(
  state: { sheets: readonly { id: string; features?: Record<string, unknown> | undefined }[] },
  op: WorkbookOperation,
) {
  if (!('sheetId' in op)) return
  const mergesFor = (id: string) =>
    sessionMerges(state.sheets.find((sheet) => sheet.id === id)?.features)
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
  if (
    (op.op === 'sort_range' || op.op === 'find_replace' || op.op === 'set_filter') &&
    merges.some((merge) => intersects(parseRange(op.range), merge))
  )
    throw new Error('merged_range_operation: select a range without merged cells')
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
