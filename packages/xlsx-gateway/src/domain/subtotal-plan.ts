import { columnLabel, type RangeBounds } from './cell-address'
import { MAX_GRID_ROWS } from '../shared/grid-bounds'

export interface SubtotalConfig {
  readonly groupCol: number
  readonly valueCol: number
  readonly agg: 'sum' | 'count' | 'average'
}

/** Desktop adjacent-group subtotals, in final coordinates after all inserts. */
export function buildSubtotalPlan(
  range: RangeBounds,
  labels: readonly unknown[],
  config: SubtotalConfig,
) {
  if (range.endRow <= range.startRow) throw new Error('appSelectRangeWithHeader')
  for (const column of [config.groupCol, config.valueCol])
    if (!Number.isInteger(column) || column < range.startColumn || column > range.endColumn)
      throw new Error('Invalid Subtotal column.')
  if (config.groupCol === config.valueCol) throw new Error('appSubtotalSameColumns')
  if (!['sum', 'count', 'average'].includes(config.agg))
    throw new Error('Invalid Subtotal function.')
  if (labels.length !== range.endRow - range.startRow)
    throw new Error('Incomplete Subtotal labels.')
  const groups: { label: string; start: number; end: number }[] = []
  let start = range.startRow + 1
  let label = String(labels[0] ?? '')
  for (let index = 1; index <= labels.length; index++) {
    const next = index < labels.length ? String(labels[index] ?? '') : null
    if (next !== label) {
      groups.push({ label, start, end: range.startRow + index })
      start = range.startRow + index + 1
      if (next !== null) label = next
    }
  }
  if (groups.length > 200) throw new Error('appSubtotalTooManyGroups')
  const grandRow = range.endRow + groups.length + 1
  if (grandRow >= MAX_GRID_ROWS) throw new Error('Subtotal exceeds the worksheet row limit.')
  const fn = config.agg === 'sum' ? 9 : config.agg === 'count' ? 3 : 1
  const column = columnLabel(config.valueCol)
  const totals = groups.map((group, shift) => ({
    row: group.end + shift + 1,
    label: `${group.label} Total`,
    formula: `=SUBTOTAL(${fn},$${column}$${group.start + shift + 1}:$${column}$${group.end + shift + 1})`,
    detailStart: group.start + shift,
    detailEnd: group.end + shift,
  }))
  return {
    groups,
    insertBefore: [...groups]
      .reverse()
      .map((group) => group.end + 1)
      .concat(grandRow),
    totals,
    grand: {
      row: grandRow,
      label: 'Grand Total',
      formula: `=SUBTOTAL(${fn},$${column}$${range.startRow + 2}:$${column}$${grandRow})`,
    },
  }
}
