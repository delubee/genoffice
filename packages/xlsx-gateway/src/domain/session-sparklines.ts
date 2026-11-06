import { columnLabel, parseAddress, parseRange } from './cell-address'

export const MAX_SESSION_SPARKLINES = 200

export interface SessionSparklineGroup {
  readonly type: 'line' | 'column' | 'stacked'
  readonly color?: string | undefined
  readonly dataRange: string
  readonly targetCell?: string | undefined
}

/** Host cells and sheet-qualified source refs of a sparkline group, one per data row. */
export function sparklineCells(sheetName: string, group: SessionSparklineGroup) {
  const bounds = parseRange(group.dataRange)
  const rows = bounds.endRow - bounds.startRow + 1
  const base =
    group.targetCell === undefined
      ? { row: bounds.startRow, column: bounds.endColumn + 1 }
      : parseAddress(group.targetCell)
  const quoted = `'${sheetName.replace(/'/g, "''")}'`
  return Array.from({ length: rows }, (_, offset) => ({
    cell: `${columnLabel(base.column)}${base.row + offset + 1}`,
    sourceRef:
      `${quoted}!$${columnLabel(bounds.startColumn)}$${bounds.startRow + offset + 1}` +
      `:$${columnLabel(bounds.endColumn)}$${bounds.startRow + offset + 1}`,
  }))
}
