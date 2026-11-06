import { columnLabel } from '@genoffice/xlsx-gateway/domain/cell-address'

export function definedNameFromLabel(label: string): string | null {
  const cleaned = label.trim().replace(/[^\p{L}\p{N}_.]/gu, '_')
  if (!cleaned || /^\.+$/.test(cleaned)) return null
  const named = /^[\p{L}_]/u.test(cleaned) ? cleaned : `_${cleaned}`
  const cellLike = /^[A-Za-z]{1,3}\d+$/.test(named) || /^[Rr]\d*([Cc]\d*)?$/.test(named)
  return cellLike ? `_${named}` : named
}

/// Absolute A1 ref over one column, rows in 0-based coordinates.
export function a1RangeRef(
  sheetName: string,
  column: number,
  fromRow: number,
  toRow: number,
): string {
  const col = columnLabel(column)
  const name = sheetName.replace(/'/g, "''")
  return `'${name}'!$${col}$${fromRow + 1}:$${col}$${toRow + 1}`
}

/// Absolute A1 ref over one row, columns in 0-based coordinates.
export function a1RowRangeRef(
  sheetName: string,
  row: number,
  fromColumn: number,
  toColumn: number,
): string {
  const name = sheetName.replace(/'/g, "''")
  return `'${name}'!$${columnLabel(fromColumn)}$${row + 1}:$${columnLabel(toColumn)}$${row + 1}`
}
