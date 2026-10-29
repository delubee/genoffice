import { CellValueType, isDefaultFormat } from '@univerjs/core'
import { getDecimalFromPattern, setPatternDecimal } from '@univerjs/sheets-numfmt'

export interface DecimalCell {
  row: number
  col: number
  pattern?: string | null | undefined
  raw?: { v?: unknown; t?: unknown } | null
}

/** Decimal digits of the raw value's shortest string form (0.77 -> 2). */
export function rawValueDecimals(cell: DecimalCell['raw']): number {
  if (!cell || cell.t !== CellValueType.NUMBER || !cell.v) return 0
  const match = /\.(\d*)$/.exec(String(cell.v))
  return match?.[1]?.length ?? 0
}

/** Shared desktop/embedded decimal stepping, including General-format handling. */
export function decimalStepValues(cells: readonly DecimalCell[], step: 1 | -1) {
  const explicit = (cell: DecimalCell) =>
    cell.pattern && !isDefaultFormat(cell.pattern) ? cell.pattern : null
  let maxDecimals = 0
  for (const cell of cells) {
    const pattern = explicit(cell)
    if (pattern) maxDecimals = Math.max(maxDecimals, getDecimalFromPattern(pattern))
    // Preserve the desktop/Univer rule: raw values only seed a zero count.
    else if (!maxDecimals) maxDecimals = Math.max(maxDecimals, rawValueDecimals(cell.raw))
  }
  const decimals = maxDecimals + step
  const defaultPattern =
    step > 0
      ? setPatternDecimal(`0${decimals > 0 ? '.0' : ''}`, decimals)
      : setPatternDecimal(`0${decimals > 0 ? '.0' : '.'}`, decimals)
  return cells.flatMap((cell) => {
    const pattern = explicit(cell)
    const next = pattern
      ? setPatternDecimal(pattern, getDecimalFromPattern(pattern) + step)
      : defaultPattern
    return !pattern || step < 0 || next !== pattern
      ? [{ row: cell.row, col: cell.col, pattern: next }]
      : []
  })
}
