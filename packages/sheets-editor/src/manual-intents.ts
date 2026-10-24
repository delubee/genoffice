import { formatAddress, parseAddress } from '@genoffice/xlsx-gateway/domain/cell-address'
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import type { GridRange } from './ports'
import { workbookOperationSchema } from '@genoffice/xlsx-gateway/domain/workbook-dsl'

/** Scalar native style commands use the same format DSL as toolbar actions. */
export function manualStyleIntent(
  sheetId: string,
  range: GridRange,
  style: { type: string; value: unknown },
): WorkbookOperation {
  const { type, value } = style
  let format: Record<string, unknown>
  switch (type) {
    case 'bl':
      format = { bold: value === 1 }
      break
    case 'it':
      format = { italic: value === 1 }
      break
    case 'ff':
      format = { fontFamily: value }
      break
    case 'fs':
      format = { fontSize: value }
      break
    case 'cl':
    case 'bg': {
      const color =
        value === null
          ? null
          : typeof value === 'object' && value && 'rgb' in value
            ? value.rgb
            : undefined
      if (color === undefined) throw new Error('unsupported_manual_style')
      format = { [type === 'cl' ? 'fontColor' : 'fillColor']: color }
      break
    }
    case 'ht':
      format = {
        horizontalAlign: ({ 1: 'left', 2: 'center', 3: 'right' } as Record<string, string>)[
          String(value)
        ],
      }
      break
    case 'vt':
      format = {
        verticalAlign: ({ 1: 'top', 2: 'center', 3: 'bottom' } as Record<string, string>)[
          String(value)
        ],
      }
      break
    case 'tb':
      if (value !== 1 && value !== 3) throw new Error('unsupported_manual_style')
      format = { wrapText: value === 3 }
      break
    default:
      throw new Error('unsupported_manual_style')
  }
  if ((type === 'bl' || type === 'it') && value !== 0 && value !== 1)
    throw new Error('unsupported_manual_style')
  return workbookOperationSchema.parse({
    op: 'format_range',
    sheetId,
    range: `${formatAddress(range.startRow, range.startColumn)}:${formatAddress(range.endRow, range.endColumn)}`,
    format,
  })
}

/** The value command is intercepted before it can mutate the visible model. */
export function manualValueIntents(
  sheetId: string,
  input: { value: unknown; range?: GridRange },
  merges: readonly GridRange[] = [],
): WorkbookOperation[] {
  const result: WorkbookOperation[] = []
  function cell(row: number, column: number, raw: unknown) {
    if (result.length >= 2000) throw new Error('range_limit_exceeded')
    const address = formatAddress(row, column)
    if (raw === null) {
      result.push({ op: 'clear_cell', sheetId, address })
      return
    }
    if (!raw || typeof raw !== 'object') throw new Error('unsupported_cell_payload')
    const value = raw as {
      v?: unknown
      f?: unknown
      p?: { body?: { dataStream?: string } }
      s?: unknown
    }
    if (typeof value.f === 'string' && value.f.startsWith('=')) {
      result.push({ op: 'set_formula', sheetId, address, formula: value.f })
      return
    }
    const scalar = value.v ?? value.p?.body?.dataStream?.replace(/\r?\n$/, '') ?? null
    if (
      typeof scalar !== 'string' &&
      typeof scalar !== 'number' &&
      typeof scalar !== 'boolean' &&
      scalar !== null
    )
      throw new Error('unsupported_cell_value')
    if (typeof scalar === 'number' && !Number.isFinite(scalar))
      throw new Error('invalid_numeric_value')
    result.push({ op: 'set_cell', sheetId, address, value: scalar })
  }
  if (Array.isArray(input.value)) {
    const origin = input.range ?? { startRow: 0, startColumn: 0 }
    input.value.forEach((row, rowOffset) => {
      if (!Array.isArray(row)) throw new Error('invalid_cell_matrix')
      row.forEach((value, columnOffset) =>
        cell(origin.startRow + rowOffset, origin.startColumn + columnOffset, value),
      )
    })
  } else if (
    input.value &&
    typeof input.value === 'object' &&
    ['v', 'f', 'p'].some((key) => Object.hasOwn(input.value!, key))
  ) {
    if (!input.range) throw new Error('missing_edit_range')
    const { range } = input
    if ((range.endRow - range.startRow + 1) * (range.endColumn - range.startColumn + 1) > 2000)
      throw new Error('range_limit_exceeded')
    for (let row = range.startRow; row <= range.endRow; row++)
      for (let column = range.startColumn; column <= range.endColumn; column++)
        cell(row, column, input.value)
  } else {
    for (const [row, columns] of Object.entries((input.value as object) ?? {}))
      for (const [column, value] of Object.entries(columns as object))
        cell(Number(row), Number(column), value)
  }
  if (!result.length) throw new Error('empty_edit_payload')
  // Univer expands a scalar edit across the merged selection. Only its anchor
  // stores a value in Excel; forwarding the covered cells creates hidden data.
  return result.filter((op) => {
    if (!('address' in op)) return true
    const point = parseAddress(op.address)
    const merge = merges.find(
      (merge) =>
        point.row >= merge.startRow &&
        point.row <= merge.endRow &&
        point.column >= merge.startColumn &&
        point.column <= merge.endColumn &&
        (point.row !== merge.startRow || point.column !== merge.startColumn),
    )
    if (!merge) return true
    const anchor = result.find(
      (candidate) =>
        'address' in candidate &&
        candidate.address === formatAddress(merge.startRow, merge.startColumn),
    )
    if (
      !anchor ||
      JSON.stringify({ ...op, address: '' }) !== JSON.stringify({ ...anchor, address: '' })
    )
      throw new Error('merged_cell_target: edit the merged cell from its top-left anchor')
    return false
  })
}
