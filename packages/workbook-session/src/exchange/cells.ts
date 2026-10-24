import type { CellRecord } from '../domain'
import { typedCellSchema, type TypedCell } from './schemas'

export function toTypedCell(cell: CellRecord): TypedCell {
  const numberFormat = cell.numberFormat ?? cell.style?.numberFormat
  const style = numberFormat === undefined ? {} : { numberFormat }
  if (cell.formula) {
    const cached =
      cell.cachedValue === undefined
        ? null
        : toTypedCell({
            ...cell,
            rawValue: cell.cachedValue,
            formula: undefined,
            calculationState: undefined,
            cachedValue: undefined,
          })
    return typedCellSchema.parse({
      kind: 'formula',
      formula: cell.formula,
      cached,
      calculationState: cell.calculationState ?? 'pending',
      ...style,
    })
  }
  if (cell.kind === 'date')
    return typedCellSchema.parse({
      kind: 'date',
      serial: cell.rawValue,
      dateSystem: cell.dateSystem,
      numberFormat: numberFormat ?? 'yyyy-mm-dd',
    })
  if (cell.kind === 'error')
    return typedCellSchema.parse({ kind: 'error', excelCode: cell.rawValue, ...style })
  return typedCellSchema.parse({ kind: cell.kind, value: cell.rawValue, ...style })
}

export function fromTypedCell(
  raw: TypedCell,
  options: { trustFormulaCache?: boolean } = {},
): CellRecord {
  const cell = typedCellSchema.parse(raw)
  const style =
    cell.numberFormat === undefined
      ? {}
      : { numberFormat: cell.numberFormat, style: { numberFormat: cell.numberFormat } }
  if (cell.kind === 'formula') {
    const cached = cell.cached
      ? fromTypedCell(cell.cached)
      : { kind: 'blank' as const, rawValue: null }
    return {
      ...cached,
      ...style,
      formula: cell.formula,
      ...(cell.cached ? { cachedValue: cached.rawValue } : {}),
      calculationState: options.trustFormulaCache ? cell.calculationState : 'pending',
    }
  }
  if (cell.kind === 'date')
    return { kind: 'date', rawValue: cell.serial, dateSystem: cell.dateSystem, ...style }
  if (cell.kind === 'error') return { kind: 'error', rawValue: cell.excelCode, ...style }
  return { kind: cell.kind, rawValue: cell.value, ...style }
}
