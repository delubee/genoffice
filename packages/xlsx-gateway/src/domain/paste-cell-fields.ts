import { isDateFormat } from 'numfmt'
import type { z } from 'zod'
import type { cellRecordSchema } from '../shared/cell-record'
export type ClipboardCell = z.infer<typeof cellRecordSchema>

/** Compose selective clipboard fields without rewriting unrelated target data. */
export function pasteCellFields(
  target: ClipboardCell,
  source: ClipboardCell,
  options: {
    writeValues?: boolean | undefined
    format?: 'all' | 'none' | 'number' | 'except-borders' | undefined
  },
): ClipboardCell {
  const mode = options.format ?? 'all'
  const value = structuredClone(options.writeValues === false ? target : source)
  delete value.style
  delete value.styleReset
  delete value.numberFormat
  const donor = mode === 'none' || mode === 'number' ? target : source
  const style = structuredClone(donor.style ?? {})
  let numberFormat = donor.numberFormat ?? donor.style?.numberFormat
  if (mode === 'number') {
    numberFormat = source.numberFormat ?? source.style?.numberFormat ?? 'General'
    style.numberFormat = numberFormat
  }
  if (mode === 'except-borders')
    for (const edge of ['borderTop', 'borderBottom', 'borderLeft', 'borderRight'] as const)
      style[edge] = structuredClone(target.style?.[edge] ?? null)
  if (numberFormat) {
    value.numberFormat = numberFormat
    style.numberFormat = numberFormat
  }
  value.style = style
  if (mode === 'all' || mode === 'except-borders' || target.styleReset) value.styleReset = true
  // Values-only dates are serial numbers unless the retained destination format
  // renders them as dates. Date systems remain explicit for actual date cells.
  if (
    (mode === 'none' || mode === 'number') &&
    typeof value.rawValue === 'number' &&
    !value.formula &&
    (value.kind === 'date' || value.kind === 'number')
  ) {
    value.kind = numberFormat && isDateFormat(numberFormat) ? 'date' : 'number'
    if (value.kind === 'date') value.dateSystem ??= target.dateSystem ?? source.dateSystem ?? '1900'
    else delete value.dateSystem
  }
  return value
}
