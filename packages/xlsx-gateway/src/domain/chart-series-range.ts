import {
  columnLabel,
  formatAddress,
  parseRange,
  rangeAddresses,
  rangeCellCount,
} from './cell-address'
import type { ChartStateEdit } from './chart-visual'
import type { WorkbookSnapshot } from './workbook.types'

/** Resolve bounded chart vectors without accepting names, external books or unions. */
export function resolveChartRange(
  input: string,
  sheets: readonly { id: string; name: string }[],
  defaultSheetId: string,
) {
  const text = input.trim()
  const qualified = /^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$/.exec(text)
  const name = qualified?.[1]?.replace(/''/g, "'") ?? qualified?.[2]
  const sheet =
    name === undefined
      ? sheets.find((item) => item.id === defaultSheetId)
      : sheets.find((item) => item.name.toLowerCase() === name.toLowerCase())
  if (!sheet) throw new Error('chart_source_missing')
  const range = (qualified?.[3] ?? text).replace(/\$/g, '').toUpperCase()
  const bounds = parseRange(range)
  if (bounds.startRow !== bounds.endRow && bounds.startColumn !== bounds.endColumn)
    throw new Error('chart_vector_required')
  if (rangeCellCount(bounds) > 1000) throw new Error('chart_vector_limit')
  const absolute = (row: number, column: number) => `$${columnLabel(column)}$${row + 1}`
  return {
    sheetId: sheet.id,
    bounds,
    range: `${formatAddress(bounds.startRow, bounds.startColumn)}:${formatAddress(bounds.endRow, bounds.endColumn)}`,
    ref: `'${sheet.name.replace(/'/g, "''")}'!${absolute(bounds.startRow, bounds.startColumn)}:${absolute(bounds.endRow, bounds.endColumn)}`,
  }
}

/** Match the desktop vector editor's finite numeric coercion. */
export function chartNumericValue(
  value: string | number | boolean | null | undefined,
): number | null {
  if (typeof value === 'number') return Number.isFinite(value) ? value : null
  if (typeof value !== 'string' || !value.trim()) return null
  const numeric = Number(value.trim())
  return Number.isFinite(numeric) ? numeric : null
}

export interface ChartSeriesRangeEdit {
  index: number
  name?: string | undefined
  sheetId?: string | undefined
  valuesRange?: string | undefined
  categoriesRange?: string | undefined
}

/** Host/snapshot-derived caches and references always travel together. */
export function chartSeriesRangeEdits(
  book: WorkbookSnapshot,
  ownerId: string,
  seriesCount: number,
  entries: readonly ChartSeriesRangeEdit[],
  assertReady?: (sheetId: string, address: string) => void,
): NonNullable<ChartStateEdit['series']> {
  const edits = entries.map((entry) => {
    if (entry.index >= seriesCount) throw new Error('unknown_chart_series')
    const vector = (input: string) => {
      const resolved = resolveChartRange(input, book.sheets, entry.sheetId ?? ownerId)
      const sheet = book.sheets.find((item) => item.id === resolved.sheetId)!
      return {
        ref: resolved.ref,
        values: rangeAddresses(resolved.bounds).map((address) => {
          assertReady?.(sheet.id, address)
          return sheet.cells[address]?.value ?? null
        }),
      }
    }
    const values = entry.valuesRange === undefined ? undefined : vector(entry.valuesRange)
    const categories =
      entry.categoriesRange === undefined ? undefined : vector(entry.categoriesRange)
    return {
      index: entry.index,
      ...(entry.name === undefined ? {} : { name: entry.name }),
      ...(values
        ? {
            valuesRef: values.ref,
            values: values.values.map((value) => chartNumericValue(value) ?? 0),
          }
        : {}),
      ...(categories
        ? {
            categoriesRef: categories.ref,
            categories: categories.values.map((value) => String(value ?? '').slice(0, 255)),
          }
        : {}),
    }
  })
  const merged = new Map<number, NonNullable<ChartStateEdit['series']>[number]>()
  for (const entry of edits) merged.set(entry.index, { ...merged.get(entry.index), ...entry })
  return [...merged.values()]
}
