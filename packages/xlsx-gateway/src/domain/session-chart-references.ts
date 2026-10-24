import { parseRange, rangeAddresses } from './cell-address'
import type { WorkbookSnapshot } from './workbook.types'

function readReference(book: WorkbookSnapshot, reference: string) {
  const match = /^(?:'((?:[^']|'')*)'|([^!]+))!(.+)$/.exec(reference)
  if (!match) throw new Error('unsupported_chart_reference')
  const name = match[1]?.replace(/''/g, "'") ?? match[2]!
  const sheet = book.sheets.find((item) => item.name === name)
  if (!sheet) throw new Error('chart_source_missing')
  const bounds = parseRange(match[3]!)
  if ((bounds.endRow - bounds.startRow + 1) * (bounds.endColumn - bounds.startColumn + 1) > 2000)
    throw new Error('chart_source_limit')
  return {
    sheetId: sheet.id,
    bounds,
    values: rangeAddresses(bounds).map((address) => sheet.cells[address]?.value ?? null),
  }
}

/** Updates JSON chart caches without retaining renderer callbacks or handles. */
export function refreshChartReferences(
  book: WorkbookSnapshot,
  completeSheetIds: readonly string[],
  referenceRanges: Record<string, string[]> = {},
) {
  for (const owner of book.sheets)
    for (const visual of owner.visuals ?? [])
      for (const series of visual.chart.series) {
        for (const key of ['valuesRef', 'categoriesRef', 'nameRef'] as const) {
          if (!series[key]) continue
          const source = readReference(book, series[key]!)
          const covered = (referenceRanges[source.sheetId] ?? []).some((range) => {
            const bounds = parseRange(range)
            return (
              bounds.startRow <= source.bounds.startRow &&
              bounds.endRow >= source.bounds.endRow &&
              bounds.startColumn <= source.bounds.startColumn &&
              bounds.endColumn >= source.bounds.endColumn
            )
          })
          if (!completeSheetIds.includes(source.sheetId) && !covered)
            throw new Error('incomplete_chart_material')
          if (key === 'valuesRef') {
            series.values = source.values.map((value) => (typeof value === 'number' ? value : 0))
            const blanks = source.values.flatMap((value, index) =>
              typeof value !== 'number' ? [index] : [],
            )
            if (blanks.length) series.blanks = blanks
            else delete series.blanks
          } else if (key === 'categoriesRef')
            series.categories = source.values.map((value) => (value === null ? '' : String(value)))
          else series.name = source.values[0] === null ? '' : String(source.values[0])
        }
      }
}
