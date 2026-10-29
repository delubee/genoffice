import type { ParsedChartData } from './chart-visual'

const MONTH_NAMES = new Set(
  'jan feb mar apr may jun jul aug sep oct nov dec january february march april june july august september october november december'.split(
    ' ',
  ),
)

export function isTimeLike(categories: readonly string[]): boolean {
  if (categories.length < 3) return false
  let dated = 0
  let numeric = 0
  let lastNumber = Number.NEGATIVE_INFINITY
  let increasing = true
  for (const raw of categories) {
    const label = raw.trim()
    if (label === '') return false
    const lower = label.toLowerCase().replace(/[.月]$/u, '')
    if (MONTH_NAMES.has(lower) || /^\d{1,2}月$/u.test(label) || /^[qQ][1-4]$/.test(label)) {
      dated += 1
      continue
    }
    if (!Number.isNaN(Date.parse(label)) && /[-/年]/u.test(label)) {
      dated += 1
      continue
    }
    const asNumber = Number(label)
    if (Number.isInteger(asNumber) && asNumber >= 1900 && asNumber <= 2200) {
      numeric += 1
      if (asNumber <= lastNumber) increasing = false
      lastNumber = asNumber
      continue
    }
    return false
  }
  return dated === categories.length || (numeric === categories.length && increasing)
}

/// Year/Value tables parse with the year column as a bare numeric series
/// (hasCategoryColumn false), which would read as x/y scatter pairs. Detect
/// that shape so the year vector can become the category axis instead.
export function hasNumericYearAxis(parsed: ParsedChartData): boolean {
  const firstSeries = parsed.series[0]
  return (
    !parsed.hasCategoryColumn &&
    parsed.series.length >= 2 &&
    firstSeries !== undefined &&
    isTimeLike(firstSeries.values.map((value) => String(value)))
  )
}
