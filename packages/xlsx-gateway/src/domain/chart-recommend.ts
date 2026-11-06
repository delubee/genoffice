import { isTimeLike, hasNumericYearAxis } from './chart-data-kind'
export { hasNumericYearAxis } from './chart-data-kind'
import { chartDataFromValues, type ChartGridValue, type ParsedChartData } from './chart-visual'

/// Excel's Insert → Recommended Charts, heuristic edition: rank chart kinds
/// by the shape of the selected data. Pure so it unit-tests without Univer.

export type RecommendedKind =
  'column' | 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter' | 'radar' | 'combo'

export type RecommendReason =
  'time' | 'proportion' | 'correlation' | 'comparison' | 'manyPoints' | 'longLabels' | 'mixedScales'

export interface ChartRecommendations {
  parsed: ParsedChartData
  items: { kind: RecommendedKind; reason: RecommendReason }[]
}

export function recommendCharts(
  values: readonly (readonly ChartGridValue[])[],
): ChartRecommendations | null {
  let parsed = chartDataFromValues(values)
  if (!parsed) return null
  if (hasNumericYearAxis(parsed)) {
    parsed = chartDataFromValues(values, { numericCategoryColumn: true }) ?? parsed
  }
  const rows = parsed.categories.length
  const seriesCount = parsed.series.length
  const firstSeries = parsed.series[0]
  if (!firstSeries || rows === 0) return null

  const time = parsed.hasCategoryColumn && isTimeLike(parsed.categories)
  const allPositive = parsed.series.every((series) => series.values.every((value) => value >= 0))
  const longLabels =
    parsed.hasCategoryColumn &&
    parsed.categories.reduce((sum, label) => sum + label.length, 0) / rows > 12
  // Two numeric vectors with no label column reads as x/y pairs.
  const scatterCandidate = !parsed.hasCategoryColumn && seriesCount >= 2 && rows >= 5
  const spread = (series: { values: number[] }): number =>
    series.values.reduce((max, value) => Math.max(max, Math.abs(value)), 0)
  const secondSeries = parsed.series[1]
  const mixedScales =
    seriesCount === 2 &&
    secondSeries !== undefined &&
    spread(firstSeries) > 0 &&
    spread(secondSeries) > 0 &&
    (spread(firstSeries) / spread(secondSeries) > 10 ||
      spread(secondSeries) / spread(firstSeries) > 10)

  const items: ChartRecommendations['items'] = []
  const add = (kind: RecommendedKind, reason: RecommendReason): void => {
    if (!items.some((item) => item.kind === kind)) items.push({ kind, reason })
  }

  if (scatterCandidate) add('scatter', 'correlation')
  if (time) {
    add('line', 'time')
    if (seriesCount > 1 && allPositive) add('area', 'time')
    if (rows <= 12) add('column', 'time')
  }
  if (seriesCount === 1 && rows >= 2 && rows <= 8 && allPositive) {
    add('pie', 'proportion')
    add('doughnut', 'proportion')
  }
  if (mixedScales) add('combo', 'mixedScales')
  if (longLabels) add('bar', 'longLabels')
  if (rows > 20 || seriesCount >= 4) add('line', 'manyPoints')
  add('column', 'comparison')
  add('bar', 'comparison')
  if (!time && rows >= 3 && rows <= 8 && seriesCount >= 2 && seriesCount <= 3) {
    add('radar', 'comparison')
  }
  add('line', 'comparison')

  return { parsed, items: items.slice(0, 5) }
}
