import { parseRange, rangeCellCount } from './cell-address'
import { applyChartStateEdit, type ChartStateEdit, type ChartVisualState } from './chart-visual'

function pointCount(series: ChartVisualState['series'][number]): number {
  let count = Math.max(series.values.length, series.categories.length)
  for (const ref of [series.valuesRef, series.categoriesRef]) {
    if (!ref) continue
    try {
      count = Math.max(
        count,
        rangeCellCount(parseRange(ref.slice(ref.lastIndexOf('!') + 1).replace(/\$/g, ''))),
      )
    } catch {
      /* Cached count remains authoritative for non-cell references. */
    }
  }
  return count
}

/** Validate effective settings before authoring either session or source XML. */
export function assertChartFormat(chart: ChartVisualState, edit: ChartStateEdit): void {
  const next = applyChartStateEdit(chart, edit)
  for (const index of new Set([
    ...Object.keys(edit.seriesColors ?? {}),
    ...Object.keys(edit.pointColors ?? {}),
  ])) {
    const series = next.series[Number(index)]
    if (!series) throw new Error('unknown_chart_series')
    if (
      Object.keys(edit.pointColors?.[index] ?? {}).some(
        (point) => Number(point) >= pointCount(series),
      )
    )
      throw new Error('unknown_chart_point')
  }
  const pie = next.chartTypes.some((type) => /^(pieChart|doughnutChart)$/.test(type))
  if (
    pie &&
    (edit.gridlines !== undefined || edit.valueAxis !== undefined || edit.axisTitles !== undefined)
  )
    throw new Error('unsupported_chart_axes')
  if (edit.gapWidthPct !== undefined && !next.chartTypes.includes('barChart'))
    throw new Error('unsupported_chart_gap')
  if (edit.holeSizePct !== undefined && !next.chartTypes.includes('doughnutChart'))
    throw new Error('unsupported_chart_hole')
  if ((edit.explosionPct !== undefined || edit.pointExplosions !== undefined) && !pie)
    throw new Error('unsupported_chart_explosion')
  if (
    edit.pointExplosions &&
    Object.keys(edit.pointExplosions).some(
      (index) => Number(index) >= (next.series[0] ? pointCount(next.series[0]) : 0),
    )
  )
    throw new Error('unknown_chart_point')
  const axis = next.valueAxis
  if (axis?.min !== undefined && axis.max !== undefined && axis.min >= axis.max)
    throw new Error('invalid_chart_axis_bounds')
}
