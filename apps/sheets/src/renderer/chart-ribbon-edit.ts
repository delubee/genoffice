import type { ChartEditData } from './WorkbookVisuals'
import type { SelectedChartRibbon } from './SheetsRibbon'
import {
  CHART_TYPE_COMMANDS,
  CHART_LEGEND_COMMANDS,
  CHART_LABEL_COMMANDS,
  CHART_PALETTES,
} from './app-constants'

/** Identical chart-design presets for the desktop and embedded editor. */
export function chartRibbonEdit(
  command: string,
  selected: Pick<
    SelectedChartRibbon,
    'canLabel' | 'isPie' | 'categoryCount' | 'seriesCount'
  > | null,
): ChartEditData | undefined {
  if (command.startsWith('chart-type-')) {
    const chartType = CHART_TYPE_COMMANDS.find((type) => command === `chart-type-${type}`)
    if (chartType) return { chartType }
  } else if (command.startsWith('chart-legend:')) {
    const legend = CHART_LEGEND_COMMANDS.find((pos) => command === `chart-legend:${pos}`)
    if (legend) return { legend }
  } else if (command.startsWith('chart-labels:')) {
    const labels = CHART_LABEL_COMMANDS.find((mode) => command === `chart-labels:${mode}`)
    // scatter/radar cannot save plot-level labels — fail here, not at ⌘S.
    if (labels && selected?.canLabel) {
      return { dataLabels: labels }
    }
  } else if (command.startsWith('chart-grouping:')) {
    const grouping = (['clustered', 'stacked', 'percentStacked'] as const).find(
      (mode) => command === `chart-grouping:${mode}`,
    )
    if (grouping) return { grouping }
  } else if (command.startsWith('chart-colors:')) {
    const palette = CHART_PALETTES[command.slice('chart-colors:'.length)]
    if (!palette) return undefined
    if (selected?.isPie) {
      // Pie recolor varies by point, not by series.
      const count = Math.min(selected.categoryCount, 64)
      if (count === 0) return undefined
      return {
        pointColors: {
          '0': Object.fromEntries(
            Array.from({ length: count }, (_, index) => [
              String(index),
              palette[index % palette.length] ?? '#4472c4',
            ]),
          ),
        },
      }
    }
    const count = Math.min(selected?.seriesCount ?? 0, 24)
    if (count === 0) return undefined
    return {
      seriesColors: Object.fromEntries(
        Array.from({ length: count }, (_, index) => [
          String(index),
          palette[index % palette.length] ?? '#4472c4',
        ]),
      ),
    }
  } else if (command.startsWith('chart-layout:')) {
    // Quick Layout presets: legend placement + data labels together,
    // labels picked per chart family (pie vs. axis charts).
    const labelsOn = selected?.isPie ? ('category-percent' as const) : ('value' as const)
    const preset = {
      'chart-layout:1': { legend: 'right', dataLabels: labelsOn },
      'chart-layout:2': { legend: 'top', dataLabels: labelsOn },
      'chart-layout:3': { legend: 'bottom', dataLabels: 'none' },
      'chart-layout:4': { legend: 'none', dataLabels: labelsOn },
    }[command] as Pick<ChartEditData, 'legend' | 'dataLabels'> | undefined
    if (preset) {
      // Label-less families take the preset's legend only.
      const applied = selected?.canLabel
        ? preset
        : { ...(preset.legend === undefined ? {} : { legend: preset.legend }) }
      return applied
    }
  } else if (command.startsWith('chart-title:')) {
    return {
      title: command.slice('chart-title:'.length).slice(0, 255),
    }
  } else if (command.startsWith('chart-axis-cat:')) {
    const text = command.slice('chart-axis-cat:'.length).slice(0, 255)
    return { axisTitles: { category: text || null } }
  } else if (command.startsWith('chart-axis-val:')) {
    const text = command.slice('chart-axis-val:'.length).slice(0, 255)
    return { axisTitles: { value: text || null } }
  }
}
