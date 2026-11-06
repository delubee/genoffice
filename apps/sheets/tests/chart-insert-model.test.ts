import { expect, it } from 'vitest'
import { buildChartVisual, chartKindFromVisual } from '@genoffice/xlsx-gateway/domain/chart-visual'

const input = {
  id: 'chart',
  sheetId: 'logical',
  sheetName: "Data's",
  dataRange: 'D8:F11',
  values: [
    ['Year', 'Sales', 'Cost'],
    [2020, 12, 6],
    [2021, 8, 10],
    [2022, 18, 14],
  ],
}
it('uses year categories for ordinary charts and the first numeric vector as X for scatter', () => {
  for (const chartType of [
    'column',
    'bar',
    'line',
    'area',
    'pie',
    'doughnut',
    'scatter',
    'radar',
    'combo',
  ] as const) {
    const visual = buildChartVisual({ ...input, chartType })
    expect(chartKindFromVisual(visual.chart)).toBe(chartType)
    expect(visual.chart.series.map((series) => series.values)).toEqual([
      [12, 8, 18],
      [6, 10, 14],
    ])
    expect(visual.chart.series[0]).toMatchObject({
      categories: ['2020', '2021', '2022'],
      valuesRef: "'Data''s'!$E$9:$E$11",
      categoriesRef: "'Data''s'!$D$9:$D$11",
    })
    expect(visual.chart.dataLabels).toBeUndefined()
    expect(visual.anchor).toMatchObject({ fromRow: 7, fromColumn: 7, toRow: 22, toColumn: 14 })
  }
})
it('does not plot a numeric X vector as an additional Y series', () => {
  const visual = buildChartVisual({
    ...input,
    chartType: 'scatter',
    values: [
      ['X', 'Y', 'Z'],
      [2, 12, 6],
      [4, 8, 10],
      [6, 18, 14],
    ],
  })
  expect(visual.chart.series.map((series) => series.values)).toEqual([
    [12, 8, 18],
    [6, 10, 14],
  ])
  expect(visual.chart.series[0]!.categories).toEqual(['2', '4', '6'])
})
it('rejects unknown or mixed plot sets instead of guessing a serializable chart kind', () => {
  for (const chartTypes of [
    [],
    ['surfaceChart'],
    ['barChart', 'areaChart'],
    ['barChart', 'barChart'],
    ['barChart', 'lineChart', 'areaChart'],
  ])
    expect(chartKindFromVisual({ chartTypes })).toBeNull()
})
