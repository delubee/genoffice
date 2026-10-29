import { Fragment, useEffect, useRef, useState } from 'react'
import { numfmt } from '@univerjs/core'
import {
  chartCategoryFormat,
  formatCategoryLabel,
  scatterAxisBounds,
  valueAxisScale,
} from '@genoffice/xlsx-gateway/domain/chart-visual'
import type { WorkbookVisualObject } from '../shared/desktop-api'
import { t } from './i18n/locale'
import {
  AXIS_LABEL_PT,
  AXIS_TITLE_PT,
  CHAR_EM,
  axisSideReservePt,
  bottomAxisLayout,
  chartAreaInk,
  chartTextStyle,
  chartTextUnit,
  valueAxisLayout,
} from './chart-text-scale'
import type { ChartAxisText, ChartTextBox } from './chart-text-scale'
export type ChartMetadata = NonNullable<WorkbookVisualObject['chart']>
export type ChartSeries = ChartMetadata['series'][number]
type PointLabel = NonNullable<ChartSeries['pointLabels']>[number]
export type ChartElementRef =
  | { kind: 'chart' | 'title' | 'legend' | 'value-axis' | 'category-axis' }
  | { kind: 'series'; seriesIndex: number }
  | { kind: 'point'; seriesIndex: number; pointIndex: number }

const chartColors = [
  '#4472c4',
  '#ed7d31',
  '#a5a5a5',
  '#ffc000',
  '#5b9bd5',
  '#70ad47',
  '#264478',
  '#9e480e',
  '#636363',
  '#997300',
]

type SeriesLike = {
  readonly color?: string | undefined
  readonly pointColors?: readonly { readonly index: number; readonly color: string }[] | undefined
}

export function seriesColor(series: SeriesLike | undefined, index: number): string {
  return series?.color ?? chartColors[index % chartColors.length] ?? '#4472c4'
}

/// Excel narrowing: first click on a series member selects the series, a
/// second click on the same series narrows to the single point.
function narrowSelection(
  current: ChartElementRef | null,
  seriesIndex: number,
  pointIndex: number,
): ChartElementRef {
  const sameSeries =
    current &&
    (current.kind === 'series' || current.kind === 'point') &&
    current.seriesIndex === seriesIndex
  return sameSeries ? { kind: 'point', seriesIndex, pointIndex } : { kind: 'series', seriesIndex }
}

interface ChartElementProps {
  readonly onElement?: ((element: ChartElementRef) => void) | undefined
  readonly selectedEl?: ChartElementRef | null | undefined
}

function isSelectedPoint(
  element: ChartElementRef | null | undefined,
  seriesIndex: number,
  pointIndex: number,
): boolean {
  if (!element) return false
  if (element.kind === 'series') return element.seriesIndex === seriesIndex
  if (element.kind === 'point') {
    return element.seriesIndex === seriesIndex && element.pointIndex === pointIndex
  }
  return false
}

/** The exact desktop chart body, independent of document, media and editor services. */
export function SheetChartContent({
  chart,
  frameWidth,
  selectedEl,
  onElement: selectElement,
}: {
  readonly chart: ChartMetadata
  readonly frameWidth?: number | undefined
} & ChartElementProps): React.JSX.Element | null {
  const bodyRef = useRef<HTMLDivElement | null>(null)
  const populated = chart.series.filter((series) => series.values.length > 0)
  const primarySeries = populated[0]
  const chartKindKey = chart.chartTypes.join()
  const hasData = primarySeries !== undefined
  useEffect(() => {
    bodyRef.current?.querySelector('svg.chart-svg')?.setAttribute('aria-label', chart.title)
  }, [chart.title, chartKindKey, hasData])
  // The SVG is re-created on a chart-type edit and when hydration delivers
  // the first points, so the measurement re-attaches on both.
  const textBox = useChartTextBox(
    bodyRef,
    frameWidth,
    `${chart.chartTypes.join()}|${primarySeries !== undefined}`,
  )
  if (!primarySeries) {
    // No cached points and nothing hydrated: Excel shows a silent empty
    // frame, never a diagnostic.
    return null
  }
  const types = chart.chartTypes
  const isDoughnut = types.includes('doughnutChart')
  const isPie = types.includes('pieChart') || isDoughnut
  const isScatter = types.includes('scatterChart')
  const isCombo = types.includes('barChart') && types.includes('lineChart')
  const comboLineIdx = isCombo ? comboLineIndices(populated) : new Set<number>()
  const comboBarIdx = comboBarIndices(populated, comboLineIdx)
  const comboBars = comboBarIdx.map((index) => populated[index] as ChartSeries)
  const comboLines = populated.filter((_, index) => comboLineIdx.has(index))
  // Element selection is resolved against chart.series, so the drawn bars
  // carry their source indices instead of their filtered positions.
  const populatedIdx = chart.series.flatMap((series, index) =>
    series.values.length > 0 ? [index] : [],
  )
  const barSeriesIdx = isCombo
    ? comboBarIdx.map((index) => populatedIdx[index] as number)
    : populatedIdx
  const isArea = types.includes('areaChart') && !types.includes('barChart')
  const isLine =
    (types.includes('lineChart') && !types.includes('barChart')) ||
    (isCombo && comboBars.length === 0)
  const isRadar = types.includes('radarChart') && !types.includes('barChart')
  const categoryFormat = chartCategoryFormat(chart)
  // Per-side axes (xAxis/yAxis, paired by c:axPos) win over the legacy
  // element-kind fields — scatter charts have two value axes and the legacy
  // pairing flips their titles. The slots here are SEMANTIC
  // (category/value); on a horizontal bar chart the category axis sits on
  // the left (yAxis position) and the value axis on the bottom, so the
  // positional fields map crosswise — BarChart swaps them back for layout.
  const isHorizontalBar = chart.barDirection === 'bar'
  const effectiveAxisTitles = {
    category: (isHorizontalBar ? chart.yAxis : chart.xAxis)?.title ?? chart.axisTitles?.category,
    value: (isHorizontalBar ? chart.xAxis : chart.yAxis)?.title ?? chart.axisTitles?.value,
  }
  // The value axis sits on the bottom (xAxis slot) for horizontal bars.
  const valueAxisSide = isHorizontalBar ? chart.xAxis : chart.yAxis
  const valueScaleInput = {
    min: valueAxisSide?.min ?? chart.valueAxis?.min,
    max: valueAxisSide?.max ?? chart.valueAxis?.max,
    majorUnit: valueAxisSide?.majorUnit,
    numFmt: valueAxisSide?.numFmt,
    hidden: valueAxisSide?.hidden === true,
    position: valueAxisSide?.position,
    labelSize: valueAxisSide?.labelSize,
    labelColor: valueAxisSide?.labelColor,
    titleSize: valueAxisSide?.titleSize,
    titleColor: valueAxisSide?.titleColor,
    displayUnit: valueAxisSide?.displayUnit,
    displayUnitLabel: valueAxisSide?.displayUnitLabel,
  }
  const categoryAxisSide = isHorizontalBar ? chart.yAxis : chart.xAxis
  const categoryAxisHidden = categoryAxisSide?.hidden === true
  const categoryAxisText: ChartAxisText = {
    labelSize: categoryAxisSide?.labelSize,
    labelColor: categoryAxisSide?.labelColor,
    titleSize: categoryAxisSide?.titleSize,
    titleColor: categoryAxisSide?.titleColor,
  }
  const plotAreaFill = chart.plotAreaFill
  return (
    <>
      {chart.title !== '' && (
        <figcaption
          className={selectedEl?.kind === 'title' ? 'chart-el-selected' : undefined}
          style={{
            ...(chart.titleStyle?.size ? { fontSize: `${chart.titleStyle.size}pt` } : {}),
            ...(chart.titleStyle?.bold === false ? { fontWeight: 400 } : {}),
            ...(chart.titleStyle?.color ? { color: chart.titleStyle.color } : {}),
          }}
          onClick={
            selectElement
              ? (event) => {
                  event.stopPropagation()
                  selectElement({ kind: 'title' })
                }
              : undefined
          }
        >
          {chart.title}
        </figcaption>
      )}
      <div className={`chart-body chart-legend-${chart.legend ?? 'default'}`} ref={bodyRef}>
        {isPie ? (
          <PieChart
            series={primarySeries}
            isDoughnut={isDoughnut}
            legend={chart.legend}
            dataLabels={chart.dataLabels}
            dataLabelPosition={chart.dataLabelPosition}
            dataLabelFormat={chart.dataLabelFormat}
            dataLabelStyle={chart.dataLabelStyle}
            holeSizePct={chart.holeSizePct}
            categoryFormat={categoryFormat}
            onElement={selectElement}
            selectedEl={selectedEl}
          />
        ) : isRadar ? (
          <RadarChart
            seriesList={populated}
            dataLabels={chart.dataLabels}
            categoryFormat={categoryFormat}
            textBox={textBox}
            categoryAxis={categoryAxisText}
            valueAxis={valueScaleInput}
            onElement={selectElement}
            selectedEl={selectedEl}
          />
        ) : isScatter ? (
          <ScatterChart
            seriesList={populated}
            axisTitles={effectiveAxisTitles}
            dataLabels={chart.dataLabels}
            gridlines={chart.gridlines}
            valueAxis={valueScaleInput}
            xAxis={chart.xAxis}
            scatterStyle={chart.scatterStyle}
            categoryFormat={categoryFormat}
            textBox={textBox}
            categoryAxis={categoryAxisText}
            plotAreaFill={plotAreaFill}
            onElement={selectElement}
            selectedEl={selectedEl}
          />
        ) : isArea ? (
          <AreaChart
            seriesList={populated}
            axisTitles={effectiveAxisTitles}
            dataLabels={chart.dataLabels}
            grouping={chart.grouping}
            gridlines={chart.gridlines}
            valueAxis={valueScaleInput}
            categoryFormat={categoryFormat}
            categoryHidden={categoryAxisHidden}
            textBox={textBox}
            categoryAxis={categoryAxisText}
            plotAreaFill={plotAreaFill}
            onElement={selectElement}
            selectedEl={selectedEl}
          />
        ) : isLine ? (
          <LineChart
            seriesList={populated}
            axisTitles={effectiveAxisTitles}
            dataLabels={chart.dataLabels}
            grouping={chart.grouping}
            gridlines={chart.gridlines}
            valueAxis={valueScaleInput}
            categoryFormat={categoryFormat}
            categoryHidden={categoryAxisHidden}
            lineMarkers={chart.lineMarkers}
            dispBlanksAs={chart.dispBlanksAs}
            textBox={textBox}
            categoryAxis={categoryAxisText}
            plotAreaFill={plotAreaFill}
            onElement={selectElement}
            selectedEl={selectedEl}
          />
        ) : (
          <BarChart
            seriesList={isCombo ? comboBars : populated}
            seriesIndices={barSeriesIdx}
            categorySeries={isCombo ? comboCategorySeries(populated) : undefined}
            isHorizontal={chart.barDirection === 'bar'}
            lineSeriesList={isCombo ? comboLines : undefined}
            axisTitles={effectiveAxisTitles}
            dataLabels={chart.dataLabels}
            dataLabelPosition={chart.dataLabelPosition}
            dataLabelFormat={chart.dataLabelFormat}
            dataLabelStyle={chart.dataLabelStyle}
            grouping={chart.grouping}
            gridlines={chart.gridlines}
            valueAxis={valueScaleInput}
            secondaryAxis={chart.secondaryYAxis}
            gapWidthPct={chart.gapWidthPct}
            categoryFormat={categoryFormat}
            categoryReversed={(isHorizontalBar ? chart.yAxis : chart.xAxis)?.reversed === true}
            categoryHidden={categoryAxisHidden}
            lineMarkers={chart.lineMarkers}
            dispBlanksAs={chart.dispBlanksAs}
            textBox={textBox}
            categoryAxis={categoryAxisText}
            plotAreaFill={plotAreaFill}
            onElement={selectElement}
            selectedEl={selectedEl}
          />
        )}
        {!isPie &&
          chart.legend !== 'none' &&
          (chart.legend !== undefined || populated.length > 1) && (
            <SeriesLegend
              seriesList={populated}
              lineSwatches={legendUsesLineSwatches(types)}
              lineSwatchIndices={isCombo ? comboLineIdx : undefined}
              selected={selectedEl?.kind === 'legend'}
              onSelect={
                selectElement
                  ? (event) => {
                      event.stopPropagation()
                      selectElement({ kind: 'legend' })
                    }
                  : undefined
              }
            />
          )}
      </div>
    </>
  )
}

export function sheetChartAreaStyle(chart: ChartMetadata): React.CSSProperties {
  // The chart area fill sits under the HTML legend too; the file's axis or
  // title text color (else light ink on a dark fill) stays readable on it
  // and reaches every SVG label through --chart-ink.
  const areaInk = chart.chartAreaFill
    ? chartAreaInk(
        chart.chartAreaFill,
        chart.yAxis?.labelColor ?? chart.xAxis?.labelColor ?? chart.titleStyle?.color,
      )
    : undefined
  return {
    ...(chart.chartAreaFill ? { background: chart.chartAreaFill } : {}),
    ...(areaInk ? { color: areaInk, '--chart-ink': areaInk } : {}),
    ...(chart.dataLabelStyle?.color ? { '--chart-dlbl-color': chart.dataLabelStyle.color } : {}),
    ...(chart.dataLabelStyle?.size ? { '--chart-dlbl-pt': chart.dataLabelStyle.size } : {}),
  } as React.CSSProperties
}

/** Read-only desktop chart; editing and host actions are supplied by its owner. */
export function SheetChart({
  chart,
  frameWidth,
}: {
  readonly chart: ChartMetadata
  readonly frameWidth?: number | undefined
}) {
  if (!chart.series.some((series) => series.values.length > 0))
    return <figure className="xlsx-chart" style={sheetChartAreaStyle(chart)} />
  return (
    <figure
      className="xlsx-chart"
      role="img"
      aria-label={chart.title}
      style={sheetChartAreaStyle(chart)}
    >
      <SheetChartContent chart={chart} frameWidth={frameWidth} />
    </figure>
  )
}
type ChartAxisTitles = ChartMetadata['axisTitles']
type ChartDataLabels = ChartMetadata['dataLabels']
type ChartGrouping = ChartMetadata['grouping']

/// Category title along the bottom, value title rotated up the left edge
/// (already swapped by the caller for horizontal bars).
function AxisTitleTexts({
  bottom,
  left,
  bottomY = 317,
  leftEdge = 0,
  leftUnits,
  bottomUnits,
  titleFont = 10,
  labelFont = 9,
}: {
  readonly bottom?: string | null | undefined
  readonly left?: string | null | undefined
  /// Charts with an outer category tier extend the viewBox and push the
  /// bottom title below the group-label band; a top value axis moves the
  /// value title (and its unit label) above the plot.
  readonly bottomY?: number
  /// Left edge of the viewBox; charts that widen it for long value labels
  /// pass the (negative) extension so the title column moves with it.
  readonly leftEdge?: number
  /// c:dispUnitsLbl of the left / bottom value axis, drawn beside its title.
  readonly leftUnits?: string | undefined
  readonly bottomUnits?: string | undefined
  readonly titleFont?: number
  readonly labelFont?: number
}): React.JSX.Element {
  // Rotated text hangs left of its baseline by the ascent; the unit label
  // column follows the title column.
  const leftX = leftEdge + 2 + 0.9 * titleFont
  const unitsX = left
    ? leftX + 0.25 * titleFont + 0.85 * labelFont
    : leftEdge + 2 + 0.85 * labelFont
  return (
    <g>
      {bottom ? (
        <text x="320" y={bottomY} textAnchor="middle" className="axis-title">
          {truncateLabel(bottom, 60)}
        </text>
      ) : null}
      {bottomUnits ? (
        <text x="580" y={bottomY} textAnchor="end" className="axis-label">
          {bottomUnits}
        </text>
      ) : null}
      {left ? (
        <text
          x={leftX}
          y="155"
          transform={`rotate(-90 ${leftX} 155)`}
          textAnchor="middle"
          className="axis-title"
        >
          {truncateLabel(left, 40)}
        </text>
      ) : null}
      {leftUnits ? (
        <text
          x={unitsX}
          y="155"
          transform={`rotate(-90 ${unitsX} 155)`}
          textAnchor="middle"
          className="axis-label"
        >
          {leftUnits}
        </text>
      ) : null}
    </g>
  )
}

/// Rendered size of the chart SVG plus the sheet zoom (figure width over the
/// unzoomed frame), so chart text can be sized in points on any frame.
function useChartTextBox(
  bodyRef: React.RefObject<HTMLDivElement | null>,
  frameWidth: number | undefined,
  chartKey: string,
): ChartTextBox | undefined {
  const [box, setBox] = useState<ChartTextBox | undefined>(undefined)
  useEffect(() => {
    const body = bodyRef.current
    const svg = body?.querySelector('svg.chart-svg')
    const figure = body?.parentElement
    if (!body || !svg || typeof ResizeObserver === 'undefined') return undefined
    const measure = (): void => {
      const rect = svg.getBoundingClientRect()
      const zoom = frameWidth && figure ? figure.getBoundingClientRect().width / frameWidth || 1 : 1
      const next = {
        width: Math.round(rect.width),
        height: Math.round(rect.height),
        zoom: Number(zoom.toFixed(3)),
      }
      setBox((previous) =>
        previous &&
        previous.width === next.width &&
        previous.height === next.height &&
        previous.zoom === next.zoom
          ? previous
          : next,
      )
    }
    const observer = new ResizeObserver(measure)
    observer.observe(svg)
    if (figure) observer.observe(figure)
    measure()
    return () => observer.disconnect()
  }, [bodyRef, frameWidth, chartKey])
  return box
}

/// Widest value tick label in em, for the gutter the labels need; a
/// deleted axis draws none and needs no gutter.
function valueLabelEm(
  ticks: readonly number[],
  displayUnit: number | undefined,
  numberFormat: string | undefined,
  hidden = false,
): number {
  if (hidden) return 0
  let widest = 0
  for (const tick of ticks) {
    widest = Math.max(
      widest,
      formatAxisValue(displayUnit ? tick / displayUnit : tick, numberFormat).length,
    )
  }
  return widest * CHAR_EM
}

/// Excel's automatic slice cycle past the six accents: the same accents
/// lightened (HSL L' = 0.6·L + 0.4), verified against Excel-rendered pies.
const pieSliceColorsLight = ['#8faadc', '#f4b183', '#c9c9c9', '#ffd966', '#9dc3e6', '#a9d18e']

/// Pie/doughnut slices color per point: explicit `c:dPt` fills win, the
/// Office palette cycles underneath (Excel's varyColors default) — six
/// accents, then their lighter variants, then around again.
export function pieSliceColor(series: SeriesLike, index: number): string {
  const explicit = series.pointColors?.find((point) => point.index === index)?.color
  if (explicit) return explicit
  const cycle = index % 12
  return (cycle < 6 ? chartColors[cycle] : pieSliceColorsLight[cycle - 6]) ?? '#4472c4'
}

function SeriesLegend({
  seriesList,
  selected,
  onSelect,
  lineSwatches = false,
  lineSwatchIndices,
}: {
  readonly seriesList: readonly ChartSeries[]
  readonly selected?: boolean | undefined
  readonly onSelect?: ((event: React.MouseEvent) => void) | undefined
  /// Line-family charts: swatches mirror the drawn stroke color.
  readonly lineSwatches?: boolean | undefined
  /// Combo charts: the series drawn as lines.
  readonly lineSwatchIndices?: ReadonlySet<number> | undefined
}): React.JSX.Element {
  return (
    <div className={`chart-legend${selected ? ' chart-el-selected' : ''}`} onClick={onSelect}>
      {seriesList.slice(0, 8).map((series, index) => (
        <span key={`${series.name}-${index}`}>
          <i
            style={{
              background: legendSwatchColor(
                series,
                index,
                lineSwatches || lineSwatchIndices?.has(index) === true,
              ),
            }}
          />
          {truncateLabel(series.name, 18)}
        </span>
      ))}
      {seriesList.length > 8 && <span>{t('appMoreSeries', { count: seriesList.length - 8 })}</span>}
    </div>
  )
}

export function BarChart({
  seriesList,
  seriesIndices,
  categorySeries,
  lineSeriesList = [],
  isHorizontal,
  axisTitles,
  dataLabels,
  dataLabelPosition,
  dataLabelFormat,
  dataLabelStyle,
  grouping,
  gridlines,
  valueAxis,
  secondaryAxis,
  gapWidthPct,
  categoryFormat,
  categoryReversed,
  categoryHidden = false,
  lineMarkers,
  dispBlanksAs,
  textBox,
  categoryAxis,
  plotAreaFill,
  onElement,
  selectedEl,
}: {
  readonly seriesList: readonly ChartSeries[]
  /// chart.series index of each seriesList entry (selection targets).
  readonly seriesIndices?: readonly number[] | undefined
  readonly textBox?: ChartTextBox | undefined
  readonly categoryAxis?: ChartAxisText | undefined
  readonly plotAreaFill?: string | undefined
  /// Series whose cached categories label the axis; defaults to seriesList[0].
  readonly categorySeries?: ChartSeries | undefined
  readonly lineSeriesList?: readonly ChartSeries[] | undefined
  readonly isHorizontal: boolean
  /// Plot-level c:lineChart/c:marker flag for the combo lines.
  readonly lineMarkers?: boolean | undefined
  readonly dispBlanksAs?: ChartMetadata['dispBlanksAs']
  /// c:catAx orientation maxMin — categories read top-down (bar) /
  /// right-to-left (column) instead of Excel's minMax default.
  readonly categoryReversed?: boolean | undefined
  readonly categoryHidden?: boolean | undefined
  readonly axisTitles?: ChartAxisTitles
  readonly dataLabels?: ChartDataLabels
  readonly dataLabelPosition?: ChartLabelPosition
  readonly dataLabelFormat?: string | undefined
  readonly dataLabelStyle?: ChartMetadata['dataLabelStyle']
  readonly grouping?: ChartGrouping
  readonly gridlines?: boolean | undefined
  readonly valueAxis?: ChartValueAxis
  readonly secondaryAxis?:
    | {
        min?: number | undefined
        max?: number | undefined
        majorUnit?: number | undefined
        numFmt?: string | undefined
        hidden?: boolean | undefined
        displayUnit?: number | undefined
      }
    | undefined
  readonly gapWidthPct?: number | undefined
  readonly categoryFormat?: string | undefined
} & ChartElementProps): React.JSX.Element {
  const primary = categorySeries ?? seriesList[0]
  if (!primary || seriesList.length === 0) return <></>
  const categories = primary.categories.map((value) => formatCategoryLabel(value, categoryFormat))
  // Every series on the shared category axis sizes it, not just the one that
  // carries the labels.
  const pointCount = Math.max(
    primary.values.length,
    ...seriesList.map((series) => series.values.length),
    ...lineSeriesList.map((series) => series.values.length),
  )
  // 20 was too tight for real corpora (38-county bar charts); 48 keeps
  // bars ≥ ~5px in the 600px plot while very wide data still truncates.
  const visibleCount = Math.min(pointCount, 48)
  const chartIndex = (seriesIndex: number): number => seriesIndices?.[seriesIndex] ?? seriesIndex
  const isStacked =
    (grouping === 'stacked' || grouping === 'percentStacked') && seriesList.length > 1
  const isPercent = grouping === 'percentStacked' && seriesList.length > 1
  const categoryTotal = (index: number): number =>
    seriesList.reduce((sum, series) => sum + Math.max(0, series.values[index] ?? 0), 0)
  const barMax = isStacked
    ? Math.max(...Array.from({ length: visibleCount }, (_, index) => categoryTotal(index)), 0)
    : Math.max(...seriesList.flatMap((series) => [...series.values]), 0)
  // Stacked segments clamp negatives to 0 (categoryTotal), so only the plain
  // branch can plot below the baseline.
  const barMin = isStacked ? 0 : Math.min(...seriesList.flatMap((series) => [...series.values]))
  // Combo lines without their own value axis share the primary scale.
  const lineValues = lineSeriesList.flatMap((series) => [...series.values])
  const lineOnPrimary = lineSeriesList.length > 0 && secondaryAxis === undefined
  const bounds = isPercent
    ? { min: 0, max: 1, ticks: [0, 0.25, 0.5, 0.75, 1] }
    : axisBounds(
        lineOnPrimary ? Math.max(barMax, ...lineValues) : barMax,
        valueAxis,
        lineOnPrimary ? Math.min(barMin, ...lineValues) : barMin,
      )
  const span = bounds.max - bounds.min
  const norm = (value: number): number => Math.max(0, Math.min(1, (value - bounds.min) / span))
  // Bars grow from the zero line, not the plot floor, once the axis dips below 0.
  const zero = norm(0)
  // Stacked segments share the category slot; each value scales against the
  // axis maximum (percentStacked normalizes per category first).
  const segment = (seriesIndex: number, index: number): number => {
    const value = Math.max(0, seriesList[seriesIndex]?.values[index] ?? 0)
    if (!isPercent) return Math.min(1, value / (bounds.max || 1))
    const total = categoryTotal(index)
    return total === 0 ? 0 : value / total
  }
  const axisNumberFormat = isPercent ? '0%' : (valueAxis?.numFmt ?? seriesList[0]?.numberFormat)
  // Excel gap width: the space between category groups, in % of one bar.
  const gap = (gapWidthPct ?? 150) / 100
  const pickBar = onElement
    ? (event: React.MouseEvent, seriesIndex: number, pointIndex: number): void => {
        event.stopPropagation()
        onElement(narrowSelection(selectedEl ?? null, chartIndex(seriesIndex), pointIndex))
      }
    : undefined
  const barStroke = (seriesIndex: number, pointIndex: number): Record<string, string> =>
    isSelectedPoint(selectedEl, chartIndex(seriesIndex), pointIndex)
      ? { stroke: '#107C41', strokeWidth: '2' }
      : {}
  const selectCategoryAxis = onElement
    ? (event: React.MouseEvent): void => {
        event.stopPropagation()
        onElement({ kind: 'category-axis' })
      }
    : undefined
  const selectValueAxis = onElement
    ? (event: React.MouseEvent): void => {
        event.stopPropagation()
        onElement({ kind: 'value-axis' })
      }
    : undefined
  // Excel resolves labels per series (own dLbls, else the plot's), and a
  // per-point showVal/delete overrides both; a layout-only dLbl inherits.
  // Single-series charts with no dLbls at all keep the auto labels.
  const autoLabels =
    !isStacked && dataLabels === undefined && visibleCount <= 12 && seriesList.length === 1
  const pointLabel = (series: ChartSeries, index: number): PointLabel | undefined =>
    series.pointLabels?.find((label) => label.index === index)
  const showsLabel = (series: ChartSeries, index: number): boolean =>
    pointLabel(series, index)?.showVal ??
    (autoLabels || (series.dataLabels ?? dataLabels) === 'value')
  // manualLayout x/y are fractions of the chart space (the 600x320 viewBox).
  const labelOffset = (series: ChartSeries, index: number): { dx: number; dy: number } => {
    const label = pointLabel(series, index)
    return { dx: (label?.offsetX ?? 0) * 600, dy: (label?.offsetY ?? 0) * 320 }
  }
  // Stacked segments start where the previous series' segments end.
  const stackBase = (seriesIndex: number, index: number): number =>
    seriesList.slice(0, seriesIndex).reduce((sum, _, before) => sum + segment(before, index), 0)

  if (isHorizontal) {
    const rowHeight = 270 / visibleCount
    const rowGroups = visibleCategoryGroups(
      categoryHidden ? undefined : primary.categoryGroups,
      visibleCount,
    )
    const unit = chartTextUnit(textBox, 600, 320)
    const catFont = (categoryAxis?.labelSize ?? AXIS_LABEL_PT) * unit
    const charUnits = CHAR_EM * catFont
    // Excel gives the category labels the width they need up to about a
    // third of the chart, wrapping to two lines past that; the plot takes
    // the rest. A rotated category title and the outer tier column (62
    // units) sit left of the labels.
    const labelLeft = (axisTitles?.category ? 26 : 6) + (rowGroups.length > 0 ? 62 : 0)
    const widest = categoryHidden
      ? 0
      : Math.max(0, ...Array.from({ length: visibleCount }, (_, i) => categories[i]?.length ?? 0)) *
        charUnits
    const labelWidth = Math.min(200, Math.max(30, widest + 2))
    const labelRight = labelLeft + labelWidth
    const plotLeft = labelRight + 10
    const plotWidth = 548 - plotLeft
    const canWrap = rowHeight >= 2.3 * catFont
    const rowLabelLines = (index: number): readonly string[] => {
      const label = categories[index] ?? String(index + 1)
      const budget = Math.max(3, Math.floor(labelWidth / charUnits))
      const lines = canWrap ? categoryTickLines(label, labelWidth, charUnits) : [label]
      return lines.map((line) => truncateLabel(line, budget))
    }
    const scaledTick = (tick: number): number =>
      valueAxis?.displayUnit ? tick / valueAxis.displayUnit : tick
    const textStyle = chartTextStyle(unit, categoryAxis, valueAxis)
    const barHeight = Math.max(
      3,
      isStacked ? rowHeight / (1 + gap) : rowHeight / (seriesList.length + gap),
    )
    // The value scale sits along the bottom unless c:axPos puts it on top
    // (Excel's placement for maxMin categories crossing at autoZero); the
    // tick gutter and the value title move with it.
    const valueAxisTop = valueAxis?.position === 't'
    const plotTop = valueAxisTop ? (axisTitles?.value ? 40 : 26) : 12
    const plotBottom = plotTop + 274
    const rowsTop = plotTop + 2
    // Excel draws bar-chart categories bottom-up and series 0 nearest the
    // category axis under the default minMax orientation; maxMin flips both
    // back to top-down reading order.
    const rowSlot = (index: number): number => (categoryReversed ? index : visibleCount - 1 - index)
    const seriesSlot = (seriesIndex: number): number =>
      categoryReversed ? seriesIndex : seriesList.length - 1 - seriesIndex
    const groupTop = (index: number): number =>
      rowsTop +
      rowHeight * rowSlot(index) +
      (rowHeight - barHeight * (isStacked ? 1 : seriesList.length)) / 2
    const plotX = (value: number): number =>
      plotLeft +
      Math.max(0, Math.min(1, (value - bounds.min) / (bounds.max - bounds.min || 1))) * plotWidth
    const labelFor = (
      series: ChartSeries,
      seriesIndex: number,
      index: number,
      y: number,
    ): React.JSX.Element | null => {
      const value = series.values[index] ?? 0
      const { dx, dy } = labelOffset(series, index)
      // Stacked segments center their label (Excel's default); the white
      // fill only helps while the label still sits on its own bar.
      const centered = isStacked || dataLabelPosition === 'center'
      const inside = centered || dataLabelPosition === 'inside-end'
      const width = (isStacked ? segment(seriesIndex, index) : norm(value) - zero) * plotWidth
      const start = plotLeft + (isStacked ? stackBase(seriesIndex, index) : zero) * plotWidth
      const below = width < 0
      const x = centered
        ? start + width / 2
        : dataLabelPosition === 'inside-end'
          ? start + width + (below ? 4 : -4)
          : start + width + (below ? -6 : 6)
      const anchor = (dataLabelPosition === 'inside-end') !== below ? 'end' : 'start'
      return (
        <text
          x={x + dx}
          y={y + dy}
          textAnchor={centered ? 'middle' : anchor}
          className="data-label"
          {...(inside && !isStacked && dx === 0 && dy === 0 ? { fill: '#fff' } : {})}
          style={labelTextStyle(dataLabelStyle)}
        >
          {formatLabelValue(value, dataLabelFormat, series.numberFormat)}
        </text>
      )
    }
    return (
      <svg className="chart-svg" viewBox="0 0 600 320" role="img" style={textStyle}>
        {plotAreaFill && (
          <rect x={plotLeft} y={plotTop} width={548 - plotLeft} height="274" fill={plotAreaFill} />
        )}
        {bounds.ticks.map((tick, index) => (
          <g key={`vt-${index}`} onClick={selectValueAxis}>
            {gridlines !== false && (
              <line
                x1={plotX(tick)}
                y1={plotTop}
                x2={plotX(tick)}
                y2={plotBottom}
                stroke="#e3e3e3"
                strokeWidth="1"
              />
            )}
            {valueAxis?.hidden !== true && (
              <text
                x={plotX(tick)}
                y={valueAxisTop ? plotTop - 6 : plotBottom + 12}
                textAnchor="middle"
                className="axis-label"
              >
                {formatAxisValue(scaledTick(tick), axisNumberFormat)}
              </text>
            )}
          </g>
        ))}
        {Array.from({ length: visibleCount }, (_, index) => {
          let cursor = 0
          return (
            <g key={`${categories[index] ?? index}-${index}`}>
              {!categoryHidden && (
                <text
                  x={labelRight}
                  y={
                    rowsTop +
                    rowHeight * (rowSlot(index) + 0.5) +
                    catFont * (rowLabelLines(index).length > 1 ? -0.25 : 0.35)
                  }
                  textAnchor="end"
                  onClick={selectCategoryAxis}
                >
                  {rowLabelLines(index).map((line, lineIndex) => (
                    <tspan key={lineIndex} x={labelRight} dy={lineIndex === 0 ? 0 : catFont * 1.15}>
                      {line}
                    </tspan>
                  ))}
                </text>
              )}
              {seriesList.map((series, seriesIndex) => {
                const share = isStacked
                  ? segment(seriesIndex, index)
                  : norm(series.values[index] ?? 0) - zero
                const x = plotLeft + (isStacked ? cursor : Math.min(zero, zero + share)) * plotWidth
                if (isStacked) cursor += share
                return (
                  <rect
                    key={seriesIndex}
                    x={x}
                    y={groupTop(index) + (isStacked ? 0 : barHeight * seriesSlot(seriesIndex))}
                    width={Math.abs(share) * plotWidth}
                    height={barHeight}
                    fill={seriesColor(series, seriesIndex)}
                    {...barStroke(seriesIndex, index)}
                    onClick={pickBar ? (event) => pickBar(event, seriesIndex, index) : undefined}
                  />
                )
              })}
              {seriesList.map(
                (series, seriesIndex) =>
                  showsLabel(series, index) && (
                    <Fragment key={`lbl-${seriesIndex}`}>
                      {labelFor(
                        series,
                        seriesIndex,
                        index,
                        groupTop(index) +
                          (isStacked ? 0 : barHeight * seriesSlot(seriesIndex)) +
                          barHeight / 2 +
                          3,
                      )}
                    </Fragment>
                  ),
              )}
            </g>
          )
        })}
        {/* Outer multiLvlStrCache level: a rotated label column left of the
            category labels with separator lines at the group edges. */}
        {rowGroups.map((group, groupIndex) => {
          const first = rowSlot(group.start)
          const last = rowSlot(group.end - 1)
          const yTop = rowsTop + rowHeight * Math.min(first, last)
          const yBottom = rowsTop + rowHeight * (Math.max(first, last) + 1)
          const cy = (yTop + yBottom) / 2
          return (
            <g key={`grp-${groupIndex}`} onClick={selectCategoryAxis}>
              <line
                x1={labelLeft - 62}
                y1={yTop}
                x2={plotLeft}
                y2={yTop}
                stroke="#d9d9d9"
                strokeWidth="1"
              />
              <line
                x1={labelLeft - 62}
                y1={yBottom}
                x2={plotLeft}
                y2={yBottom}
                stroke="#d9d9d9"
                strokeWidth="1"
              />
              <text
                x={labelLeft - 54}
                y={cy}
                transform={`rotate(-90 ${labelLeft - 54} ${cy})`}
                textAnchor="middle"
              >
                {truncateLabel(group.label, Math.max(3, Math.floor((yBottom - yTop) / charUnits)))}
              </text>
            </g>
          )
        })}
        <TruncationNote shown={visibleCount} total={pointCount} />
        <AxisTitleTexts
          bottom={axisTitles?.value}
          left={axisTitles?.category}
          bottomY={valueAxisTop ? 12 : 317}
          bottomUnits={valueAxis?.displayUnitLabel}
          titleFont={(categoryAxis?.titleSize ?? AXIS_TITLE_PT) * unit}
        />
      </svg>
    )
  }

  const columnWidth = 480 / visibleCount
  const barWidth = Math.max(
    2,
    isStacked ? columnWidth / (1 + gap) : columnWidth / (seriesList.length + gap),
  )
  const groupWidth = barWidth * (isStacked ? 1 : seriesList.length)
  // maxMin mirrors the column order (and the series order inside a group).
  const catSlot = (index: number): number => (categoryReversed ? visibleCount - 1 - index : index)
  const seriesSlot = (seriesIndex: number): number =>
    categoryReversed ? seriesList.length - 1 - seriesIndex : seriesIndex
  const groupLeft = (index: number): number =>
    62 + columnWidth * catSlot(index) + (columnWidth - groupWidth) / 2
  // Combo lines ride the secondary value axis when the file has one; a
  // single axis pair means they share the primary scale (Excel never
  // invents a right-hand axis).
  const lineScale =
    lineSeriesList.length > 0
      ? secondaryAxis
        ? valueAxisScale(Math.max(...lineValues, 0), secondaryAxis, Math.min(...lineValues))
        : bounds
      : undefined
  const lineSpan = lineScale ? lineScale.max - lineScale.min || 1 : 1
  const comboX = (index: number): number => groupLeft(index) + groupWidth / 2
  const comboLines = lineSeriesList.map((series, lineIndex) => {
    // Palette slots continue after the bar series, as Excel cycles accents.
    const paletteIndex = seriesList.length + lineIndex
    const stroke = lineStroke(series, paletteIndex)
    const y = (index: number): number =>
      280 -
      Math.max(0, Math.min(1, ((series.values[index] ?? 0) - (lineScale?.min ?? 0)) / lineSpan)) *
        240
    const symbol =
      series.marker === 'none'
        ? null
        : lineMarkers !== true && series.marker === undefined
          ? null
          : series.marker !== undefined && series.marker !== 'auto'
            ? series.marker
            : (AUTO_MARKER_SYMBOLS[paletteIndex % AUTO_MARKER_SYMBOLS.length] ?? 'circle')
    return {
      stroke,
      width: series.lineWidth ?? 3,
      markerColor: stroke ?? seriesColor(series, paletteIndex),
      segments: lineSegments(
        Math.min(series.values.length, visibleCount),
        series.blanks,
        dispBlanksAs,
      ),
      y,
      symbol,
    }
  })
  const columnGroups = visibleCategoryGroups(
    categoryHidden ? undefined : primary.categoryGroups,
    visibleCount,
  )
  // The group band occupies y 284-314; a bottom axis title moves below it
  // on an extended canvas instead of overprinting.
  const shiftBottomTitle = columnGroups.length > 0 && Boolean(axisTitles?.category)
  const layoutFor = (viewBoxHeight: number): { unit: number; extra: number } =>
    valueAxisLayout(
      textBox,
      viewBoxHeight,
      valueLabelEm(
        bounds.ticks,
        valueAxis?.displayUnit,
        axisNumberFormat,
        valueAxis?.hidden === true,
      ),
      valueAxis?.labelSize ?? AXIS_LABEL_PT,
      axisSideReservePt(
        axisTitles?.value,
        valueAxis?.displayUnitLabel,
        valueAxis?.titleSize ?? AXIS_TITLE_PT,
        valueAxis?.labelSize ?? AXIS_LABEL_PT,
      ),
    )
  const catPt = categoryAxis?.labelSize ?? AXIS_LABEL_PT
  const bottom = bottomAxisLayout(
    categoryTicksWrap(categories, visibleCount, columnWidth, CHAR_EM * catPt * layoutFor(320).unit),
    Boolean(axisTitles?.category),
    shiftBottomTitle,
    catPt * layoutFor(320).unit,
    (categoryAxis?.titleSize ?? AXIS_TITLE_PT) * layoutFor(320).unit,
  )
  const viewBoxHeight = bottom.viewBoxHeight
  const { unit, extra } = layoutFor(viewBoxHeight)
  const catFont = catPt * unit
  const tickStride = categoryTickStride(categories, visibleCount, columnWidth, CHAR_EM * catFont)
  const textStyle = chartTextStyle(unit, categoryAxis, valueAxis)
  return (
    <svg
      className="chart-svg"
      viewBox={`${-extra} 0 ${600 + extra} ${viewBoxHeight}`}
      role="img"
      style={textStyle}
    >
      {plotAreaFill && <rect x="58" y="40" width="522" height="240" fill={plotAreaFill} />}
      <VerticalAxis
        minimum={bounds.min}
        maximum={bounds.max}
        ticks={isPercent ? undefined : bounds.ticks}
        numberFormat={axisNumberFormat}
        showGridlines={gridlines !== false}
        hideLabels={valueAxis?.hidden === true}
        displayUnit={valueAxis?.displayUnit}
        onSelect={selectValueAxis}
      />
      {Array.from({ length: visibleCount }, (_, index) => {
        let cursor = 0
        return (
          <g key={`${categories[index] ?? index}-${index}`}>
            {seriesList.map((series, seriesIndex) => {
              const share = isStacked
                ? segment(seriesIndex, index)
                : norm(series.values[index] ?? 0) - zero
              const height = Math.abs(share) * 240
              const y = 280 - (isStacked ? cursor + share : Math.max(zero, zero + share)) * 240
              if (isStacked) cursor += share
              return (
                <rect
                  key={seriesIndex}
                  x={groupLeft(index) + (isStacked ? 0 : barWidth * seriesSlot(seriesIndex))}
                  y={y}
                  width={barWidth}
                  height={height}
                  fill={seriesColor(series, seriesIndex)}
                  {...barStroke(seriesIndex, index)}
                  onClick={pickBar ? (event) => pickBar(event, seriesIndex, index) : undefined}
                />
              )
            })}
            {seriesList.map((series, seriesIndex) => {
              if (!showsLabel(series, index)) return null
              const { dx, dy } = labelOffset(series, index)
              const inside =
                isStacked || dataLabelPosition === 'center' || dataLabelPosition === 'inside-end'
              const share = isStacked
                ? segment(seriesIndex, index)
                : norm(series.values[index] ?? 0) - zero
              const below = share < 0
              const barEnd = 280 - (isStacked ? 0 : zero + share) * 240
              const y = isStacked
                ? 283 - (stackBase(seriesIndex, index) + share / 2) * 240
                : dataLabelPosition === 'center'
                  ? barEnd + 3 + share * 120
                  : dataLabelPosition === 'inside-end'
                    ? barEnd + (below ? -4 : 12)
                    : barEnd + (below ? 12 : -8)
              const x =
                groupLeft(index) +
                (isStacked ? 0 : barWidth * seriesSlot(seriesIndex)) +
                barWidth / 2
              return (
                <text
                  key={`lbl-${seriesIndex}`}
                  x={x + dx}
                  y={y + dy}
                  textAnchor="middle"
                  className="data-label"
                  {...(inside && !isStacked && dx === 0 && dy === 0 ? { fill: '#fff' } : {})}
                  style={labelTextStyle(dataLabelStyle)}
                >
                  {formatLabelValue(
                    series.values[index] ?? 0,
                    dataLabelFormat,
                    series.numberFormat,
                  )}
                </text>
              )
            })}
            {!categoryHidden && (
              <CategoryTick
                x={62 + columnWidth * catSlot(index) + columnWidth / 2}
                label={categories[index] ?? String(index + 1)}
                slotWidth={columnWidth}
                stride={tickStride}
                index={index}
                fontUnits={catFont}
                onClick={selectCategoryAxis}
              />
            )}
          </g>
        )
      })}
      {!isStacked &&
        seriesList.map((series, seriesIndex) =>
          series.trendline === 'linear' ? (
            <TrendLine
              key={`trend-${seriesIndex}`}
              values={series.values.slice(0, visibleCount)}
              minimum={bounds.min}
              maximum={bounds.max}
              color={seriesColor(series, seriesIndex)}
              xFor={(index) => groupLeft(index) + groupWidth / 2}
            />
          ) : null,
        )}
      {comboLines.map(({ stroke, symbol, ...line }, lineIndex) => (
        <g key={`combo-${lineIndex}`}>
          {stroke !== null &&
            line.segments.map((segment, segmentIndex) => (
              <polyline
                key={segmentIndex}
                points={segment.map((index) => `${comboX(index)},${line.y(index)}`).join(' ')}
                fill="none"
                stroke={stroke}
                strokeWidth={line.width}
              />
            ))}
          {symbol !== null &&
            line.segments
              .flat()
              .map((index) => (
                <MarkerGlyph
                  key={`marker-${index}`}
                  x={comboX(index)}
                  y={line.y(index)}
                  symbol={symbol}
                  color={line.markerColor}
                />
              ))}
        </g>
      ))}
      <CategoryGroupBand
        spans={columnGroups.map((group) => {
          const first = catSlot(group.start)
          const last = catSlot(group.end - 1)
          return {
            label: group.label,
            xStart: 62 + columnWidth * Math.min(first, last),
            xEnd: 62 + columnWidth * (Math.max(first, last) + 1),
          }
        })}
        charUnits={CHAR_EM * catFont}
        onClick={selectCategoryAxis}
      />
      {lineScale &&
        secondaryAxis !== undefined &&
        secondaryAxis.hidden !== true &&
        lineScale.ticks.map((tick, index) => (
          <text
            key={index}
            x="596"
            y={284 - ((tick - lineScale.min) / lineSpan) * 240}
            textAnchor="end"
            className="axis-label"
          >
            {formatAxisValue(
              secondaryAxis.displayUnit ? tick / secondaryAxis.displayUnit : tick,
              secondaryAxis.numFmt ?? lineSeriesList[0]?.numberFormat,
            )}
          </text>
        ))}
      <TruncationNote shown={visibleCount} total={pointCount} />
      <AxisTitleTexts
        bottom={axisTitles?.category}
        left={axisTitles?.value}
        bottomY={bottom.bottomY}
        leftEdge={-extra}
        leftUnits={valueAxis?.displayUnitLabel}
        titleFont={(valueAxis?.titleSize ?? AXIS_TITLE_PT) * unit}
        labelFont={(valueAxis?.labelSize ?? AXIS_LABEL_PT) * unit}
      />
    </svg>
  )
}

/// Least-squares regression over (index, value), rendered as a dashed line.
function TrendLine({
  values,
  minimum = 0,
  maximum,
  color,
  xFor,
}: {
  readonly values: readonly number[]
  readonly minimum?: number
  readonly maximum: number
  readonly color: string
  readonly xFor: (index: number) => number
}): React.JSX.Element | null {
  const count = values.length
  if (count < 2) return null
  const meanX = (count - 1) / 2
  const meanY = values.reduce((sum, value) => sum + value, 0) / count
  let numerator = 0
  let denominator = 0
  for (let index = 0; index < count; index += 1) {
    numerator += (index - meanX) * ((values[index] ?? 0) - meanY)
    denominator += (index - meanX) ** 2
  }
  const slope = denominator === 0 ? 0 : numerator / denominator
  const intercept = meanY - slope * meanX
  const span = maximum - minimum || 1
  const yFor = (index: number): number =>
    280 - Math.max(0, Math.min(1, (intercept + slope * index - minimum) / span)) * 240
  return (
    <line
      x1={xFor(0)}
      y1={yFor(0)}
      x2={xFor(count - 1)}
      y2={yFor(count - 1)}
      stroke={color}
      strokeWidth="2"
      strokeDasharray="6 4"
      opacity="0.85"
    />
  )
}

/// Excel's automatic marker cycle when the plot flag is on but the series
/// sets no symbol (diamond, square, triangle, then repeat).
const AUTO_MARKER_SYMBOLS = ['diamond', 'square', 'triangle'] as const

function MarkerGlyph({
  x,
  y,
  symbol,
  color,
}: {
  readonly x: number
  readonly y: number
  readonly symbol: string
  readonly color: string
}): React.JSX.Element {
  const r = 4
  if (symbol === 'diamond') {
    return (
      <polygon points={`${x},${y - r} ${x + r},${y} ${x},${y + r} ${x - r},${y}`} fill={color} />
    )
  }
  if (symbol === 'square') {
    return <rect x={x - r} y={y - r} width={2 * r} height={2 * r} fill={color} />
  }
  if (symbol === 'triangle') {
    return <polygon points={`${x},${y - r} ${x + r},${y + r} ${x - r},${y + r}`} fill={color} />
  }
  return <circle cx={x} cy={y} r={r} fill={color} />
}

/// Series line stroke: an explicit spPr/a:ln color wins over the fill/accent
/// default; null for an explicit noFill line.
export function lineStroke(
  series: SeriesLike & { readonly lineColor?: string | undefined },
  index: number,
): string | null {
  if (series.lineColor === 'none') return null
  return series.lineColor ?? seriesColor(series, index)
}

/// Index runs a line plots as one stroke. Blank cells (c:dispBlanksAs)
/// break the line on 'gap', are bridged on 'span', and plot their 0 filler
/// on 'zero' — the OOXML default when the element is absent.
export function lineSegments(
  count: number,
  blanks: readonly number[] | undefined,
  mode: string | undefined,
): number[][] {
  const indices = Array.from({ length: count }, (_, index) => index)
  if ((mode !== 'gap' && mode !== 'span') || blanks === undefined || blanks.length === 0) {
    return count > 0 ? [indices] : []
  }
  const blankSet = new Set(blanks)
  if (mode === 'span') {
    const kept = indices.filter((index) => !blankSet.has(index))
    return kept.length > 0 ? [kept] : []
  }
  const segments: number[][] = []
  let current: number[] = []
  for (const index of indices) {
    if (blankSet.has(index)) {
      if (current.length > 0) segments.push(current)
      current = []
    } else {
      current.push(index)
    }
  }
  if (current.length > 0) segments.push(current)
  return segments
}

/// Whether the legend mirrors stroke colors: only when a line-drawing
/// component wins the render cascade (pie → radar → scatter → area → line →
/// bar) — a chart that also carries an areaChart/scatterChart paints fills.
export function legendUsesLineSwatches(types: readonly string[]): boolean {
  if (types.includes('pieChart') || types.includes('doughnutChart')) return false
  const noBar = !types.includes('barChart')
  if (types.includes('radarChart') && noBar) return true
  if (types.includes('scatterChart')) return false
  if (types.includes('areaChart') && noBar) return false
  return types.includes('lineChart') && noBar
}

/// Bar+line combo: series whose plot group is lineChart draw as lines.
/// Untagged series (AI-built charts, older snapshots) keep the legacy rule of
/// the last series being the line.
export function comboLineIndices(
  seriesList: readonly { readonly plot?: string | undefined }[],
): Set<number> {
  if (seriesList.every((series) => series.plot === undefined)) {
    return new Set(seriesList.length > 1 ? [seriesList.length - 1] : [])
  }
  return new Set(
    seriesList.flatMap((series, index) => (series.plot === 'lineChart' ? [index] : [])),
  )
}

export function comboBarIndices(
  seriesList: readonly unknown[],
  lineIndices: ReadonlySet<number>,
): number[] {
  return seriesList.flatMap((_, index) => (lineIndices.has(index) ? [] : [index]))
}

/// Excel often caches the category list on the first plot group only, so a
/// line-first combo must not read categories from the first bar series.
export function comboCategorySeries<
  T extends { readonly categories: readonly unknown[]; readonly values: readonly unknown[] },
>(seriesList: readonly T[]): T | undefined {
  return (
    seriesList.find((series) => series.categories.length > 0) ??
    seriesList.reduce<T | undefined>(
      (best, series) => (best && best.values.length >= series.values.length ? best : series),
      undefined,
    )
  )
}

/// Legend swatch color: line-family charts mirror the drawn stroke (an
/// explicit noFill line keeps the series color); everything else keeps the
/// fill resolution.
export function legendSwatchColor(
  series: SeriesLike & { readonly lineColor?: string | undefined },
  index: number,
  lineSwatches: boolean,
): string {
  if (!lineSwatches) return seriesColor(series, index)
  return lineStroke(series, index) ?? seriesColor(series, index)
}

function linePoints(values: readonly number[], maximum: number, minimum = 0): string {
  const count = Math.max(1, values.length - 1)
  const span = maximum - minimum || 1
  return values
    .map(
      (value, index) =>
        `${60 + (index / count) * 500},${280 - Math.max(0, Math.min(1, (value - minimum) / span)) * 240}`,
    )
    .join(' ')
}

/// 5-tick nice ceiling so the top gridline sits at or above the data maximum.
function niceAxisMaximum(maximum: number): number {
  if (maximum <= 0) return 1
  const magnitude = 10 ** Math.floor(Math.log10(maximum))
  const normalized = maximum / magnitude
  const nice =
    normalized <= 1 ? 1 : normalized <= 2 ? 2 : normalized <= 2.5 ? 2.5 : normalized <= 5 ? 5 : 10
  return nice * magnitude
}

type ChartLabelPosition = ChartMetadata['dataLabelPosition']

/// Minimal formatCode support for data labels (percent / thousands / fixed
/// decimals); anything fancier falls back to the axis heuristics.
/// Pie value labels honor the full source number format (currency symbols,
/// accounting padding) the way Excel's sourceLinked labels do.
function formatPieValue(value: number, formatCode: string | undefined): string {
  if (formatCode && formatCode !== 'General') {
    try {
      return numfmt.format(formatCode, value, { throws: false }).trim()
    } catch {
      // Unparseable format: fall through to the plain rendering.
    }
  }
  return formatLabelValue(value, formatCode, undefined)
}

/// Excel's showPercent default format is "0%" — integer rounding — unless
/// the dLbls carry their own percent numFmt.
export function formatPiePercent(share: number, formatCode: string | undefined): string {
  if (formatCode !== undefined && formatCode.includes('%')) {
    return formatLabelValue(share, formatCode, undefined)
  }
  return `${Math.round(share * 100)}%`
}

/// Explicit c:dLbls/c:txPr font as inline style: the stylesheet's label
/// rules would beat presentation attributes. Size is in points, scaled by
/// the chart's text unit like every other label.
function labelTextStyle(style: ChartMetadata['dataLabelStyle']): React.CSSProperties | undefined {
  if (!style) return undefined
  return {
    ...(style.color ? { fill: style.color } : {}),
    ...(style.size ? { fontSize: `calc(${style.size} * var(--chart-pt, 1px))` } : {}),
    ...(style.bold ? { fontWeight: 700 } : {}),
  }
}

function formatLabelValue(
  value: number,
  formatCode: string | undefined,
  numberFormat: string | undefined,
): string {
  if (!formatCode || formatCode === 'General') return formatAxisValue(value, numberFormat)
  // Full format codes (currency symbols, literals) go through numfmt like
  // the axis; the digit-count shorthand below only backs up codes it rejects.
  try {
    const text = numfmt.format(formatCode, value, { throws: false })
    if (typeof text === 'string' && text.trim() !== '') return text.trim()
  } catch {
    // fall through to the shorthand rendering
  }
  const decimals = /0\.(0+)/.exec(formatCode)?.[1]?.length ?? 0
  if (formatCode.includes('%')) return `${(value * 100).toFixed(decimals)}%`
  const fixed = value.toFixed(decimals)
  if (!formatCode.includes(',')) return fixed
  const [whole, fraction] = fixed.split('.')
  const grouped = Number(whole).toLocaleString('en-US')
  return fraction ? `${grouped}.${fraction}` : grouped
}

/// Corner note when the drawing shows fewer categories than the data has.
function TruncationNote({
  shown,
  total,
}: {
  readonly shown: number
  readonly total: number
}): React.JSX.Element | null {
  if (total <= shown) return null
  return (
    <text x="580" y="14" textAnchor="end" className="axis-label" opacity="0.75">
      {t('appTruncationNote', { shown, total })}
    </text>
  )
}

/// Approximate glyph width of the small axis font in the 640-unit viewBox.
const AXIS_LABEL_CHAR_UNITS = 5.4

/// Excel wraps a category label at spaces when its slot is narrow ("KW 01"
/// stacks as KW / 01); greedy two-line split, single line when it fits or
/// has no break point.
export function categoryTickLines(
  label: string,
  slotWidth: number,
  charUnits = AXIS_LABEL_CHAR_UNITS,
): readonly string[] {
  const fits = (text: string): boolean => text.length * charUnits <= slotWidth
  if (fits(label) || !label.includes(' ')) return [label]
  const words = label.split(/\s+/)
  let first = words[0] ?? ''
  let index = 1
  while (index < words.length && fits(`${first} ${words[index]}`)) {
    first = `${first} ${words[index]}`
    index += 1
  }
  const rest = words.slice(index).join(' ')
  return rest ? [first, rest] : [first]
}

/// Excel thins overlapping category labels to every n-th instead of
/// overprinting: when even the wrapped label is wider than its slot, only
/// every ceil(width/slot)-th category keeps its label.
export function categoryTickStride(
  labels: readonly string[],
  count: number,
  slotWidth: number,
  charUnits = AXIS_LABEL_CHAR_UNITS,
): number {
  if (!(slotWidth > 0)) return 1
  let widest = 0
  for (let index = 0; index < count; index += 1) {
    for (const line of categoryTickLines(
      labels[index] ?? String(index + 1),
      slotWidth,
      charUnits,
    )) {
      widest = Math.max(widest, Math.min(line.length, 16) * charUnits)
    }
  }
  return Math.max(1, Math.ceil((widest + 2) / slotWidth))
}

/// Whether any drawn category tick needs its second line.
function categoryTicksWrap(
  labels: readonly string[],
  count: number,
  slotWidth: number,
  charUnits: number,
): boolean {
  const stride = categoryTickStride(labels, count, slotWidth, charUnits)
  for (let index = 0; index < count; index += stride) {
    if (
      categoryTickLines(labels[index] ?? String(index + 1), slotWidth * stride, charUnits).length >
      1
    ) {
      return true
    }
  }
  return false
}

function CategoryTick({
  x,
  label,
  slotWidth,
  stride,
  index,
  fontUnits = 10,
  onClick,
}: {
  readonly x: number
  readonly label: string
  readonly slotWidth: number
  readonly stride: number
  readonly index: number
  /// Tick font size in viewBox units (glyph budget and line pitch).
  readonly fontUnits?: number
  readonly onClick?: ((event: React.MouseEvent) => void) | undefined
}): React.JSX.Element | null {
  if (index % stride !== 0) return null
  const charUnits = CHAR_EM * fontUnits
  const budget = Math.max(5, Math.min(16, Math.floor((slotWidth * stride) / charUnits)))
  const lines = categoryTickLines(label, slotWidth * stride, charUnits)
  // Two-line ticks start higher; the plot floor is 280, so line one at 292
  // stays below the bars. bottomAxisLayout moves the axis title down when
  // the second line would reach it.
  return (
    <text x={x} y={lines.length === 1 ? 298 : 292} textAnchor="middle" onClick={onClick}>
      {lines.length === 1 ? (
        truncateLabel(lines[0] ?? '', budget)
      ) : (
        <>
          <tspan x={x} dy="0">
            {truncateLabel(lines[0] ?? '', budget)}
          </tspan>
          <tspan x={x} dy={fontUnits * 1.1}>
            {truncateLabel(lines[1] ?? '', budget)}
          </tspan>
        </>
      )}
    </text>
  )
}

/// Clamp parsed outer-level category spans to the drawn category count.
function visibleCategoryGroups(
  groups: ChartSeries['categoryGroups'],
  visibleCount: number,
): { label: string; start: number; end: number }[] {
  return (groups ?? [])
    .map((group) => ({ ...group, end: Math.min(group.end, visibleCount) }))
    .filter((group) => group.start < group.end && group.start < visibleCount)
}

/// Outer multiLvlStrCache level under the tick labels: each group label
/// centered under its span with separator ticks at the span edges (Excel's
/// grouped category axis).
function CategoryGroupBand({
  spans,
  charUnits,
  onClick,
}: {
  readonly spans: readonly { label: string; xStart: number; xEnd: number }[]
  /// Glyph advance of the point-sized group label font, in viewBox units.
  readonly charUnits: number
  readonly onClick?: ((event: React.MouseEvent) => void) | undefined
}): React.JSX.Element {
  return (
    <g onClick={onClick}>
      {spans.map((span, index) => (
        <Fragment key={index}>
          <line
            x1={span.xStart}
            y1="284"
            x2={span.xStart}
            y2="314"
            stroke="#d9d9d9"
            strokeWidth="1"
          />
          <line x1={span.xEnd} y1="284" x2={span.xEnd} y2="314" stroke="#d9d9d9" strokeWidth="1" />
          <text x={(span.xStart + span.xEnd) / 2} y="312" textAnchor="middle">
            {truncateLabel(
              span.label,
              Math.max(3, Math.floor((span.xEnd - span.xStart) / charUnits)),
            )}
          </text>
        </Fragment>
      ))}
    </g>
  )
}

export function formatAxisValue(value: number, numberFormat: string | undefined): string {
  if (numberFormat && numberFormat !== 'General' && !numberFormat.includes('%')) {
    try {
      const text = numfmt.format(numberFormat, value, { throws: false })
      if (typeof text === 'string' && text !== '') return text
    } catch {
      // fall through to the plain rendering
    }
  }
  if (numberFormat?.includes('%')) return `${Math.round(value * 100)}%`
  // Excel prints full numbers on value axes; abbreviating disagreed with the data labels.
  const clean = Number(value.toPrecision(12))
  return clean.toLocaleString('en-US', { maximumFractionDigits: 4 })
}

function VerticalAxis({
  minimum = 0,
  maximum,
  ticks,
  numberFormat,
  showGridlines = true,
  hideLabels = false,
  displayUnit,
  onSelect,
}: {
  readonly minimum?: number
  readonly maximum: number
  readonly ticks?: readonly number[] | undefined
  readonly numberFormat: string | undefined
  readonly showGridlines?: boolean
  /// c:delete on the axis: gridlines survive, the scale labels do not.
  readonly hideLabels?: boolean
  /// c:dispUnits divisor applied to the shown tick values.
  readonly displayUnit?: number | undefined
  readonly onSelect?: ((event: React.MouseEvent) => void) | undefined
}): React.JSX.Element {
  const span = maximum - minimum || 1
  const list =
    ticks && ticks.length >= 2
      ? ticks
      : [0, 0.25, 0.5, 0.75, 1].map((fraction) => minimum + fraction * span)
  return (
    <g onClick={onSelect}>
      {list.map((tick, index) => {
        const y = 280 - ((tick - minimum) / span) * 240
        return (
          <g key={index}>
            {(showGridlines || index === 0) && (
              <line x1="58" y1={y} x2="580" y2={y} stroke="#e3e3e3" strokeWidth="1" />
            )}
            {!hideLabels && (
              <text x="54" y={y + 4} textAnchor="end" className="axis-label">
                {formatAxisValue(displayUnit ? tick / displayUnit : tick, numberFormat)}
              </text>
            )}
          </g>
        )
      })}
    </g>
  )
}

type ChartValueAxis =
  | ({
      min?: number | undefined
      max?: number | undefined
      majorUnit?: number | undefined
      numFmt?: string | undefined
      hidden?: boolean | undefined
      /// c:axPos side; a horizontal bar's value scale moves to the top on 't'.
      position?: 'l' | 'r' | 't' | 'b' | undefined
      displayUnit?: number | undefined
      displayUnitLabel?: string | undefined
    } & ChartAxisText)
  | undefined

/// Explicit axis bounds/unit win; otherwise an Excel-like auto scale.
/// `dataMin` is the plotted series minimum: stacked charts clamp their
/// segments at 0 and pass 0, but without it a below-zero series scales to
/// 0..1 and every bar draws at full height.
function axisBounds(
  dataMax: number,
  valueAxis: ChartValueAxis,
  dataMin = 0,
): { min: number; max: number; ticks: number[] } {
  return valueAxisScale(dataMax, valueAxis, dataMin)
}

export function LineChart({
  seriesList,
  axisTitles,
  dataLabels,
  grouping,
  gridlines,
  valueAxis,
  categoryFormat,
  categoryHidden = false,
  lineMarkers,
  dispBlanksAs,
  textBox,
  categoryAxis,
  plotAreaFill,
  onElement,
  selectedEl,
}: {
  readonly seriesList: readonly ChartSeries[]
  readonly axisTitles?: ChartAxisTitles
  readonly dataLabels?: ChartDataLabels
  readonly grouping?: ChartGrouping
  readonly gridlines?: boolean | undefined
  readonly valueAxis?: ChartValueAxis
  readonly categoryFormat?: string | undefined
  readonly categoryHidden?: boolean | undefined
  readonly lineMarkers?: boolean | undefined
  readonly dispBlanksAs?: ChartMetadata['dispBlanksAs']
  readonly textBox?: ChartTextBox | undefined
  readonly categoryAxis?: ChartAxisText | undefined
  readonly plotAreaFill?: string | undefined
} & ChartElementProps): React.JSX.Element {
  const primary = seriesList[0]
  if (!primary) return <></>
  const isStacked =
    (grouping === 'stacked' || grouping === 'percentStacked') && seriesList.length > 1
  const isPercent = grouping === 'percentStacked' && seriesList.length > 1
  // Running per-category totals; each line rides on the ones below it
  // (negatives clamp to 0, matching AreaChart).
  const stackTotals = isStacked
    ? seriesList.reduce<number[][]>((totals, series) => {
        const previous = totals[totals.length - 1] ?? primary.values.map(() => 0)
        totals.push(previous.map((base, index) => base + Math.max(0, series.values[index] ?? 0)))
        return totals
      }, [])
    : []
  const percentTotals = isPercent
    ? primary.values.map(
        (_, index) =>
          seriesList.reduce((sum, series) => sum + Math.max(0, series.values[index] ?? 0), 0) || 1,
      )
    : []
  const displayValues = (seriesIndex: number): number[] => {
    if (!isStacked) return [...(seriesList[seriesIndex]?.values ?? [])]
    const totals = stackTotals[seriesIndex] ?? []
    return isPercent ? totals.map((value, index) => value / (percentTotals[index] ?? 1)) : totals
  }
  const bounds = isPercent
    ? { min: 0, max: 1, ticks: [0, 0.25, 0.5, 0.75, 1] }
    : axisBounds(
        isStacked
          ? Math.max(...(stackTotals[stackTotals.length - 1] ?? [0]), 0)
          : Math.max(...seriesList.flatMap((series) => [...series.values]), 0),
        isStacked ? undefined : valueAxis,
        // Stack totals clamp negatives to 0; only the plain branch goes below.
        isStacked ? 0 : Math.min(...seriesList.flatMap((series) => [...series.values])),
      )
  const span = bounds.max - bounds.min
  const categories = primary.categories.map((value) => formatCategoryLabel(value, categoryFormat))
  const count = Math.max(1, primary.values.length - 1)
  // Stacked lines keep the 0 fillers (a blank contributes nothing to the
  // stack); only flat lines honor gap/span.
  const seriesSegments = (series: ChartSeries): number[][] =>
    lineSegments(series.values.length, isStacked ? undefined : series.blanks, dispBlanksAs)
  const isSkippedBlank = (series: ChartSeries, index: number): boolean =>
    !isStacked &&
    (dispBlanksAs === 'gap' || dispBlanksAs === 'span') &&
    (series.blanks?.includes(index) ?? false)
  const categoryGroups = visibleCategoryGroups(
    categoryHidden ? undefined : primary.categoryGroups,
    primary.values.length,
  )
  // The group band occupies y 284-314; a bottom axis title moves below it
  // on an extended canvas instead of overprinting.
  const shiftBottomTitle = categoryGroups.length > 0 && Boolean(axisTitles?.category)
  const lineNumberFormat = isPercent ? '0%' : (valueAxis?.numFmt ?? primary.numberFormat)
  const layoutFor = (viewBoxHeight: number): { unit: number; extra: number } =>
    valueAxisLayout(
      textBox,
      viewBoxHeight,
      valueLabelEm(
        bounds.ticks,
        valueAxis?.displayUnit,
        lineNumberFormat,
        valueAxis?.hidden === true,
      ),
      valueAxis?.labelSize ?? AXIS_LABEL_PT,
      axisSideReservePt(
        axisTitles?.value,
        valueAxis?.displayUnitLabel,
        valueAxis?.titleSize ?? AXIS_TITLE_PT,
        valueAxis?.labelSize ?? AXIS_LABEL_PT,
      ),
    )
  const catPt = categoryAxis?.labelSize ?? AXIS_LABEL_PT
  const slotWidth = 500 / Math.max(1, count)
  const bottom = bottomAxisLayout(
    categoryTicksWrap(
      categories,
      primary.values.length,
      slotWidth,
      CHAR_EM * catPt * layoutFor(320).unit,
    ),
    Boolean(axisTitles?.category),
    shiftBottomTitle,
    catPt * layoutFor(320).unit,
    (categoryAxis?.titleSize ?? AXIS_TITLE_PT) * layoutFor(320).unit,
  )
  const viewBoxHeight = bottom.viewBoxHeight
  const { unit, extra } = layoutFor(viewBoxHeight)
  const catFont = catPt * unit
  const textStyle = chartTextStyle(unit, categoryAxis, valueAxis)
  return (
    <svg
      className="chart-svg"
      viewBox={`${-extra} 0 ${600 + extra} ${viewBoxHeight}`}
      role="img"
      style={textStyle}
    >
      {plotAreaFill && <rect x="58" y="40" width="522" height="240" fill={plotAreaFill} />}
      <VerticalAxis
        minimum={bounds.min}
        maximum={bounds.max}
        ticks={isPercent ? undefined : bounds.ticks}
        numberFormat={isPercent ? '0%' : (valueAxis?.numFmt ?? primary.numberFormat)}
        showGridlines={gridlines !== false}
        hideLabels={valueAxis?.hidden === true}
        displayUnit={valueAxis?.displayUnit}
        onSelect={
          onElement
            ? (event) => {
                event.stopPropagation()
                onElement({ kind: 'value-axis' })
              }
            : undefined
        }
      />
      {seriesList.map((series, seriesIndex) => {
        const stroke = lineStroke(series, seriesIndex)
        if (stroke === null) return null
        const width = series.lineWidth ?? 3
        const values = displayValues(seriesIndex)
        const denominator = Math.max(1, values.length - 1)
        const pointAt = (index: number): string =>
          `${60 + (index / denominator) * 500},${
            280 - Math.max(0, Math.min(1, ((values[index] ?? 0) - bounds.min) / (span || 1))) * 240
          }`
        return seriesSegments(series).map((segment, segmentIndex) => (
          <polyline
            key={`${seriesIndex}-${segmentIndex}`}
            points={segment.map(pointAt).join(' ')}
            fill="none"
            stroke={stroke}
            strokeWidth={
              selectedEl?.kind === 'series' && selectedEl.seriesIndex === seriesIndex
                ? Math.max(5, width + 2)
                : width
            }
            onClick={
              onElement
                ? (event) => {
                    event.stopPropagation()
                    onElement({ kind: 'series', seriesIndex })
                  }
                : undefined
            }
          />
        ))
      })}
      {seriesList.map((series, seriesIndex) => {
        // Plot flag on or explicit symbol; a series-level "none" always wins.
        if (series.marker === 'none') return null
        if (lineMarkers !== true && series.marker === undefined) return null
        const symbol =
          series.marker !== undefined && series.marker !== 'auto'
            ? series.marker
            : (AUTO_MARKER_SYMBOLS[seriesIndex % AUTO_MARKER_SYMBOLS.length] ?? 'circle')
        const values = displayValues(seriesIndex)
        const pointCount = Math.max(1, values.length - 1)
        // Markers match the polyline stroke; an explicit noFill line keeps
        // its markers in the series color.
        const color = lineStroke(series, seriesIndex) ?? seriesColor(series, seriesIndex)
        return (
          <g key={`markers-${seriesIndex}`}>
            {values.map((value, index) =>
              isSkippedBlank(series, index) ? null : (
                <MarkerGlyph
                  key={index}
                  x={60 + (index / pointCount) * 500}
                  y={280 - Math.max(0, Math.min(1, (value - bounds.min) / (span || 1))) * 240}
                  symbol={symbol}
                  color={color}
                />
              ),
            )}
          </g>
        )
      })}
      {seriesList.map((series, seriesIndex) =>
        series.trendline === 'linear' ? (
          <TrendLine
            key={`trend-${seriesIndex}`}
            values={displayValues(seriesIndex)}
            minimum={bounds.min}
            maximum={bounds.max}
            color={seriesColor(series, seriesIndex)}
            xFor={(index) => 60 + (index / Math.max(1, series.values.length - 1)) * 500}
          />
        ) : null,
      )}
      {!categoryHidden &&
        primary.values.map((_, index) => (
          <CategoryTick
            key={index}
            x={60 + (index / count) * 500}
            label={categories[index] ?? String(index + 1)}
            slotWidth={500 / Math.max(1, count)}
            stride={categoryTickStride(
              categories,
              primary.values.length,
              500 / Math.max(1, count),
              CHAR_EM * catFont,
            )}
            index={index}
            fontUnits={catFont}
            onClick={
              onElement
                ? (event) => {
                    event.stopPropagation()
                    onElement({ kind: 'category-axis' })
                  }
                : undefined
            }
          />
        ))}
      {dataLabels === 'value' &&
        displayValues(0).map((displayed, index) =>
          isSkippedBlank(primary, index) ? null : (
            <text
              key={`label-${index}`}
              x={60 + (index / count) * 500}
              y={272 - Math.max(0, Math.min(1, (displayed - bounds.min) / span)) * 240}
              textAnchor="middle"
              className="data-label"
            >
              {formatAxisValue(primary.values[index] ?? 0, primary.numberFormat)}
            </text>
          ),
        )}
      <CategoryGroupBand
        spans={categoryGroups.map((group) => {
          const xAt = (index: number): number => 60 + (index / count) * 500
          // Group edges fall midway between the last point of one group
          // and the first point of the next (plot edges at the ends).
          return {
            label: group.label,
            xStart: group.start === 0 ? 60 : (xAt(group.start - 1) + xAt(group.start)) / 2,
            xEnd:
              group.end >= primary.values.length ? 560 : (xAt(group.end - 1) + xAt(group.end)) / 2,
          }
        })}
        charUnits={CHAR_EM * catFont}
        onClick={
          onElement
            ? (event) => {
                event.stopPropagation()
                onElement({ kind: 'category-axis' })
              }
            : undefined
        }
      />
      <AxisTitleTexts
        bottom={axisTitles?.category}
        left={axisTitles?.value}
        bottomY={bottom.bottomY}
        leftEdge={-extra}
        leftUnits={valueAxis?.displayUnitLabel}
        titleFont={(valueAxis?.titleSize ?? AXIS_TITLE_PT) * unit}
        labelFont={(valueAxis?.labelSize ?? AXIS_LABEL_PT) * unit}
      />
    </svg>
  )
}

function AreaChart({
  seriesList,
  axisTitles,
  dataLabels,
  grouping,
  gridlines,
  valueAxis,
  categoryFormat,
  categoryHidden = false,
  textBox,
  categoryAxis,
  plotAreaFill,
  onElement,
  selectedEl,
}: {
  readonly seriesList: readonly ChartSeries[]
  readonly axisTitles?: ChartAxisTitles
  readonly dataLabels?: ChartDataLabels
  readonly grouping?: ChartGrouping
  readonly gridlines?: boolean | undefined
  readonly valueAxis?: ChartValueAxis
  readonly categoryFormat?: string | undefined
  readonly categoryHidden?: boolean | undefined
  readonly textBox?: ChartTextBox | undefined
  readonly categoryAxis?: ChartAxisText | undefined
  readonly plotAreaFill?: string | undefined
} & ChartElementProps): React.JSX.Element {
  const primary = seriesList[0]
  if (!primary) return <></>
  const categories = primary.categories.map((value) => formatCategoryLabel(value, categoryFormat))
  const count = Math.max(1, primary.values.length - 1)
  const isStacked =
    (grouping === 'stacked' || grouping === 'percentStacked') && seriesList.length > 1
  const isPercent = grouping === 'percentStacked' && seriesList.length > 1
  // Cumulative per-category boundaries: series N fills between the running
  // total below it and the one including it.
  const stackBounds = isStacked
    ? seriesList.reduce<number[][]>((bounds, series) => {
        const previous = bounds[bounds.length - 1] ?? primary.values.map(() => 0)
        bounds.push(previous.map((base, index) => base + Math.max(0, series.values[index] ?? 0)))
        return bounds
      }, [])
    : []
  const percentTotals = isPercent
    ? primary.values.map(
        (_, index) =>
          seriesList.reduce((sum, series) => sum + Math.max(0, series.values[index] ?? 0), 0) || 1,
      )
    : []
  const bounds = isPercent
    ? { min: 0, max: 1, ticks: [0, 0.25, 0.5, 0.75, 1] }
    : axisBounds(
        isStacked
          ? Math.max(...(stackBounds[stackBounds.length - 1] ?? [0]), 0)
          : Math.max(...seriesList.flatMap((series) => [...series.values]), 0),
        isStacked ? undefined : valueAxis,
        // Stack bounds clamp negatives to 0; only the plain branch goes below.
        isStacked ? 0 : Math.min(...seriesList.flatMap((series) => [...series.values])),
      )
  const axisSpan = bounds.max - bounds.min
  const stackY = (raw: number, index: number): number =>
    280 - (isPercent ? raw / (percentTotals[index] ?? 1) : raw / (bounds.max || 1)) * 240
  const selectSeries = onElement
    ? (event: React.MouseEvent, seriesIndex: number): void => {
        event.stopPropagation()
        onElement({ kind: 'series', seriesIndex })
      }
    : undefined
  const areaNumberFormat = isPercent ? '0%' : (valueAxis?.numFmt ?? primary.numberFormat)
  const layoutFor = (viewBoxHeight: number): { unit: number; extra: number } =>
    valueAxisLayout(
      textBox,
      viewBoxHeight,
      valueLabelEm(
        bounds.ticks,
        valueAxis?.displayUnit,
        areaNumberFormat,
        valueAxis?.hidden === true,
      ),
      valueAxis?.labelSize ?? AXIS_LABEL_PT,
      axisSideReservePt(
        axisTitles?.value,
        valueAxis?.displayUnitLabel,
        valueAxis?.titleSize ?? AXIS_TITLE_PT,
        valueAxis?.labelSize ?? AXIS_LABEL_PT,
      ),
    )
  const catPt = categoryAxis?.labelSize ?? AXIS_LABEL_PT
  const bottom = bottomAxisLayout(
    categoryTicksWrap(
      categories,
      primary.values.length,
      500 / Math.max(1, count),
      CHAR_EM * catPt * layoutFor(320).unit,
    ),
    Boolean(axisTitles?.category),
    false,
    catPt * layoutFor(320).unit,
    (categoryAxis?.titleSize ?? AXIS_TITLE_PT) * layoutFor(320).unit,
  )
  const { unit, extra } = layoutFor(bottom.viewBoxHeight)
  const catFont = catPt * unit
  return (
    <svg
      className="chart-svg"
      viewBox={`${-extra} 0 ${600 + extra} ${bottom.viewBoxHeight}`}
      role="img"
      style={chartTextStyle(unit, categoryAxis, valueAxis)}
    >
      {plotAreaFill && <rect x="58" y="40" width="522" height="240" fill={plotAreaFill} />}
      <VerticalAxis
        minimum={bounds.min}
        maximum={bounds.max}
        ticks={isPercent ? undefined : bounds.ticks}
        numberFormat={isPercent ? '0%' : (valueAxis?.numFmt ?? primary.numberFormat)}
        showGridlines={gridlines !== false}
        hideLabels={valueAxis?.hidden === true}
        displayUnit={valueAxis?.displayUnit}
        onSelect={
          onElement
            ? (event) => {
                event.stopPropagation()
                onElement({ kind: 'value-axis' })
              }
            : undefined
        }
      />
      {seriesList.map((series, seriesIndex) => {
        const isSel = selectedEl?.kind === 'series' && selectedEl.seriesIndex === seriesIndex
        if (isStacked) {
          const upper = stackBounds[seriesIndex] ?? []
          const lower =
            seriesIndex === 0 ? upper.map(() => 0) : (stackBounds[seriesIndex - 1] ?? [])
          const xFor = (index: number): number => 60 + (index / count) * 500
          const upperPoints = upper.map((value, index) => `${xFor(index)},${stackY(value, index)}`)
          const lowerPoints = lower
            .map((value, index) => `${xFor(index)},${stackY(value, index)}`)
            .reverse()
          return (
            <g
              key={seriesIndex}
              onClick={selectSeries ? (event) => selectSeries(event, seriesIndex) : undefined}
            >
              <polygon
                points={[...upperPoints, ...lowerPoints].join(' ')}
                fill={seriesColor(series, seriesIndex)}
                opacity="0.75"
              />
              <polyline
                points={upperPoints.join(' ')}
                fill="none"
                stroke={isSel ? '#107C41' : seriesColor(series, seriesIndex)}
                strokeWidth={isSel ? '3.5' : '2'}
              />
            </g>
          )
        }
        const points = linePoints(series.values, bounds.max, bounds.min)
        const lastX = 60 + (Math.max(0, series.values.length - 1) / count) * 500
        const baseY = 280 - Math.max(0, Math.min(1, -bounds.min / (axisSpan || 1))) * 240
        return (
          <g
            key={seriesIndex}
            onClick={selectSeries ? (event) => selectSeries(event, seriesIndex) : undefined}
          >
            <polygon
              points={`60,${baseY} ${points} ${lastX},${baseY}`}
              fill={seriesColor(series, seriesIndex)}
              opacity="0.35"
            />
            <polyline
              points={points}
              fill="none"
              stroke={isSel ? '#107C41' : seriesColor(series, seriesIndex)}
              strokeWidth={isSel ? '4' : '2.5'}
            />
          </g>
        )
      })}
      {!categoryHidden &&
        primary.values.map((_, index) => (
          <CategoryTick
            key={index}
            x={60 + (index / count) * 500}
            label={categories[index] ?? String(index + 1)}
            slotWidth={500 / Math.max(1, count)}
            stride={categoryTickStride(
              categories,
              primary.values.length,
              500 / Math.max(1, count),
              CHAR_EM * catFont,
            )}
            index={index}
            fontUnits={catFont}
          />
        ))}
      {dataLabels === 'value' &&
        primary.values.map((value, index) => (
          <text
            key={`label-${index}`}
            x={60 + (index / count) * 500}
            y={272 - Math.max(0, Math.min(1, (value - bounds.min) / axisSpan)) * 240}
            textAnchor="middle"
            className="data-label"
          >
            {formatAxisValue(value, primary.numberFormat)}
          </text>
        ))}
      <AxisTitleTexts
        bottom={axisTitles?.category}
        left={axisTitles?.value}
        bottomY={bottom.bottomY}
        leftEdge={-extra}
        leftUnits={valueAxis?.displayUnitLabel}
        titleFont={(valueAxis?.titleSize ?? AXIS_TITLE_PT) * unit}
        labelFont={(valueAxis?.labelSize ?? AXIS_LABEL_PT) * unit}
      />
    </svg>
  )
}

/// Spider chart: one spoke per category, rings at quarter steps, one closed
/// polygon per series.
function RadarChart({
  seriesList,
  dataLabels,
  categoryFormat,
  textBox,
  categoryAxis,
  valueAxis,
  onElement,
  selectedEl,
}: {
  readonly seriesList: readonly ChartSeries[]
  readonly dataLabels?: ChartDataLabels
  readonly categoryFormat?: string | undefined
  readonly textBox?: ChartTextBox | undefined
  readonly categoryAxis?: ChartAxisText | undefined
  readonly valueAxis?: ChartValueAxis
} & ChartElementProps): React.JSX.Element {
  const primary = seriesList[0]
  if (!primary) return <></>
  const categories = primary.categories.map((value) => formatCategoryLabel(value, categoryFormat))
  const count = Math.min(Math.max(primary.values.length, 3), 12)
  const cx = 300
  const cy = 168
  const r = 118
  const maximum = niceAxisMaximum(
    Math.max(...seriesList.flatMap((series) => [...series.values]), 0),
  )
  const vertex = (index: number, radius: number): [number, number] => {
    const theta = (index / count) * 2 * Math.PI
    return [cx + Math.sin(theta) * radius, cy - Math.cos(theta) * radius]
  }
  const ringPoints = (fraction: number): string =>
    Array.from({ length: count }, (_, index) => vertex(index, r * fraction).join(',')).join(' ')
  const seriesPoints = (series: ChartSeries): string =>
    Array.from({ length: count }, (_, index) => {
      const value = Math.max(0, series.values[index] ?? 0)
      return vertex(index, Math.min(1, value / maximum) * r).join(',')
    }).join(' ')
  return (
    <svg
      className="chart-svg"
      viewBox="0 0 600 320"
      role="img"
      style={chartTextStyle(chartTextUnit(textBox, 600, 320), categoryAxis, valueAxis)}
    >
      {[0.25, 0.5, 0.75, 1].map((fraction) => (
        <polygon key={fraction} points={ringPoints(fraction)} fill="none" stroke="#e3e3e3" />
      ))}
      {Array.from({ length: count }, (_, index) => {
        const [x, y] = vertex(index, r)
        const [lx, ly] = vertex(index, r * 1.12)
        return (
          <g key={index}>
            <line x1={cx} y1={cy} x2={x} y2={y} stroke="#e3e3e3" />
            <text
              x={lx}
              y={ly + 3}
              textAnchor={Math.abs(lx - cx) < 8 ? 'middle' : lx > cx ? 'start' : 'end'}
              onClick={
                onElement
                  ? (event) => {
                      event.stopPropagation()
                      onElement({ kind: 'category-axis' })
                    }
                  : undefined
              }
            >
              {truncateLabel(categories[index] ?? String(index + 1), 8)}
            </text>
          </g>
        )
      })}
      {seriesList.map((series, seriesIndex) => {
        const isSel = selectedEl?.kind === 'series' && selectedEl.seriesIndex === seriesIndex
        return (
          <polygon
            key={seriesIndex}
            points={seriesPoints(series)}
            fill={seriesColor(series, seriesIndex)}
            fillOpacity="0.18"
            stroke={isSel ? '#107C41' : (lineStroke(series, seriesIndex) ?? 'none')}
            strokeWidth={
              isSel ? Math.max(4, (series.lineWidth ?? 2.5) + 1.5) : (series.lineWidth ?? 2.5)
            }
            onClick={
              onElement
                ? (event) => {
                    event.stopPropagation()
                    onElement({ kind: 'series', seriesIndex })
                  }
                : undefined
            }
          />
        )
      })}
      {dataLabels === 'value' &&
        Array.from({ length: count }, (_, index) => {
          const value = primary.values[index] ?? 0
          const [x, y] = vertex(index, Math.min(1, Math.max(0, value) / maximum) * r + 10)
          return (
            <text key={index} x={x} y={y} textAnchor="middle" className="data-label">
              {formatAxisValue(value, primary.numberFormat)}
            </text>
          )
        })}
    </svg>
  )
}

/// Scatter tick labels keep sub-integer precision that formatAxisValue's
/// one-decimal rounding would collapse (0.04 → "0.0").
function formatScatterTick(value: number, format: string | undefined): string {
  if (format !== undefined && format !== 'General') {
    return formatCategoryLabel(String(value), format)
  }
  if (Number.isInteger(value) || Math.abs(value) >= 1e4) return formatAxisValue(value, undefined)
  return String(Number(value.toPrecision(4)))
}

export function ScatterChart({
  seriesList,
  axisTitles,
  dataLabels,
  gridlines,
  valueAxis,
  xAxis,
  scatterStyle,
  categoryFormat,
  textBox,
  categoryAxis,
  plotAreaFill,
  onElement,
  selectedEl,
}: {
  readonly seriesList: readonly ChartSeries[]
  readonly axisTitles?: ChartAxisTitles
  readonly dataLabels?: ChartDataLabels
  readonly gridlines?: boolean | undefined
  readonly valueAxis?: ChartValueAxis
  readonly textBox?: ChartTextBox | undefined
  readonly categoryAxis?: ChartAxisText | undefined
  readonly plotAreaFill?: string | undefined
  readonly xAxis?:
    | {
        min?: number | undefined
        max?: number | undefined
        majorUnit?: number | undefined
        numFmt?: string | undefined
        majorGridlines: boolean
        hidden?: boolean | undefined
        displayUnit?: number | undefined
      }
    | undefined
  readonly scatterStyle?: string | undefined
  readonly categoryFormat?: string | undefined
} & ChartElementProps): React.JSX.Element {
  const points = seriesList.map((series) => {
    const xValues = series.categories.map((value, index) => {
      const parsed = Number.parseFloat(value)
      return Number.isFinite(parsed) ? parsed : index
    })
    // Blank cache slots are 0 fillers, not data — Excel plots no point.
    const blankSet = new Set(series.blanks ?? [])
    return { series, xValues, blankSet }
  })
  const allX = points.flatMap((entry) =>
    entry.xValues.filter((_, index) => !entry.blankSet.has(index)),
  )
  const allY = points.flatMap((entry) =>
    entry.series.values.filter((_, index) => !entry.blankSet.has(index)),
  )
  if (allY.length === 0) return <></>
  const boundsX = scatterAxisBounds(allX, xAxis)
  const boundsY = scatterAxisBounds(allY, valueAxis)
  // lineMarker / smoothMarker / line / smooth connect the points; a series
  // whose spPr line is an explicit noFill stays marker-only.
  const styleWantsLines = scatterStyle !== undefined && scatterStyle !== 'marker'
  const xFormat = xAxis?.numFmt ?? categoryFormat
  const plotX = (value: number): number =>
    60 + Math.max(0, Math.min(1, (value - boundsX.min) / (boundsX.max - boundsX.min))) * 520
  const plotY = (value: number): number =>
    280 - Math.max(0, Math.min(1, (value - boundsY.min) / (boundsY.max - boundsY.min))) * 240
  const yNumberFormat = valueAxis?.numFmt ?? seriesList[0]?.numberFormat
  const { unit, extra } = valueAxisLayout(
    textBox,
    320,
    valueLabelEm(boundsY.ticks, valueAxis?.displayUnit, yNumberFormat, valueAxis?.hidden === true),
    valueAxis?.labelSize ?? AXIS_LABEL_PT,
    axisSideReservePt(
      axisTitles?.value,
      valueAxis?.displayUnitLabel,
      valueAxis?.titleSize ?? AXIS_TITLE_PT,
      valueAxis?.labelSize ?? AXIS_LABEL_PT,
    ),
  )
  return (
    <svg
      className="chart-svg"
      viewBox={`${-extra} 0 ${600 + extra} 320`}
      role="img"
      style={chartTextStyle(unit, categoryAxis, valueAxis)}
    >
      {plotAreaFill && <rect x="58" y="40" width="522" height="240" fill={plotAreaFill} />}
      <VerticalAxis
        minimum={boundsY.min}
        maximum={boundsY.max}
        ticks={boundsY.ticks}
        numberFormat={yNumberFormat}
        showGridlines={gridlines !== false}
        hideLabels={valueAxis?.hidden === true}
        displayUnit={valueAxis?.displayUnit}
        onSelect={
          onElement
            ? (event) => {
                event.stopPropagation()
                onElement({ kind: 'value-axis' })
              }
            : undefined
        }
      />
      {boundsX.ticks.map((tick, index) => (
        <g key={index}>
          {xAxis?.majorGridlines && (
            <line
              x1={plotX(tick)}
              y1="40"
              x2={plotX(tick)}
              y2="280"
              stroke="#e3e3e3"
              strokeWidth="1"
            />
          )}
          {xAxis?.hidden !== true && (
            <text
              x={plotX(tick)}
              y="296"
              textAnchor="middle"
              className="axis-label"
              onClick={
                onElement
                  ? (event) => {
                      event.stopPropagation()
                      onElement({ kind: 'category-axis' })
                    }
                  : undefined
              }
            >
              {formatScatterTick(xAxis?.displayUnit ? tick / xAxis.displayUnit : tick, xFormat)}
            </text>
          )}
        </g>
      ))}
      {points.map(({ series, xValues, blankSet }, seriesIndex) => (
        <g key={seriesIndex}>
          {styleWantsLines && series.lineColor !== 'none' && series.values.length > 1 && (
            <polyline
              points={series.values
                .map((value, index) =>
                  blankSet.has(index) ? null : `${plotX(xValues[index] ?? 0)},${plotY(value)}`,
                )
                .filter((point) => point !== null)
                .join(' ')}
              fill="none"
              stroke={
                series.lineColor && series.lineColor !== 'none'
                  ? series.lineColor
                  : seriesColor(series, seriesIndex)
              }
              strokeWidth={series.lineWidth ?? 2}
            />
          )}
          {series.marker !== 'none' &&
            series.values.map((value, index) =>
              blankSet.has(index) ? null : (
                <circle
                  key={index}
                  cx={plotX(xValues[index] ?? 0)}
                  cy={plotY(value)}
                  r={isSelectedPoint(selectedEl, seriesIndex, index) ? 6 : 4}
                  fill={seriesColor(series, seriesIndex)}
                  opacity="0.85"
                  {...(isSelectedPoint(selectedEl, seriesIndex, index)
                    ? { stroke: '#107C41', strokeWidth: 2 }
                    : {})}
                  onClick={
                    onElement
                      ? (event) => {
                          event.stopPropagation()
                          onElement(narrowSelection(selectedEl ?? null, seriesIndex, index))
                        }
                      : undefined
                  }
                />
              ),
            )}
          {dataLabels === 'value' &&
            series.values.map((value, index) =>
              blankSet.has(index) ? null : (
                <text
                  key={`label-${index}`}
                  x={plotX(xValues[index] ?? 0)}
                  y={plotY(value) - 8}
                  textAnchor="middle"
                  className="data-label"
                >
                  {formatAxisValue(value, series.numberFormat)}
                </text>
              ),
            )}
        </g>
      ))}
      <AxisTitleTexts
        bottom={axisTitles?.category}
        left={axisTitles?.value}
        leftEdge={-extra}
        leftUnits={valueAxis?.displayUnitLabel}
        titleFont={(valueAxis?.titleSize ?? AXIS_TITLE_PT) * unit}
        labelFont={(valueAxis?.labelSize ?? AXIS_LABEL_PT) * unit}
      />
    </svg>
  )
}

interface PieSliceLabel {
  key: number
  inside: boolean
  x: number
  y: number
  anchor: 'start' | 'middle' | 'end'
  lines: string[]
  leader?: string
}

/// Label placement: slices big enough hold their label at
/// two-thirds radius; small slices move it outside with a leader line, and
/// outside labels on each side push apart until they no longer overlap.
function pieSliceLabels(
  values: readonly number[],
  categories: readonly string[],
  total: number,
  mode: NonNullable<ChartDataLabels>,
  geometry: { cx: number; cy: number; r: number; inner: number; offsets: readonly number[] },
  position?: ChartLabelPosition,
  formatCode?: string | undefined,
  /// dLbls' own numFmt only — the source cells' format is not a percent.
  percentFormat?: string | undefined,
): PieSliceLabel[] {
  const { cx, cy, r, inner, offsets } = geometry
  const labels: PieSliceLabel[] = []
  let cursor = 0
  for (const [index, value] of values.entries()) {
    const share = Math.max(0, value) / total
    const mid = (cursor + share / 2) * 2 * Math.PI
    cursor += share
    // Excel labels every visible slice; only slivers too thin to aim a
    // leader line at go unlabeled.
    if (share < 0.005) continue
    const percentText = formatPiePercent(share, percentFormat)
    const valueText = formatPieValue(value, formatCode)
    const lines =
      mode === 'category-value-percent'
        ? [truncateLabel(categories[index] ?? '', 12), valueText, percentText]
        : mode === 'category-percent'
          ? [truncateLabel(categories[index] ?? '', 12), percentText]
          : [mode === 'value' ? valueText : percentText]
    const sin = Math.sin(mid)
    const cos = Math.cos(mid)
    // An exploded slice carries its label out with it.
    const offset = offsets[index] ?? 0
    // An explicit position overrides the best-fit heuristic.
    const inside =
      position === 'outside-end' ? false : position !== undefined ? true : share >= 0.09
    if (inside) {
      const baseRadius =
        inner > 0
          ? (r + inner) / 2
          : position === 'inside-end'
            ? r * 0.8
            : position === 'center'
              ? r * 0.5
              : r * 0.62
      const radius = baseRadius + offset
      labels.push({
        key: index,
        inside: true,
        x: cx + sin * radius,
        y: cy - cos * radius,
        anchor: 'middle',
        lines,
      })
      continue
    }
    const elbowR = r * 1.1 + offset
    const tick = 10
    const rightSide = sin >= 0
    const elbowX = cx + sin * elbowR
    const elbowY = cy - cos * elbowR
    const textX = elbowX + (rightSide ? tick : -tick)
    labels.push({
      key: index,
      inside: false,
      x: textX + (rightSide ? 3 : -3),
      y: elbowY,
      anchor: rightSide ? 'start' : 'end',
      lines,
      leader: `${cx + sin * (r + offset)},${cy - cos * (r + offset)} ${elbowX},${elbowY} ${textX},${elbowY}`,
    })
  }
  // Outside labels: per side, top-down, keep a minimum vertical gap.
  for (const side of ['start', 'end'] as const) {
    const outside = labels
      .filter((label) => !label.inside && label.anchor === side)
      .sort((a, b) => a.y - b.y)
    const gap = 15
    for (let i = 1; i < outside.length; i += 1) {
      const previous = outside[i - 1]
      const current = outside[i]
      if (previous && current && current.y < previous.y + gap) current.y = previous.y + gap
    }
  }
  return labels
}

function PieChart({
  series,
  isDoughnut,
  legend,
  dataLabels,
  dataLabelPosition,
  dataLabelFormat,
  dataLabelStyle,
  holeSizePct,
  categoryFormat,
  onElement,
  selectedEl,
}: {
  readonly series: ChartSeries
  readonly isDoughnut: boolean
  readonly legend?: ChartMetadata['legend']
  readonly dataLabels?: ChartDataLabels
  readonly dataLabelPosition?: ChartLabelPosition
  readonly dataLabelFormat?: string | undefined
  readonly dataLabelStyle?: ChartMetadata['dataLabelStyle']
  readonly holeSizePct?: number | undefined
  readonly categoryFormat?: string | undefined
} & ChartElementProps): React.JSX.Element {
  const { values } = series
  const categories = series.categories.map((value) => formatCategoryLabel(value, categoryFormat))
  const total = values.reduce((sum, value) => sum + Math.max(0, value), 0) || 1
  const cx = 210
  const cy = 160
  const r = 100
  // 50 mirrors what buildChartXml writes for new doughnuts, so the render
  // does not jump after the first save.
  const inner = isDoughnut ? (r * Math.min(90, Math.max(10, holeSizePct ?? 50))) / 100 : 0
  const point = (radius: number, theta: number): [number, number] => [
    cx + Math.sin(theta) * radius,
    cy - Math.cos(theta) * radius,
  ]
  // Explosion: per-slice dPt override, else the series-level value; render
  // caps the shift at 40% of the radius so labels stay inside the viewbox.
  const explosionOffset = (index: number): number => {
    const pct =
      series.pointExplosions?.find((entry) => entry.index === index)?.pct ??
      series.explosionPct ??
      0
    return (Math.min(100, Math.max(0, pct)) / 100) * r * 0.4
  }
  const offsets = values.map((_, index) => explosionOffset(index))
  let angle = 0
  const slices = values
    .map((value, index) => {
      const share = Math.max(0, value) / total
      const start = angle
      angle += share * 2 * Math.PI
      return { index, share, start, end: angle, color: pieSliceColor(series, index) }
    })
    .filter((slice) => slice.share > 0.0005)
  const slicePath = (slice: (typeof slices)[number]): string => {
    // A lone full-circle slice has coincident endpoints; nudge the end angle
    // so the arc still spans the whole ring.
    const end = slice.share > 0.9995 ? slice.start + 2 * Math.PI - 0.001 : slice.end
    const large = end - slice.start > Math.PI ? 1 : 0
    const [x1, y1] = point(r, slice.start)
    const [x2, y2] = point(r, end)
    if (inner === 0) {
      return `M ${cx} ${cy} L ${x1} ${y1} A ${r} ${r} 0 ${large} 1 ${x2} ${y2} Z`
    }
    const [ix1, iy1] = point(inner, slice.start)
    const [ix2, iy2] = point(inner, end)
    return (
      `M ${x1} ${y1} A ${r} ${r} 0 ${large} 1 ${x2} ${y2}` +
      ` L ${ix2} ${iy2} A ${inner} ${inner} 0 ${large} 0 ${ix1} ${iy1} Z`
    )
  }
  const labels =
    dataLabels !== undefined && dataLabels !== 'none'
      ? pieSliceLabels(
          values,
          categories,
          total,
          dataLabels,
          { cx, cy, r, inner, offsets },
          dataLabelPosition,
          // Labels without their own numFmt inherit the source cells' format.
          dataLabelFormat ?? series.numberFormat,
          dataLabelFormat,
        )
      : []
  return (
    <div className={`pie-layout pie-legend-${legend ?? 'right'}`}>
      <svg className="chart-svg pie-svg" viewBox="0 0 420 320" role="img">
        {slices.map((slice) => {
          const mid = (slice.start + slice.end) / 2
          const offset = offsets[slice.index] ?? 0
          const selected = isSelectedPoint(selectedEl, 0, slice.index)
          return (
            <path
              key={slice.index}
              d={slicePath(slice)}
              transform={
                offset > 0
                  ? `translate(${Math.sin(mid) * offset} ${-Math.cos(mid) * offset})`
                  : undefined
              }
              fill={slice.color}
              stroke={selected ? '#107C41' : '#fff'}
              strokeWidth={selected ? 2.5 : 1.5}
              onClick={
                onElement
                  ? (event) => {
                      event.stopPropagation()
                      onElement({ kind: 'point', seriesIndex: 0, pointIndex: slice.index })
                    }
                  : undefined
              }
            />
          )
        })}
        {labels.map((label) => (
          <g key={label.key}>
            {label.leader && (
              <polyline points={label.leader} fill="none" stroke="#9a9a9a" strokeWidth="1" />
            )}
            <text
              x={label.x}
              y={label.y + (label.lines.length > 1 ? -2 : 3)}
              textAnchor={label.anchor}
              className={label.inside ? 'pie-label-inside' : 'pie-label-outside'}
              style={labelTextStyle(dataLabelStyle)}
            >
              {label.lines.map((line, lineIndex) => (
                <tspan key={lineIndex} x={label.x} dy={lineIndex === 0 ? 0 : 12}>
                  {line}
                </tspan>
              ))}
            </text>
          </g>
        ))}
      </svg>
      {legend !== 'none' && (
        <div
          className={`chart-legend${selectedEl?.kind === 'legend' ? ' chart-el-selected' : ''}`}
          onClick={
            onElement
              ? (event) => {
                  event.stopPropagation()
                  onElement({ kind: 'legend' })
                }
              : undefined
          }
        >
          {categories.slice(0, 12).map((category, index) => (
            <span key={`${category}-${index}`}>
              <i style={{ background: pieSliceColor(series, index) }} />
              {truncateLabel(category, 12)}
            </span>
          ))}
          {categories.length > 12 && (
            <span>{t('appMoreItems', { count: categories.length - 12 })}</span>
          )}
        </div>
      )}
    </div>
  )
}

export function truncateLabel(value: string, maximumLength = 14): string {
  return value.length > maximumLength ? `${value.slice(0, maximumLength)}…` : value
}
