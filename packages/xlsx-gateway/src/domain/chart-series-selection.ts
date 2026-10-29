import { z } from 'zod'
import { chartNumericValue } from './chart-series-range'
import {
  chartSupportsSeriesReplace,
  transposeChartSeries,
  type ChartStateEdit,
  type ChartVisualState,
} from './chart-visual'

export const chartSeriesSelectionSchema = z.discriminatedUnion('mode', [
  z
    .object({
      mode: z.literal('rows'),
      rows: z
        .array(
          z
            .object({
              sourceIndex: z.number().int().min(0).max(255).optional(),
              name: z.string().max(255),
              valuesRef: z.string().min(1).max(512).optional(),
              categoriesRef: z.string().min(1).max(512).optional(),
            })
            .strict(),
        )
        .min(1)
        .max(24),
    })
    .strict(),
  z
    .object({
      mode: z.literal('transpose'),
      defaultNames: z.array(z.string().max(255)).max(24).optional(),
    })
    .strict(),
])
export type ChartSeriesSelection = z.infer<typeof chartSeriesSelectionSchema>
export interface SelectionVector {
  vector: readonly (string | number | boolean | null | undefined)[]
  ref: string
}

export function chartSelectionReferences(
  chart: ChartVisualState,
  intent: ChartSeriesSelection,
): string[] {
  return [
    ...new Set(
      [
        ...chart.series.flatMap((series) => [
          series.valuesRef,
          series.categoriesRef,
          series.nameRef,
        ]),
        ...(intent.mode === 'rows'
          ? intent.rows.flatMap((row) => [row.valuesRef, row.categoriesRef])
          : []),
      ].filter((value): value is string => !!value),
    ),
  ]
}

/** No caller-provided caches: values come from the current chart or verified reads. */
export function buildChartSeriesSelection(
  chart: ChartVisualState,
  intent: ChartSeriesSelection,
  read: (reference: string) => SelectionVector,
): NonNullable<ChartStateEdit['seriesSet']> {
  if (!chartSupportsSeriesReplace(chart.chartTypes))
    throw new Error('unsupported_chart_series_replacement')
  const base = chart.series.map((series) => ({
    ...series,
    ...(series.valuesRef
      ? {
          values: read(series.valuesRef).vector.map((value) => chartNumericValue(value) ?? 0),
          valuesRef: read(series.valuesRef).ref,
        }
      : {}),
    ...(series.categoriesRef
      ? {
          categories: read(series.categoriesRef).vector.map((value) =>
            String(value ?? '').slice(0, 255),
          ),
          categoriesRef: read(series.categoriesRef).ref,
        }
      : {}),
    ...(series.nameRef ? { name: String(read(series.nameRef).vector[0] ?? '').slice(0, 255) } : {}),
  }))
  if (intent.mode === 'transpose') {
    const result = transposeChartSeries(base, (n) => intent.defaultNames?.[n - 1] ?? `Series ${n}`)
    if (!result) throw new Error('chart_categories_required')
    return result
  }
  return intent.rows.map((row) => {
    const original = row.sourceIndex === undefined ? undefined : base[row.sourceIndex]
    if (row.sourceIndex !== undefined && !original) throw new Error('unknown_chart_series')
    if (!original && !row.valuesRef) throw new Error('chart_series_values_required')
    const values = row.valuesRef ? read(row.valuesRef) : undefined
    const categories = row.categoriesRef ? read(row.categoriesRef) : undefined
    const numbers = values
      ? values.vector.map((value) => chartNumericValue(value) ?? 0)
      : original!.values
    return {
      name: row.name,
      values: [...numbers],
      categories: categories
        ? categories.vector.map((value) => String(value ?? '').slice(0, 255))
        : [...(original?.categories ?? base[0]?.categories.slice(0, numbers.length) ?? [])],
      ...(values
        ? { valuesRef: values.ref }
        : original?.valuesRef
          ? { valuesRef: original.valuesRef }
          : {}),
      ...(categories
        ? { categoriesRef: categories.ref }
        : original?.categoriesRef
          ? { categoriesRef: original.categoriesRef }
          : {}),
      ...(original?.color ? { color: original.color } : {}),
    }
  })
}
