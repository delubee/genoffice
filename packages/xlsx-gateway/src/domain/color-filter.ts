import { z } from 'zod'
import { resolveStyleColor, type StyleColor } from './style-color'

const color = z.string().regex(/^#[0-9A-Fa-f]{6}$/)
export const colorFilterCriteriaSchema = z.union([
  z.object({ fill: z.array(color.nullable()).min(1).max(2000) }).strict(),
  z.object({ font: z.array(color).min(1).max(2000) }).strict(),
])
export type ColorFilterCriteria = z.infer<typeof colorFilterCriteriaSchema>

export function normalizedColorFilters(criteria: ColorFilterCriteria): ColorFilterCriteria {
  const colors = colorFilterCriteriaSchema.parse(criteria)
  return 'fill' in colors
    ? { fill: colors.fill.map((value) => (value === null ? null : value.toUpperCase())) }
    : { font: colors.font.map((value) => value.toUpperCase()) }
}

export function colorFilterMatches(
  criteria: ColorFilterCriteria,
  style:
    | { fillColor?: StyleColor | null | undefined; fontColor?: StyleColor | null | undefined }
    | undefined,
  palette?: readonly string[],
) {
  if ('fill' in criteria) {
    const value = style?.fillColor
      ? resolveStyleColor(style.fillColor, palette).toUpperCase()
      : null
    return criteria.fill.some((color) =>
      color === null ? value === null : color.toUpperCase() === value,
    )
  }
  const value = style?.fontColor
    ? resolveStyleColor(style.fontColor, palette).toUpperCase()
    : '#000000'
  return criteria.font.some((color) => color.toUpperCase() === value)
}
