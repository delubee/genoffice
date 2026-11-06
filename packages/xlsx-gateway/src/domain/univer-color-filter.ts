import { ColorKit } from '@univerjs/core'
import { colorFilterCriteriaSchema, type ColorFilterCriteria } from './color-filter'

export function fromUniverColorFilter(criteria: {
  cellFillColors?: (string | null)[]
  cellTextColors?: string[]
}) {
  const hex = (value: string) => new ColorKit(value).toHexString().toUpperCase()
  return colorFilterCriteriaSchema.parse(
    criteria.cellFillColors
      ? { fill: criteria.cellFillColors.map((value) => (value === null ? null : hex(value))) }
      : { font: criteria.cellTextColors?.map(hex) },
  )
}
export function toUniverColorFilter(criteria: ColorFilterCriteria) {
  const rgb = (value: string) => new ColorKit(value).toRgbString()
  return 'fill' in criteria
    ? { cellFillColors: criteria.fill.map((value) => (value === null ? null : rgb(value))) }
    : { cellTextColors: criteria.font.map(rgb) }
}
