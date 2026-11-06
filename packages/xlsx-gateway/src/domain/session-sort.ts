import { workbookFilterStateSchema } from '../shared/workbook-protocol'
import { visibilityEditsSchema, axisHiddenOverride } from './session-visibility'

/** Desktop sort keeps filtered and manually hidden rows stationary. Host rows
 * include imported visibility; session journals may have changed earlier in the batch. */
export function sortRowVisibility(
  features: Record<string, unknown>,
  sourceHiddenRows: readonly number[] = [],
) {
  const hidden = new Set(sourceHiddenRows)
  const filter = features.filterEdits
    ? workbookFilterStateSchema.parse(features.filterEdits)
    : undefined
  const filtered = new Set(filter?.hiddenRows)
  const visibility = visibilityEditsSchema.parse(features.visibilityEdits ?? [])
  return (row: number) => {
    const base =
      filter && row > filter.visibilityRange.startRow && row <= filter.visibilityRange.endRow
        ? filtered.has(row)
        : hidden.has(row)
    return !(axisHiddenOverride(visibility, 'rows', row) ?? base)
  }
}
