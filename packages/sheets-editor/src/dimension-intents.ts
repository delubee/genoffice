import { columnLabel } from '@genoffice/xlsx-gateway/domain/cell-address'
import {
  workbookOperationSchema,
  type WorkbookOperation,
} from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import type { GridRange } from './ports'

/** Translate pixels from the grid to the same validated operations used by AI. */
export function dimensionIntents(
  sheetId: string,
  axis: 'row' | 'column',
  ranges: readonly GridRange[],
  pixels: number,
): WorkbookOperation[] {
  if (!Number.isFinite(pixels)) throw new Error('invalid_dimension')
  const result: WorkbookOperation[] = []
  const spans = new Set<string>()
  for (const range of ranges) {
    const first = axis === 'row' ? range.startRow : range.startColumn
    const last = axis === 'row' ? range.endRow : range.endColumn
    const chunk = axis === 'row' ? 500 : 100
    for (let start = first; start <= last; start += chunk) {
      const count = Math.min(chunk, last - start + 1)
      const key = `${start}:${count}`
      if (spans.has(key)) continue
      if (result.length >= 1000) throw new Error('dimension_limit: select fewer rows or columns')
      spans.add(key)
      result.push(
        workbookOperationSchema.parse(
          axis === 'row'
            ? {
                op: 'set_row_height',
                sheetId,
                row: start + 1,
                count,
                heightPoints: Math.max(2, Math.min(409, Math.round(pixels * 0.75 * 100) / 100)),
              }
            : {
                op: 'set_col_width',
                sheetId,
                column: columnLabel(start),
                count,
                widthPx: Math.max(10, Math.min(2000, Math.round(pixels))),
              },
        ),
      )
    }
  }
  if (!result.length) throw new Error('missing_dimension_range')
  return result
}
