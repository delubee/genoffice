import { formatAddress, parseRange } from '@genoffice/xlsx-gateway/domain/cell-address'
import {
  workbookOperationSchema,
  type SetPageSetupOperation,
} from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { clampTitleRows } from './print-settings'
type PageFields = Omit<SetPageSetupOperation, 'op' | 'sheetId' | 'header' | 'footer'>
export type PageSettingsPatch = { [K in keyof PageFields]?: Exclude<PageFields[K], undefined> }
export interface PageSettingsContext {
  sheetId: string
  address: string
  dimensions: { rows: number; columns: number }
  prior: PageSettingsPatch
  breaks: { rowBreaks: number[]; colBreaks: number[] }
}
/** Shared desktop command geometry and field semantics; callers own persistence. */
export function pageSettingsCommand(
  rest: string,
  context: PageSettingsContext,
): PageSettingsPatch | null {
  const at = rest.indexOf(':'),
    key = at < 0 ? rest : rest.slice(0, at),
    value = at < 0 ? '' : rest.slice(at + 1)
  const range = parseRange(context.address)
  let patch: PageSettingsPatch
  switch (key) {
    case 'orientation':
      if (value !== 'portrait' && value !== 'landscape') return null
      patch = { orientation: value }
      break
    case 'margins':
      if (!['normal', 'wide', 'narrow'].includes(value)) return null
      patch = { margins: value as 'normal' }
      break
    case 'paper':
      patch = { paperSize: Number(value) }
      break
    case 'scale':
      patch = { scale: Number(value) }
      break
    case 'fit-width':
    case 'fit-height':
      patch = {
        fitToWidth: key === 'fit-width' ? Number(value) : (context.prior.fitToWidth ?? 0),
        fitToHeight: key === 'fit-height' ? Number(value) : (context.prior.fitToHeight ?? 0),
      }
      break
    case 'fit': {
      const [width, height] = value.split(',').map(Number)
      if (width === undefined || height === undefined) return null
      patch = { fitToWidth: width!, fitToHeight: height! }
      break
    }
    case 'print-gridlines':
      patch = { printGridlines: value === '1' }
      break
    case 'print-headings':
      patch = { printHeadings: value === '1' }
      break
    case 'print-area':
      patch = {
        printArea:
          value === 'clear'
            ? null
            : `${formatAddress(range.startRow, range.startColumn)}:${formatAddress(range.endRow, range.endColumn)}`,
      }
      break
    case 'print-titles':
      patch = {
        printTitles:
          value === 'clear'
            ? null
            : value === 'first-row'
              ? '1:1'
              : clampTitleRows(range.startRow + 1, range.endRow + 1),
      }
      break
    case 'breaks': {
      if (value === 'reset') {
        patch = { rowBreaks: [], colBreaks: [] }
        break
      }
      if (!['insert', 'remove'].includes(value)) return null
      const row = range.startRow,
        column = range.startColumn
      const wantRow = range.endRow - range.startRow + 1 < context.dimensions.rows && row > 0
      const wantColumn =
        range.endColumn - range.startColumn + 1 < context.dimensions.columns && column > 0
      // Full-row selection creates only the horizontal break, and vice versa.
      const rows = context.breaks.rowBreaks.filter((id) => id > 0),
        cols = context.breaks.colBreaks.filter((id) => id > 0)
      if (value === 'insert' && !wantRow && !wantColumn) return null
      if (
        value === 'remove' &&
        !rows.some((id) => wantRow && id === row) &&
        !cols.some((id) => wantColumn && id === column)
      )
        return null
      if (value === 'insert')
        patch = {
          rowBreaks: wantRow ? [...new Set([...rows, row])].sort((a, b) => a - b) : rows,
          colBreaks: wantColumn ? [...new Set([...cols, column])].sort((a, b) => a - b) : cols,
        }
      else
        patch = {
          rowBreaks: rows.filter((id) => !wantRow || id !== row),
          colBreaks: cols.filter((id) => !wantColumn || id !== column),
        }
      break
    }
    default:
      return null
  }
  const parsed = workbookOperationSchema.safeParse({
    op: 'set_page_setup',
    sheetId: context.sheetId,
    ...patch,
  })
  if (!parsed.success || parsed.data.op !== 'set_page_setup') return null
  const { op: _op, sheetId: _id, ...result } = parsed.data
  return Object.fromEntries(
    Object.entries(result).filter(([, value]) => value !== undefined),
  ) as PageSettingsPatch
}
