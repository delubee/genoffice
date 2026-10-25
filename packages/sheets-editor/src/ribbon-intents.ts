import {
  workbookOperationSchema,
  type WorkbookOperation,
  type CellFormatPatch,
} from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { parseRange, formatAddress, columnLabel } from '@genoffice/xlsx-gateway/domain/cell-address'
import type { SelectionFormat } from '../../../apps/sheets/src/renderer/selection-format'
import { stepFontSize } from '../../../apps/sheets/src/renderer/font-size-ladder'
import { CELL_STYLE_PRESETS } from '../../../apps/sheets/src/renderer/cell-style-presets'

export interface RibbonTarget {
  sheetId: string
  address: string
  format: SelectionFormat
  doubleUnderline: boolean
}

/** Decode desktop command strings to validated host operations, without touching Univer. */
export function ribbonIntents(command: string, target: RibbonTarget): WorkbookOperation[] {
  const colon = command.indexOf(':')
  const name = colon < 0 ? command : command.slice(0, colon)
  const argument = colon < 0 ? '' : command.slice(colon + 1)
  const { sheetId, address, format: echo } = target
  const bounds = parseRange(address)
  const range = (top: number, left: number, bottom: number, right: number) =>
    `${formatAddress(top, left)}:${formatAddress(bottom, right)}`
  const fmt = (format: CellFormatPatch, selection = address): WorkbookOperation => ({
    op: 'format_range',
    sheetId,
    range: selection,
    format,
  })
  const bool = (current: boolean) =>
    argument === ''
      ? !current
      : argument === 'on'
        ? true
        : argument === 'off'
          ? false
          : (() => {
              throw new Error('Invalid toggle value')
            })()
  let result: unknown[]
  switch (name) {
    case 'bold':
      result = [fmt({ bold: bool(echo.bold) })]
      break
    case 'italic':
      result = [fmt({ italic: bool(echo.italic) })]
      break
    case 'strike':
      result = [fmt({ strikethrough: bool(echo.strike) })]
      break
    case 'underline':
      result = [
        fmt(
          argument === 'double'
            ? { underline: !target.doubleUnderline, underlineStyle: 'double' }
            : { underline: bool(echo.underline), underlineStyle: 'single' },
        ),
      ]
      break
    case 'superscript':
    case 'subscript':
      result = [fmt({ vertAlign: echo.vertAlign === name ? null : name })]
      break
    case 'font-family':
      result = [fmt({ fontFamily: argument })]
      break
    case 'font-size':
      result = [fmt({ fontSize: Number(argument) })]
      break
    case 'font-size-step': {
      if (argument !== '1' && argument !== '-1') throw new Error('Invalid font-size step')
      result = [fmt({ fontSize: stepFontSize(echo.fontSize ?? 11, Number(argument) as 1 | -1) })]
      break
    }
    case 'font-color':
      result = [fmt({ fontColor: argument === 'auto' ? null : argument })]
      break
    case 'fill':
      result = [fmt({ fillColor: argument === 'none' ? null : argument })]
      break
    // Keep the complete remainder: number formats and font names can contain colons.
    case 'format':
      result = [fmt({ numberFormat: argument })]
      break
    case 'align':
      result = [fmt({ horizontalAlign: argument as CellFormatPatch['horizontalAlign'] })]
      break
    case 'valign':
      result = [
        fmt({
          verticalAlign: (argument === 'middle'
            ? 'center'
            : argument) as CellFormatPatch['verticalAlign'],
        }),
      ]
      break
    case 'wrap':
      result = [fmt({ wrapText: bool(echo.wrap) })]
      break
    case 'rotate':
      result = [fmt({ textRotation: argument === 'vertical' ? 'vertical' : Number(argument) })]
      break
    case 'indent':
      result = [fmt({ indent: Number(argument) })]
      break
    case 'cellprot': {
      const patch: CellFormatPatch = {}
      for (const value of argument.split(':')) {
        if (!value) continue
        const match = /^(locked|hidden)-(on|off)$/.exec(value)
        if (!match) throw new Error('Invalid cell protection flag')
        patch[match[1] === 'locked' ? 'protectionLocked' : 'protectionHidden'] = match[2] === 'on'
      }
      result = [fmt(patch)]
      break
    }
    case 'border': {
      const [kind, color, style = 'thin'] = argument.split(':')
      const edge = (type: string, selection = address) =>
        fmt(
          {
            border: {
              type: type as 'all',
              ...(color ? { color } : {}),
              style: (kind === 'thick-outer' ? 'thick' : style) as 'thin',
            },
          },
          selection,
        )
      result =
        kind === 'outer' || kind === 'thick-outer'
          ? [
              edge(
                'top',
                range(bounds.startRow, bounds.startColumn, bounds.startRow, bounds.endColumn),
              ),
              edge(
                'bottom',
                range(bounds.endRow, bounds.startColumn, bounds.endRow, bounds.endColumn),
              ),
              edge(
                'left',
                range(bounds.startRow, bounds.startColumn, bounds.endRow, bounds.startColumn),
              ),
              edge(
                'right',
                range(bounds.startRow, bounds.endColumn, bounds.endRow, bounds.endColumn),
              ),
            ]
          : [edge(kind ?? '')]
      break
    }
    case 'cell-style': {
      const preset = CELL_STYLE_PRESETS[argument]
      if (!preset) throw new Error('Unknown cell style')
      result = preset.map((patch) => fmt(patch))
      break
    }
    case 'merge':
      if (argument === 'unmerge') result = [{ op: 'unmerge_cells', sheetId, range: address }]
      else if (argument === 'across') {
        if (bounds.endRow - bounds.startRow >= 500) throw new Error('range_limit_exceeded')
        result = Array.from({ length: bounds.endRow - bounds.startRow + 1 }, (_, i) => ({
          op: 'merge_cells',
          sheetId,
          range: range(
            bounds.startRow + i,
            bounds.startColumn,
            bounds.startRow + i,
            bounds.endColumn,
          ),
        }))
      } else if (argument === 'cells' || argument === 'center') {
        result = [{ op: 'merge_cells', sheetId, range: address }]
        if (argument === 'center')
          result.push(
            fmt({ horizontalAlign: 'center' }, formatAddress(bounds.startRow, bounds.startColumn)),
          )
      } else throw new Error('Invalid merge option')
      break
    case 'insert-row-here':
    case 'delete-row-here':
      result = [
        {
          op: name.startsWith('insert') ? 'insert_rows' : 'delete_rows',
          sheetId,
          row: bounds.startRow + 1,
          count: 1,
        },
      ]
      break
    case 'insert-col-here':
    case 'delete-col-here':
      result = [
        {
          op: name.startsWith('insert') ? 'insert_cols' : 'delete_cols',
          sheetId,
          column: columnLabel(bounds.startColumn),
          count: 1,
        },
      ]
      break
    case 'sort':
      result = [
        {
          op: 'sort_range',
          sheetId,
          range: address,
          byColumn: columnLabel(bounds.startColumn),
          order: argument,
          hasHeader: false,
        },
      ]
      break
    case 'fill-down':
    case 'fill-up':
    case 'fill-left':
    case 'fill-right': {
      const vertical = name === 'fill-down' || name === 'fill-up'
      const row = name === 'fill-up' ? bounds.endRow : bounds.startRow
      const column = name === 'fill-left' ? bounds.endColumn : bounds.startColumn
      result = [
        {
          op: 'fill_range',
          sheetId,
          source: vertical
            ? range(row, bounds.startColumn, row, bounds.endColumn)
            : range(bounds.startRow, column, bounds.endRow, column),
          target: address,
        },
      ]
      break
    }
    case 'clear-contents':
      result = [{ op: 'clear_range', sheetId, range: address }]
      break
    default:
      throw new Error(`Desktop command is not connected yet: ${command}`)
  }
  // Validate the whole batch before any command is sent, including dialog batches.
  return result.map((operation) => workbookOperationSchema.parse(operation))
}
