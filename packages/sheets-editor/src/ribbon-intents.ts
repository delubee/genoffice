import { normalizeLinkTarget } from '../../../apps/sheets/src/renderer/link-target'
import {
  workbookOperationSchema,
  type WorkbookOperation,
  type CellFormatPatch,
} from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { parseRange, formatAddress, columnLabel } from '@genoffice/xlsx-gateway/domain/cell-address'
import type { SelectionFormat } from '../../../apps/sheets/src/renderer/selection-format'
import { stepFontSize } from '../../../apps/sheets/src/renderer/font-size-ladder'
import {
  decimalStepValues,
  type DecimalCell,
} from '../../../apps/sheets/src/renderer/decimal-step-model'
import { axisColumnWidthToPixels } from '../../../apps/sheets/src/renderer/normal-font-metrics'
import { CELL_STYLE_PRESETS } from '../../../apps/sheets/src/renderer/cell-style-presets'

export interface RibbonTarget {
  sheetId: string
  address: string
  format: SelectionFormat
  doubleUnderline: boolean
  maximumDigitWidth?: number
  dimensions?: { rows: number; columns: number }
  hyperlinkAddresses?: readonly string[]
  filteredRows?: readonly number[]
  noteAddresses?: readonly string[]
  decimalCells?: readonly DecimalCell[]
  cellIndents?: readonly { address: string; indent: number }[]
}

/** Decode desktop command strings to validated host operations, without touching Univer. */
export function ribbonIntents(command: string, target: RibbonTarget): WorkbookOperation[] {
  const colon = command.indexOf(':')
  const name = colon < 0 ? command : command.slice(0, colon)
  const argument = colon < 0 ? '' : command.slice(colon + 1)
  const { sheetId, address, format: echo } = target
  const bounds = { ...parseRange(address) }
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
    case 'link-set':
    case 'link-remove': {
      const target =
        name === 'link-remove' ? null : normalizeLinkTarget(decodeURIComponent(argument))
      if (name === 'link-set' && target === null) throw new Error('Invalid hyperlink address.')
      result = [
        {
          op: 'set_hyperlink',
          sheetId,
          address: formatAddress(bounds.startRow, bounds.startColumn),
          target,
        },
      ]
      break
    }
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
    case 'indent-step': {
      if (argument !== '1' && argument !== '-1') throw new Error('Invalid indent step')
      if (!target.cellIndents) throw new Error('Cell indentation snapshot is required')
      const expected =
        (bounds.endRow - bounds.startRow + 1) * (bounds.endColumn - bounds.startColumn + 1)
      if (expected > 10_000 || target.cellIndents.length !== expected)
        throw new Error('range_limit_exceeded')
      const seen = new Set<string>()
      result = target.cellIndents.map((cell) => {
        const cellBounds = parseRange(cell.address)
        if (
          cellBounds.startRow !== cellBounds.endRow ||
          cellBounds.startColumn !== cellBounds.endColumn ||
          cellBounds.startRow < bounds.startRow ||
          cellBounds.endRow > bounds.endRow ||
          cellBounds.startColumn < bounds.startColumn ||
          cellBounds.endColumn > bounds.endColumn ||
          !Number.isInteger(cell.indent) ||
          cell.indent < 0 ||
          cell.indent > 250 ||
          seen.has(formatAddress(cellBounds.startRow, cellBounds.startColumn))
        )
          throw new Error('Invalid cell indentation snapshot')
        seen.add(formatAddress(cellBounds.startRow, cellBounds.startColumn))
        return fmt(
          { indent: Math.min(250, Math.max(0, cell.indent + Number(argument))) },
          cell.address,
        )
      })
      break
    }
    case 'decimal-inc':
    case 'decimal-dec': {
      if (!target.decimalCells) throw new Error('Cell number format snapshot is required')
      const expected =
        (bounds.endRow - bounds.startRow + 1) * (bounds.endColumn - bounds.startColumn + 1)
      if (expected > 10_000 || target.decimalCells.length !== expected)
        throw new Error('range_limit_exceeded')
      const seen = new Set<string>()
      for (const cell of target.decimalCells) {
        if (
          !Number.isInteger(cell.row) ||
          !Number.isInteger(cell.col) ||
          cell.row < bounds.startRow ||
          cell.row > bounds.endRow ||
          cell.col < bounds.startColumn ||
          cell.col > bounds.endColumn
        )
          throw new Error('Invalid number format snapshot')
        const key = formatAddress(cell.row, cell.col)
        if (seen.has(key)) throw new Error('Invalid number format snapshot')
        seen.add(key)
      }
      result = decimalStepValues(target.decimalCells, name === 'decimal-inc' ? 1 : -1).map((cell) =>
        fmt({ numberFormat: cell.pattern }, formatAddress(cell.row, cell.col)),
      )
      break
    }
    case 'row-height':
    case 'col-width': {
      const value = Number(argument)
      if (
        !argument.trim() ||
        !Number.isFinite(value) ||
        value < 0 ||
        value > (name === 'row-height' ? 409.5 : 255)
      )
        throw new Error('Invalid axis size')
      const row = name === 'row-height'
      if (!row && (!target.maximumDigitWidth || !Number.isFinite(target.maximumDigitWidth)))
        throw new Error('Workbook font metrics are required')
      const first = row ? bounds.startRow : bounds.startColumn
      const last = row ? bounds.endRow : bounds.endColumn
      const chunk = row ? 500 : 100
      result = []
      for (let start = first; start <= last; start += chunk) {
        const count = Math.min(chunk, last - start + 1)
        result.push(
          row
            ? {
                op: 'set_row_height',
                sheetId,
                row: start + 1,
                count,
                heightPoints: Math.round((value * 4) / 3) * 0.75,
              }
            : {
                op: 'set_col_width',
                sheetId,
                column: columnLabel(start),
                count,
                widthPx: axisColumnWidthToPixels(value, target.maximumDigitWidth!),
                widthCharacters: value,
              },
        )
      }
      break
    }
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
      // Match desktop copy-down/right and the mirrored up/left commands:
      // a one-row/column selection borrows the adjacent source row/column.
      if (vertical && bounds.startRow === bounds.endRow) {
        if (name === 'fill-down') {
          if (bounds.startRow === 0) return []
          bounds.startRow--
        } else {
          if (!target.dimensions) throw new Error('Sheet dimensions are required')
          if (bounds.endRow + 1 >= target.dimensions.rows) return []
          bounds.endRow++
        }
      } else if (!vertical && bounds.startColumn === bounds.endColumn) {
        if (name === 'fill-right') {
          if (bounds.startColumn === 0) return []
          bounds.startColumn--
        } else {
          if (!target.dimensions) throw new Error('Sheet dimensions are required')
          if (bounds.endColumn + 1 >= target.dimensions.columns) return []
          bounds.endColumn++
        }
      }
      const row = name === 'fill-up' ? bounds.endRow : bounds.startRow
      const column = name === 'fill-left' ? bounds.endColumn : bounds.startColumn
      result = [
        {
          op: 'fill_range',
          copyFormat: true,
          sheetId,
          source: vertical
            ? range(row, bounds.startColumn, row, bounds.endColumn)
            : range(bounds.startRow, column, bounds.endRow, column),
          // Keep the donor outside the destination so the range operation's
          // overlap guard and immutable-source semantics remain intact.
          target: range(
            bounds.startRow + (name === 'fill-down' ? 1 : 0),
            bounds.startColumn + (name === 'fill-right' ? 1 : 0),
            bounds.endRow - (name === 'fill-up' ? 1 : 0),
            bounds.endColumn - (name === 'fill-left' ? 1 : 0),
          ),
        },
      ]
      break
    }
    case 'clear-all':
      return ['clear-contents', 'clear-formats', 'clear-comments', 'clear-hyperlinks'].flatMap(
        (command) => ribbonIntents(command, target),
      )
    case 'clear-hyperlinks':
      if (!target.hyperlinkAddresses) throw new Error('Hyperlink snapshot is required')
      if (
        (bounds.endRow - bounds.startRow + 1) * (bounds.endColumn - bounds.startColumn + 1) >
        10_000
      )
        throw new Error('range_limit_exceeded')
      result = target.hyperlinkAddresses.map((linkAddress) => {
        const link = parseRange(linkAddress)
        if (
          link.startRow !== link.endRow ||
          link.startColumn !== link.endColumn ||
          link.startRow < bounds.startRow ||
          link.endRow > bounds.endRow ||
          link.startColumn < bounds.startColumn ||
          link.endColumn > bounds.endColumn
        )
          throw new Error('Invalid hyperlink snapshot')
        return { op: 'set_hyperlink', sheetId, address: linkAddress, target: null }
      })
      break
    case 'clear-formats':
      result = [
        {
          op: 'format_range',
          sheetId,
          range: address,
          reset: true,
          format: { numberFormat: 'General' },
        },
      ]
      break
    case 'clear-comments':
      if (!target.noteAddresses) throw new Error('Note snapshot is required')
      result = target.noteAddresses.map((noteAddress) => {
        const note = parseRange(noteAddress)
        if (
          note.startRow !== note.endRow ||
          note.startColumn !== note.endColumn ||
          note.startRow < bounds.startRow ||
          note.endRow > bounds.endRow ||
          note.startColumn < bounds.startColumn ||
          note.endColumn > bounds.endColumn
        )
          throw new Error('Invalid note snapshot')
        return { op: 'set_note', sheetId, address: noteAddress, text: null }
      })
      break
    case 'clear-contents':
      result = [{ op: 'clear_range', sheetId, range: address }]
      break
    default:
      throw new Error(`Desktop command is not connected yet: ${command}`)
  }
  if ((name === 'clear-formats' || name === 'clear-contents') && target.filteredRows?.length) {
    const filtered = new Set(target.filteredRows)
    const ranges: string[] = []
    let start = bounds.startRow
    for (let row = bounds.startRow; row <= bounds.endRow + 1; row++) {
      if (row <= bounds.endRow && !filtered.has(row)) continue
      if (start < row) ranges.push(range(start, bounds.startColumn, row - 1, bounds.endColumn))
      start = row + 1
    }
    return ranges.flatMap((address) =>
      ribbonIntents(command, { ...target, address, filteredRows: [] }),
    )
  }
  // Validate the whole batch before any command is sent, including dialog batches.
  return result.map((operation) => workbookOperationSchema.parse(operation))
}
