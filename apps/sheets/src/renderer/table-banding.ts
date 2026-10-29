/**
 * Excel table-style banding painted into cell fills (header band, row and
 * column stripes, first/last column, totals and borders), shared by the
 * desktop loader and the VS Code projection.
 */
import {
  BooleanNumber,
  BorderStyleTypes,
  type ICellData,
  type IRange,
  type IStyleData,
} from '@univerjs/core'
import type { WorkbookFile } from '../shared/desktop-api'

export interface HiddenRowsInfo {
  readonly rows: ReadonlySet<number>
  readonly coveredThrough: number
}

export function mapBorderStyle(style: string): BorderStyleTypes {
  switch (style) {
    case 'hair':
      return BorderStyleTypes.HAIR
    case 'dotted':
      return BorderStyleTypes.DOTTED
    case 'dashed':
      return BorderStyleTypes.DASHED
    case 'dashDot':
      return BorderStyleTypes.DASH_DOT
    case 'dashDotDot':
      return BorderStyleTypes.DASH_DOT_DOT
    case 'double':
      return BorderStyleTypes.DOUBLE
    case 'medium':
      return BorderStyleTypes.MEDIUM
    case 'mediumDashed':
      return BorderStyleTypes.MEDIUM_DASHED
    case 'mediumDashDot':
      return BorderStyleTypes.MEDIUM_DASH_DOT
    case 'mediumDashDotDot':
      return BorderStyleTypes.MEDIUM_DASH_DOT_DOT
    case 'slantDashDot':
      return BorderStyleTypes.SLANT_DASH_DOT
    case 'thick':
      return BorderStyleTypes.THICK
    default:
      return BorderStyleTypes.THIN
  }
}

export function styleHasFill(style: IStyleData): boolean {
  return Boolean((style.bg as { rgb?: string } | null | undefined)?.rgb)
}

/// Approximates Excel table styles (header band + row stripes) for cells that
/// carry no explicit fill of their own.
export function applyTableBanding(
  matrix: ICellData[][],
  range: IRange,
  tables: WorkbookFile['sheets'][number]['tables'],
  hiddenRows?: HiddenRowsInfo,
): void {
  for (const table of tables) {
    const rowStart = Math.max(range.startRow, table.range.startRow)
    const rowEnd = Math.min(range.endRow, table.range.endRow)
    const columnStart = Math.max(range.startColumn, table.range.startColumn)
    const columnEnd = Math.min(range.endColumn, table.range.endColumn)
    if (rowStart > rowEnd || columnStart > columnEnd) continue
    // A name-less tableStyleInfo is Excel's style "None": paint nothing.
    if (!table.styleName && !table.headerFill && !table.headerFontColor && !table.stripeFill) {
      continue
    }
    // Colors are resolved sidecar-side from the workbook's real theme accents
    // (Light/Medium/Dark variant rules) or the file's custom <tableStyle>
    // dxfs; the literals are a last-resort fallback.
    const headerFill = table.headerFill
    const headerFont = table.headerFontColor ?? '#FFFFFF'
    // No stripe color means the style genuinely has none (custom styles
    // without band dxfs, Light 8-14) — a fallback would invent banding.
    const stripeFill = table.stripeFill
    const dataStartRow = table.range.startRow + table.headerRowCount
    const totalsStartRow = table.range.endRow - (table.totalsRowCount ?? 0) + 1
    // A live autoFilter makes Excel re-rank the stripes by VISIBLE row order
    // (a filtered Light19 ref alternates across the hidden gaps); manually
    // hidden rows keep the physical banding, so only filtered tables re-rank.
    // The scan needs full hidden knowledge from the table's top: past the
    // streamed coverage an absent row is merely unknown, not visible.
    let visibleParity: Map<number, number> | undefined
    const scanEnd = Math.min(rowEnd, totalsStartRow - 1)
    if (
      table.filterActive &&
      table.showRowStripes &&
      hiddenRows &&
      hiddenRows.rows.size > 0 &&
      scanEnd <= hiddenRows.coveredThrough
    ) {
      visibleParity = new Map()
      let ordinal = 0
      for (let row = dataStartRow; row <= scanEnd; row += 1) {
        if (hiddenRows.rows.has(row)) continue
        if (row >= rowStart) visibleParity.set(row, ordinal % 2)
        ordinal += 1
      }
    }
    for (let row = rowStart; row <= rowEnd; row += 1) {
      const isHeader = row < dataStartRow
      const isTotals = !isHeader && row >= totalsStartRow
      // Excel's firstRowStripe covers the FIRST data row (ref: Medium9 shades
      // data row 1 with #B8CCE4), then alternates with secondRowStripe.
      const rowParity = visibleParity?.get(row) ?? (row - dataStartRow) % 2
      const isStripe = !isHeader && !isTotals && table.showRowStripes && rowParity === 0
      const secondStripeFill =
        !isHeader && !isTotals && table.showRowStripes && rowParity === 1
          ? table.secondRowStripeFill
          : undefined
      for (let column = columnStart; column <= columnEnd; column += 1) {
        const cell = matrix[row - range.startRow]?.[column - range.startColumn]
        if (!cell) continue
        let style = (cell.s ?? {}) as IStyleData
        const hasCustomBorders =
          table.wholeTableBorderColor !== undefined ||
          table.innerHorizontalBorderColor !== undefined ||
          table.innerVerticalBorderColor !== undefined ||
          table.headerBottomBorderColor !== undefined
        if (table.borderColor || hasCustomBorders || (isTotals && table.totalRowBorderColor)) {
          const edges: IStyleData['bd'] = {}
          if (table.borderColor) {
            if (row === table.range.startRow) {
              edges.t = { s: BorderStyleTypes.MEDIUM, cl: { rgb: table.borderColor } }
            }
            if (isHeader && row === dataStartRow - 1) {
              edges.b = { s: BorderStyleTypes.THIN, cl: { rgb: table.borderColor } }
            }
            if (row === table.range.endRow) {
              edges.b = { s: BorderStyleTypes.MEDIUM, cl: { rgb: table.borderColor } }
            }
          }
          // Custom wholeTable dxf borders: inner grid first, the header rule
          // over it, the outline last — later edges win shared boundaries.
          if (table.innerHorizontalBorderColor && row < table.range.endRow) {
            edges.b = {
              s: mapBorderStyle(table.innerHorizontalBorderStyle ?? 'thin'),
              cl: { rgb: table.innerHorizontalBorderColor },
            }
          }
          if (table.innerVerticalBorderColor && column < table.range.endColumn) {
            edges.r = {
              s: mapBorderStyle(table.innerVerticalBorderStyle ?? 'thin'),
              cl: { rgb: table.innerVerticalBorderColor },
            }
          }
          if (table.headerBottomBorderColor && isHeader && row === dataStartRow - 1) {
            edges.b = {
              s: mapBorderStyle(table.headerBottomBorderStyle ?? 'thin'),
              cl: { rgb: table.headerBottomBorderColor },
            }
          }
          if (table.wholeTableBorderColor) {
            const outline = {
              s: mapBorderStyle(table.wholeTableBorderStyle ?? 'thin'),
              cl: { rgb: table.wholeTableBorderColor },
            }
            if (row === table.range.startRow) edges.t = outline
            if (row === table.range.endRow) edges.b = outline
            if (column === table.range.startColumn) edges.l = outline
            if (column === table.range.endColumn) edges.r = outline
          }
          if (isTotals && row === totalsStartRow && table.totalRowBorderColor) {
            edges.t = {
              s: mapBorderStyle(table.totalRowBorderStyle ?? 'thin'),
              cl: { rgb: table.totalRowBorderColor },
            }
          }
          if (edges.t || edges.b || edges.l || edges.r) {
            cell.s = { ...style, bd: { ...(style.bd ?? {}), ...edges } }
            style = cell.s as IStyleData
          }
        }
        if (isHeader) {
          const fontColor =
            column === table.range.startColumn && table.firstHeaderCellFontColor
              ? table.firstHeaderCellFontColor
              : headerFill
                ? headerFont
                : (table.headerFontColor ?? '#333333')
          if (styleHasFill(style)) {
            // Baked header fill: keep it, but a default-black font still takes
            // the style's header font (Excel lets table-style text win over
            // the automatic color).
            const cellFont = (style.cl as { rgb?: string } | undefined)?.rgb
            if (headerFill && (!cellFont || cellFont === '#000000')) {
              cell.s = { ...style, cl: { rgb: fontColor }, bl: BooleanNumber.TRUE }
            }
            continue
          }
          // An explicit non-automatic cell font color survives the table
          // style (Book1_custom's red "Names" header).
          const explicitFont = (style.cl as { rgb?: string } | undefined)?.rgb
          cell.s = {
            ...style,
            ...(headerFill ? { bg: { rgb: headerFill } } : {}),
            ...(explicitFont && explicitFont !== '#000000' ? {} : { cl: { rgb: fontColor } }),
            bl: BooleanNumber.TRUE,
          }
          continue
        }
        if (styleHasFill(style)) continue
        if (isTotals) {
          cell.s = {
            ...style,
            ...(table.totalRowFill ? { bg: { rgb: table.totalRowFill } } : {}),
            ...(table.totalRowFontColor ? { cl: { rgb: table.totalRowFontColor } } : {}),
            bl: BooleanNumber.TRUE,
          }
          continue
        }
        // Band precedence below header/totals: first/last column emphasis,
        // then row stripes, then column stripes, then the whole-table fill.
        const isFirstColumn = column === table.range.startColumn && table.firstColumnFill
        const isLastColumn = column === table.range.endColumn && table.lastColumnFill
        const columnStripeFill = table.showColumnStripes
          ? (column - table.range.startColumn) % 2 === 0
            ? table.columnStripeFill
            : table.secondColumnStripeFill
          : undefined
        const fill = isFirstColumn
          ? table.firstColumnFill
          : isLastColumn
            ? table.lastColumnFill
            : ((isStripe ? stripeFill : undefined) ??
              secondStripeFill ??
              columnStripeFill ??
              table.wholeTableFill)
        if (fill) {
          // Dark families set a body text color; a default-black font yields
          // to it (explicit cell colors survive, mirroring the header rule).
          const cellFont = (style.cl as { rgb?: string } | undefined)?.rgb
          const fontPatch =
            table.bodyFontColor && (!cellFont || cellFont === '#000000')
              ? { cl: { rgb: table.bodyFontColor } }
              : {}
          cell.s = { ...style, bg: { rgb: fill }, ...fontPatch }
        }
      }
    }
  }
}
