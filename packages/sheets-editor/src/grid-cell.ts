import {
  BorderStyleTypes,
  HorizontalAlign,
  VerticalAlign,
  WrapStrategy,
  BaselineOffset,
  type ICellData,
  type IStyleData,
} from '@univerjs/core'
import type { cellRecordSchema } from '@genoffice/xlsx-gateway/shared/session-state'
import type { z } from 'zod'

export type ViewCell = z.infer<typeof cellRecordSchema> & { address: string; displayValue?: string }
const borderStyles = {
  thin: BorderStyleTypes.THIN,
  medium: BorderStyleTypes.MEDIUM,
  thick: BorderStyleTypes.THICK,
  dashed: BorderStyleTypes.DASHED,
  dotted: BorderStyleTypes.DOTTED,
  double: BorderStyleTypes.DOUBLE,
  hair: BorderStyleTypes.HAIR,
  dashDot: BorderStyleTypes.DASH_DOT,
  dashDotDot: BorderStyleTypes.DASH_DOT_DOT,
  mediumDashed: BorderStyleTypes.MEDIUM_DASHED,
  mediumDashDot: BorderStyleTypes.MEDIUM_DASH_DOT,
  mediumDashDotDot: BorderStyleTypes.MEDIUM_DASH_DOT_DOT,
  slantDashDot: BorderStyleTypes.SLANT_DASH_DOT,
}

/** Document formatting is independent of the host's light/dark UI theme. */
export function gridCell(cell: Omit<ViewCell, 'address'>): ICellData {
  const style = cell.style ?? {},
    s: IStyleData = {}
  if (style.bold !== undefined) s.bl = style.bold ? 1 : 0
  if (style.italic !== undefined) s.it = style.italic ? 1 : 0
  if (style.underline !== undefined || style.underlineStyle !== undefined)
    s.ul = { s: style.underline === false ? 0 : 1, t: style.underlineStyle === 'double' ? 10 : 0 }
  if (style.vertAlign !== undefined)
    s.va =
      style.vertAlign === 'superscript'
        ? BaselineOffset.SUPERSCRIPT
        : style.vertAlign === 'subscript'
          ? BaselineOffset.SUBSCRIPT
          : BaselineOffset.NORMAL
  if (style.textRotation !== undefined)
    s.tr =
      style.textRotation === 255
        ? { v: 1, a: 0 }
        : { a: style.textRotation > 90 ? 90 - style.textRotation : style.textRotation }
  if (style.indent !== undefined) s.pd = { l: style.indent * 8 }
  if (style.strikethrough !== undefined) s.st = { s: style.strikethrough ? 1 : 0 }
  if (style.fontFamily) s.ff = style.fontFamily
  if (style.fontSize) s.fs = style.fontSize
  if (typeof style.fontColor === 'string') s.cl = { rgb: style.fontColor }
  if (typeof style.fillColor === 'string') s.bg = { rgb: style.fillColor }
  // Empty RGB pins an explicit clear; null is stripped by Univer and lets a
  // row/column's original formatting show through again.
  if (style.fontColor === null) s.cl = { rgb: '' }
  if (style.fillColor === null || style.fill === null) s.bg = { rgb: '' }
  if (style.horizontalAlignment)
    s.ht = {
      left: HorizontalAlign.LEFT,
      center: HorizontalAlign.CENTER,
      right: HorizontalAlign.RIGHT,
      justify: HorizontalAlign.JUSTIFIED,
      distributed: HorizontalAlign.DISTRIBUTED,
    }[style.horizontalAlignment]
  if (style.verticalAlignment)
    s.vt = {
      top: VerticalAlign.TOP,
      center: VerticalAlign.MIDDLE,
      bottom: VerticalAlign.BOTTOM,
    }[style.verticalAlignment]
  if (style.wrapText !== undefined)
    s.tb = style.wrapText ? WrapStrategy.WRAP : WrapStrategy.OVERFLOW
  for (const [property, edge] of [
    ['borderTop', 't'],
    ['borderBottom', 'b'],
    ['borderLeft', 'l'],
    ['borderRight', 'r'],
  ] as const) {
    const border = style[property]
    if (border === undefined) continue
    s.bd ??= {}
    s.bd[edge] =
      border === null
        ? null
        : {
            s: borderStyles[border.style],
            cl: typeof border.color === 'string' ? { rgb: border.color } : {},
          }
  }
  if (cell.numberFormat ?? style.numberFormat)
    s.n = { pattern: (cell.numberFormat ?? style.numberFormat)! }
  return {
    ...(cell.rawValue === null ? {} : { v: cell.rawValue }),
    ...(cell.formula ? { f: cell.formula } : {}),
    ...(Object.keys(s).length ? { s } : {}),
    ...(cell.kind === 'string' ? { t: 4 } : {}),
  }
}
