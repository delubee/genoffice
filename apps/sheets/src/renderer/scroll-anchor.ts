/**
 * The viewport's scroll anchor as scrollToCell arguments: the scroll state's
 * sheetViewStartRow/Column plus the freeze split round-trips exactly, and on
 * RTL sheets it is the only correct source — getVisibleRange().startColumn is
 * the logically LOWEST visible column there, while scrollToCell (and the lazy
 * streaming anchor) interpret the start column as the viewport's visually-left
 * (highest) logical index.
 */
export function getScrollAnchor(
  workbook: {
    getScrollStateBySheetId(
      sheetId: string,
    ): { sheetViewStartRow?: number; sheetViewStartColumn?: number } | null | undefined | void
  },
  sheet: {
    getSheetId(): string
    getSheet(): { getFreeze(): { xSplit: number; ySplit: number } }
  },
): { row: number; column: number } | null {
  try {
    const scroll = workbook.getScrollStateBySheetId(sheet.getSheetId())
    if (scroll?.sheetViewStartRow == null || scroll.sheetViewStartColumn == null) return null
    const freeze = sheet.getSheet().getFreeze()
    return {
      row: scroll.sheetViewStartRow + freeze.ySplit,
      column: scroll.sheetViewStartColumn + freeze.xSplit,
    }
  } catch {
    // No scroll render controller yet (still booting).
    return null
  }
}
