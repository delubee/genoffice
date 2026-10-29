import { BooleanNumber, type IRange } from '@univerjs/core'

/** newest reveal wins: rapid Find Next/Previous must not let an older
 *  correction loop keep scrolling after a newer one started */
let revealGeneration = 0

function isRtlFacadeSheet(sheet: {
  getSheet?(): { getConfig(): { rightToLeft?: BooleanNumber } }
}): boolean {
  try {
    return sheet.getSheet?.().getConfig().rightToLeft === BooleanNumber.TRUE
  } catch {
    return false
  }
}

export async function revealCellBelowFreeze(
  sheet: {
    scrollToCell(row: number, column: number): unknown
    getVisibleRange(): IRange | null
    getSheet?(): { getConfig(): { rightToLeft?: BooleanNumber } }
  },
  row: number,
  column: number,
): Promise<void> {
  // aim one row above the target: the offset error is fractional rows, so a
  // position converged exactly on the target can still leave it half-sliced
  // under the pane — with the predecessor as the boundary row the target is
  // whole. Every scrollToCell call re-applies the same broken offset, so the
  // compensation folds into one running aim; the scroll and its viewport
  // update settle asynchronously, hence the waits between measurements.
  const generation = ++revealGeneration
  if (isRtlFacadeSheet(sheet)) {
    // On an RTL sheet scrollToCell parks the target at the viewport's
    // visually-left (max-column) edge — a full jump, not a reveal. Univer's
    // own minimal reveal (rtl-grid-mirror) has usually just made the target
    // visible; once that scroll has rendered, keep its position.
    await new Promise((resolve) => setTimeout(resolve, 60))
    if (generation !== revealGeneration) return
    let visible: IRange | null
    try {
      visible = sheet.getVisibleRange()
    } catch {
      return
    }
    if (
      visible &&
      row >= visible.startRow &&
      row <= visible.endRow &&
      column >= visible.startColumn &&
      column <= visible.endColumn
    )
      return
  }
  // Already fully visible → nothing to reveal. getVisibleRange starts below
  // the frozen pane, so a target hidden under it reads as out of range and
  // still gets the corrective scroll. Without this short-circuit every
  // re-emission of an unchanged match (find research after streamed patches)
  // re-scrolls the viewport.
  try {
    const visible = sheet.getVisibleRange()
    if (
      visible &&
      row >= visible.startRow &&
      row <= visible.endRow &&
      column >= visible.startColumn &&
      column <= visible.endColumn
    ) {
      return
    }
  } catch {
    // no scroll render controller yet — proceed with the reveal
  }
  const aimRow = Math.max(0, row - 1)
  let scrollRow = aimRow
  let scrollColumn = column
  let lastStartRow = -1
  let lastStartColumn = -1
  sheet.scrollToCell(scrollRow, scrollColumn)
  for (let attempt = 0; attempt < 4; attempt += 1) {
    await new Promise((resolve) => setTimeout(resolve, 60))
    if (generation !== revealGeneration) return // superseded by a newer reveal
    let visible: IRange | null
    try {
      visible = sheet.getVisibleRange()
    } catch {
      return
    }
    if (!visible) return
    // over*: the viewport starts past the aim (target hidden under the pane).
    // under*: a correction clamped at 0 left the target below/right of the
    // viewport — the aim being visible is not enough, the target must be too.
    // The two cannot both be positive: overRows > 0 puts the whole viewport
    // at or past the target row.
    const overRows = Math.max(0, visible.startRow - aimRow)
    const underRows = Math.max(0, row - visible.endRow)
    const overColumns = Math.max(0, visible.startColumn - column)
    const underColumns = Math.max(0, column - visible.endColumn)
    if (overRows === 0 && overColumns === 0 && underRows === 0 && underColumns === 0) return
    // an aim inside the frozen pane (target on the first scrollable row) can
    // never be reached — the viewport clamps at the pane edge with the target
    // visible right below it, so a correction that moved nothing is done
    if (visible.startRow === lastStartRow && visible.startColumn === lastStartColumn) return
    lastStartRow = visible.startRow
    lastStartColumn = visible.startColumn
    scrollRow = Math.max(0, scrollRow - overRows + underRows)
    scrollColumn = Math.max(0, scrollColumn - overColumns + underColumns)
    sheet.scrollToCell(scrollRow, scrollColumn)
  }
}
