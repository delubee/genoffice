/** OOXML order: only edited anchors permute among their existing slots. */
export function projectVisualOrder<
  T extends {
    sheetId: string
    drawingPath?: string | undefined
    drawingIndex?: number | undefined
  },
>(projected: readonly T[], key: (visual: T) => number | undefined): T[] {
  const byPart = new Map<string, number[]>()
  projected.forEach((visual, at) => {
    if (key(visual) === undefined) return
    const part = visual.drawingPath ?? visual.sheetId
    byPart.set(part, [...(byPart.get(part) ?? []), at])
  })
  if (byPart.size === 0) return [...projected]
  const result = [...projected]
  for (const slots of byPart.values()) {
    const sorted = slots
      .map((at) => projected[at]!)
      .sort(
        (left, right) =>
          key(left)! - key(right)! || (left.drawingIndex ?? 0) - (right.drawingIndex ?? 0),
      )
    slots.forEach((at, index) => {
      result[at] = sorted[index]!
    })
  }
  return result
}
