/** Excel-style default name for an inserted sheet: the lowest free SheetN. */
export function nextSheetName(taken: readonly string[]): string {
  const lower = new Set(taken.map((name) => name.toLowerCase()))
  for (let n = 1; ; n += 1) {
    const candidate = `Sheet${n}`
    if (!lower.has(candidate.toLowerCase())) return candidate
  }
}
