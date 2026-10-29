/** Excel's sheet-name cap; the save path's sheetNameSchema enforces the same limit. */
const SHEET_NAME_MAX = 31

/**
 * "Name" → "Name (2)" → "Name (3)" against the taken set (case-insensitive like
 * Excel). A suffixed candidate stays within Excel's 31-character cap: the base
 * is trimmed to make room (a 30-char name that collides becomes "…26 chars (2)"),
 * since an over-long name would pass the grid but fail the whole workbook save.
 */
export function dedupeSheetName(name: string, taken: ReadonlySet<string>): string {
  const lower = new Set([...taken].map((entry) => entry.toLowerCase()))
  if (!lower.has(name.toLowerCase())) return name
  for (let n = 2; ; n += 1) {
    const suffix = ` (${n})`
    // the cut may expose trailing whitespace or an apostrophe (illegal at a sheet name's end)
    const base = name.slice(0, SHEET_NAME_MAX - suffix.length).replace(/[\s']+$/, '')
    const candidate = `${base}${suffix}`
    if (!lower.has(candidate.toLowerCase())) return candidate
  }
}
