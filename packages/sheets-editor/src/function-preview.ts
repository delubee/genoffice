import {
  containsUnresolvedNames,
  parseFormulaReferences,
} from '../../../apps/sheets/src/renderer/formula-closure'
import type { GridRange } from './ports'

/** Resolve bounded static precedents before asking the display engine for a preview. */
export function functionPreviewRanges(
  expression: string,
  ownerId: string,
  sheets: readonly { id: string; name: string; rows: number; columns: number }[],
): { sheetId: string; range: GridRange }[] | null {
  if (containsUnresolvedNames(expression) || /\b(?:INDIRECT|OFFSET)\s*\(/i.test(expression))
    return null
  const result: { sheetId: string; range: GridRange }[] = []
  let cells = 0
  for (const { qualifier, token } of parseFormulaReferences(expression)) {
    const name = qualifier?.replace(/^'|'$/g, '').replace(/''/g, "'")
    const sheet =
      name === undefined
        ? sheets.find((sheet) => sheet.id === ownerId)
        : sheets.find((sheet) => sheet.name.toLowerCase() === name.toLowerCase())
    if (!sheet) return null
    const range = {
      startRow: token.startRow ?? 0,
      endRow: token.endRow ?? Math.max(0, sheet.rows - 1),
      startColumn: token.startColumn ?? 0,
      endColumn: token.endColumn ?? Math.max(0, sheet.columns - 1),
    }
    if (
      range.startRow < 0 ||
      range.startColumn < 0 ||
      range.endRow >= 1048576 ||
      range.endColumn >= 16384
    )
      return null
    cells += (range.endRow - range.startRow + 1) * (range.endColumn - range.startColumn + 1)
    if (cells > 2000) return null
    result.push({ sheetId: sheet.id, range })
  }
  return result
}
