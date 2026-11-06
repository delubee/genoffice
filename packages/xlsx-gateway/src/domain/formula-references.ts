import { FORMULA_REFERENCE_PATTERN } from '../gateway/xlsx-structure'

interface ParsedToken {
  readonly startRow: number | null
  readonly endRow: number | null
  readonly startColumn: number | null
  readonly endColumn: number | null
}

function lettersToColumn(letters: string): number {
  let column = 0
  for (const character of letters) {
    column = column * 26 + character.charCodeAt(0) - 64
  }
  return column - 1
}

/// Null bounds mean "whole axis" (A:A or 1:5 references), clamped to the
/// target sheet's used range by the caller.
function parseReferenceToken(token: string): ParsedToken | null {
  const wholeColumn = /^\$?([A-Z]{1,3}):\$?([A-Z]{1,3})$/.exec(token)
  if (wholeColumn) {
    const start = lettersToColumn(wholeColumn[1] ?? 'A')
    const end = lettersToColumn(wholeColumn[2] ?? 'A')
    return {
      startRow: null,
      endRow: null,
      startColumn: Math.min(start, end),
      endColumn: Math.max(start, end),
    }
  }
  const wholeRow = /^\$?([0-9]+):\$?([0-9]+)$/.exec(token)
  if (wholeRow) {
    const start = Number(wholeRow[1]) - 1
    const end = Number(wholeRow[2]) - 1
    return {
      startRow: Math.min(start, end),
      endRow: Math.max(start, end),
      startColumn: null,
      endColumn: null,
    }
  }
  const cells = token.split(':').map((part) => /^\$?([A-Z]{1,3})\$?([0-9]+)$/.exec(part))
  if (cells.some((cell) => cell === null)) return null
  const rows = cells.map((cell) => Number(cell?.[2]) - 1)
  const columns = cells.map((cell) => lettersToColumn(cell?.[1] ?? 'A'))
  return {
    startRow: Math.min(...rows),
    endRow: Math.max(...rows),
    startColumn: Math.min(...columns),
    endColumn: Math.max(...columns),
  }
}

interface FormulaReference {
  readonly qualifier: string | undefined
  readonly token: ParsedToken
}

export function parseFormulaReferences(formula: string): FormulaReference[] {
  const references: FormulaReference[] = []
  const segments = formula.split('"')
  for (let index = 0; index < segments.length; index += 2) {
    for (const match of (segments[index] ?? '').matchAll(FORMULA_REFERENCE_PATTERN)) {
      const token = parseReferenceToken(match[3] ?? '')
      if (token) references.push({ qualifier: match[2], token })
    }
  }
  return references
}
