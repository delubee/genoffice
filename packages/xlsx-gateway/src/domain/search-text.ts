/** Same scalar text and literal matching used by desktop streamed Find. */
export function scalarToText(value: unknown): string | null {
  if (value === null || value === undefined) return null
  if (typeof value === 'number') return `${value}`
  if (typeof value === 'boolean') return value ? '1' : '0'
  return String(value)
}

export function createSearchTextMatcher(query: {
  query: string
  matchCase: boolean
  matchEntireCell: boolean
}) {
  const needle = (query.matchCase ? query.query : query.query.toLowerCase()).trim()
  return (text: string | null | undefined): boolean => {
    if (!needle || text === null || text === undefined) return false
    const value = query.matchCase ? text : text.toLowerCase()
    return query.matchEntireCell ? value.replace(/^ +| +$/g, '') === needle : value.includes(needle)
  }
}

/** Preserve scalar types when a literal replacement still parses as that type. */
export function coerceReplaceValue(raw: unknown, replaced: string): string | number | boolean {
  if (typeof raw === 'number') {
    if (replaced.trim() === '') return replaced
    const number = Number(replaced)
    return Number.isFinite(number) ? number : replaced
  }
  if (typeof raw === 'boolean') {
    const text = replaced.trim().toLowerCase()
    if (text === '1' || text === 'true') return true
    if (text === '0' || text === 'false') return false
  }
  return replaced
}

/** Literal replacement, including dollar signs; query whitespace follows matching. */
export function replaceSearchText(
  text: string,
  query: { query: string; matchCase: boolean },
  replacement: string,
): string {
  const needle = query.matchCase ? query.query.trim() : query.query.trim().toLowerCase()
  if (!needle) return text
  const haystack = query.matchCase ? text : text.toLowerCase()
  let result = '',
    cursor = 0
  for (;;) {
    const index = haystack.indexOf(needle, cursor)
    if (index < 0) return result + text.slice(cursor)
    result += text.slice(cursor, index) + replacement
    cursor = index + needle.length
  }
}
