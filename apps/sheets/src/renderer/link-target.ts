/// Desktop dialog input to the persistent hyperlink target.
export function normalizeLinkTarget(input: string): string | null {
  const trimmed = input.trim()
  if (trimmed.length === 0 || trimmed.length > 2083) return null
  if (/^(https?:\/\/|mailto:)/i.test(trimmed)) return trimmed
  if (/^#?'?[^'!]+'?!\$?[A-Za-z]{1,3}\$?[0-9]+$/.test(trimmed)) {
    return trimmed.startsWith('#') ? trimmed : `#${trimmed}`
  }
  if (/^[\w-]+(\.[\w-]+)+(\/\S*)?$/.test(trimmed)) return `https://${trimmed}`
  return null
}
