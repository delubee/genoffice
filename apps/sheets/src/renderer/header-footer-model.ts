import type { HeaderFooterParts } from './ribbon-model'

/// Formatting toggles and codes the layout cannot render; any other &X is literal text.
const STRIPPED_CODES = new Set(['B', 'I', 'U', 'S', 'E', 'X', 'Y', 'Z', 'O', 'H'])

/// Excel's encoded header/footer → left/center/right parts. Field codes the
/// layout resolves (&P &N &D &T &F &A &G picture, && literal) stay verbatim;
/// formatting codes (font/size/color/bold/…) and unsupported codes (&Z path)
/// are stripped, while an unrecognised &X is kept as the literal text Excel
/// prints. Text before the first section marker is centered.
export function decodeHeaderFooter(encoded: string): HeaderFooterParts | null {
  const sections = { L: '', C: '', R: '' }
  let current: 'L' | 'C' | 'R' = 'C'
  let index = 0
  while (index < encoded.length) {
    const character = encoded[index] ?? ''
    if (character !== '&') {
      sections[current] += character
      index += 1
      continue
    }
    const next = encoded[index + 1]
    if (next === undefined) break
    if (next === '&') {
      sections[current] += '&&'
      index += 2
      continue
    }
    // &"Font Name,Style" — skip the whole quoted spec.
    if (next === '"') {
      const close = encoded.indexOf('"', index + 2)
      index = close === -1 ? encoded.length : close + 1
      continue
    }
    // &nn font size.
    if (/[0-9]/.test(next)) {
      index += 2
      while (index < encoded.length && /[0-9]/.test(encoded[index] ?? '')) index += 1
      continue
    }
    const code = next.toUpperCase()
    // &K followed by an RRGGBB or theme color spec.
    if (code === 'K') {
      index += 8
      continue
    }
    if (code === 'L' || code === 'C' || code === 'R') {
      current = code
      index += 2
      continue
    }
    if (
      code === 'P' ||
      code === 'N' ||
      code === 'D' ||
      code === 'T' ||
      code === 'F' ||
      code === 'A' ||
      code === 'G'
    ) {
      sections[current] += `&${code}`
      index += 2
      continue
    }
    if (STRIPPED_CODES.has(code)) {
      index += 2
      continue
    }
    sections[current] += `&${code}`
    index += 2
  }
  if (sections.L === '' && sections.C === '' && sections.R === '') return null
  return {
    ...(sections.L === '' ? {} : { left: sections.L }),
    ...(sections.C === '' ? {} : { center: sections.C }),
    ...(sections.R === '' ? {} : { right: sections.R }),
  }
}
