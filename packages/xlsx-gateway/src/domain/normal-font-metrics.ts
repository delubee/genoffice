import { EXCEL_DIGIT_PER_PT, fontAvailable } from './font-digit-metrics'

export interface NormalFontDescriptor {
  styles?: readonly {
    fontFamily?: string | undefined
    fontSize?: number | undefined
    fontScheme?: 'major' | 'minor' | undefined
  }[]
  normalFontName?: string | undefined
  themeFonts?: { minorEa?: string | undefined } | undefined
}

const CJK_FAMILY_NAME =
  /[\u1100-\u11ff\u3000-\u303f\u3040-\u30ff\u3130-\u318f\u3400-\u4dbf\u4e00-\u9fff\uac00-\ud7af\uf900-\ufaff\uff00-\uffef]/

const clampMdw = (mdw: number): number => Math.max(4, Math.min(30, mdw))

/// The face Excel derives the column-width MDW from. For scheme fonts the
/// sidecar substitutes the theme latin face into styles[0], but Excel lays
/// CJK workbooks out per the Normal font's author-locale resolution — the
/// literal cached <name val> (e.g. MS PGothic under theme latin Calibri) —
/// with the theme's minor <a:ea> face as the fallback when no name is
/// cached. Only trust a differing literal that names a face we know.
export function resolveNormalMdwFamily(file: NormalFontDescriptor): string {
  const normal = file.styles?.[0]
  const themeResolved = normal?.fontFamily ?? 'Calibri'
  // ja vertical-text names prefix the base face with '@'.
  let literal = (file.normalFontName ?? '').replace(/^@/, '')
  if (literal === '' && normal?.fontScheme === 'minor') {
    literal = file.themeFonts?.minorEa ?? ''
  }
  if (literal === '' || literal === themeResolved) return themeResolved
  if (
    EXCEL_DIGIT_PER_PT[literal] !== undefined ||
    CJK_FAMILY_NAME.test(literal) ||
    fontAvailable(literal)
  ) {
    return literal
  }
  return themeResolved
}

const MAC_EXCEL_MDW_8_AT_11 = new Set(['Calibri', 'Malgun Gothic', '맑은 고딕'])

export function measureNormalFontMdw(file: NormalFontDescriptor): number {
  const size = file.styles?.[0]?.fontSize ?? 11
  const family = resolveNormalMdwFamily(file)
  // Mac Excel (the fidelity reference) lays Calibri 11 out at MDW 8, not the
  // Windows GDI 7: live probes and ref print geometry both fit
  // floor((w+16/256)*8) within 1pt across corpora (50.86ch→305pt,
  // 32.44ch→195pt, 13.71ch→82pt) while MDW 7
  // misses by 12%+, wrapping borderline text a line early. Malgun Gothic 11
  // (Korean Excel's default) prints the same way: a reference grid at
  // scale 64 gives 112/89/137/101/113 px for 14/11.25/17.125/12.625/14.125
  // ch — floor((w+16/256)*8) exactly, MDW 7 short by 12%.
  if (MAC_EXCEL_MDW_8_AT_11.has(family) && size === 11) return 8
  const perPt = EXCEL_DIGIT_PER_PT[family]
  if (perPt !== undefined) return clampMdw(Math.round(perPt * size))
  // CJK faces without a table entry: canvas would measure a macOS
  // substitute; every grounded legacy CJK face is em/2 → 8px at 11pt.
  if (CJK_FAMILY_NAME.test(family)) return clampMdw(Math.round((8 / 11) * size))
  if (typeof document === 'undefined') return 7
  try {
    const context = document.createElement('canvas').getContext('2d')
    if (!context) return 7
    // Quote the family: an unquoted leading digit or '@' silently no-ops
    // the whole ctx.font assignment and measures the 10px canvas default.
    context.font = `${(size * 96) / 72}px "${family.replace(/["\\]/g, '')}"`
    const width = context.measureText('0').width
    // Round with a 0.4 threshold: Arial 10 measures 7.42px and prints at
    // MDW 8 (plain round loses 12% of every column), while Arial 11 (8.16)
    // and Century Gothic 11 (8.13) print at 8 (ceil overshoots them 12% the
    // other way). Any threshold in [0.58, 0.84) fits all calibrated refs.
    return width > 0 ? clampMdw(Math.floor(width + 0.6)) : 7
  } catch {
    return 7
  }
}

export function axisColumnWidthToPixels(width: number, mdw: number): number {
  if (width === 0) return 0
  const mdwPt = mdw * 0.75
  return Math.round((Math.trunc(((256 * width + Math.trunc(128 / mdwPt)) / 256) * mdwPt) * 4) / 3)
}
