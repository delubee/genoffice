export const EXCEL_DIGIT_PER_PT: Record<string, number> = {
  Calibri: 7 / 11,
  Verdana: 8 / 10,
  // Korean Excel's default; its GDI digit width matches Calibri (MDW 7px at
  // 11pt), while the macOS substitute (Apple SD Gothic Neo, via the
  // styles.css alias) measures digits ~40% wider — Excel-auto-fitted number
  // columns turned into spurious ####.
  'Malgun Gothic': 7 / 11,
  '맑은 고딕': 7 / 11,
  // MDW 8 at 11pt, solved from Excel print geometry of the ja prod refs
  // (five columns match floor((w+16/256)*8) within 1pt; MDW 7 misses by
  // 20%+). The styles.css alias (Carlito) measures Calibri-like 7.3px.
  'Aptos Narrow': 8 / 11,
  // JIS legacy faces keep em/2 digit advances → 8px at 11pt GDI, matching
  // the grid-derived MS PGothic value (ja Excel's classic 72px default
  // column vs Calibri's 64px).
  'ＭＳ Ｐゴシック': 8 / 11,
  'MS PGothic': 8 / 11,
  'ＭＳ ゴシック': 8 / 11,
  'MS Gothic': 8 / 11,
  'ＭＳ Ｐ明朝': 8 / 11,
  'MS PMincho': 8 / 11,
  'ＭＳ 明朝': 8 / 11,
  'MS Mincho': 8 / 11,
  // Meiryo digits are 0.621em (Verdana design) → 9.1px at 11pt; the ja prod
  // ref grid (B:H) fits MDW 9, while the CJK-name fallback's 8
  // narrows every column 11%. Meiryo UI keeps the same Latin advances.
  メイリオ: 9 / 11,
  Meiryo: 9 / 11,
  'Meiryo UI': 9 / 11,
  // GB/Big5 legacy faces share the em/2 digit advance.
  宋体: 8 / 11,
  SimSun: 8 / 11,
  NSimSun: 8 / 11,
  新細明體: 8 / 11,
  PMingLiU: 8 / 11,
  細明體: 8 / 11,
  MingLiU: 8 / 11,
}

const fontAvailabilityCache = new Map<string, boolean>()

/// Canvas width comparison against generic fallbacks rather than
/// document.fonts.check, whose answer for uninstalled system fonts varies
/// across Chromium versions: a family that renders must differ from at
/// least one generic; a missing one falls back and matches both.
export function fontAvailable(family: string): boolean {
  let known = fontAvailabilityCache.get(family)
  if (known === undefined) {
    known = false
    try {
      const context =
        typeof document === 'undefined' ? null : document.createElement('canvas').getContext('2d')
      if (context) {
        const width = (font: string): number => {
          context.font = `16px ${font}`
          return context.measureText('01mWi').width
        }
        known =
          width(`"${family}", monospace`) !== width('monospace') ||
          width(`"${family}", serif`) !== width('serif')
      }
    } catch {
      known = false
    }
    fontAvailabilityCache.set(family, known)
  }
  return known
}
