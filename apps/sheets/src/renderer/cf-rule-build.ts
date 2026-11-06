/**
 * Conditional-formatting rule builders shared by the desktop projection and
 * the VS Code workbook view: file rules and dxf styles to Univer rules.
 */
import { mapBorderStyle } from './table-banding'
import {
  BooleanNumber,
  BorderStyleTypes,
  HorizontalAlign,
  type IRange,
  type IStyleData,
  VerticalAlign,
  WrapStrategy,
  BaselineOffset,
} from '@univerjs/core'
import { CFValueType, type IValueConfig } from '@univerjs/preset-sheets-conditional-formatting'
import { WORST_FIRST_ICON_SETS } from '@genoffice/xlsx-gateway/gateway/xlsx-cf'
import type { WorkbookCellStyle, WorkbookRangeResult } from '../shared/desktop-api'
import { escapeCssLeadingDigit, NO_FILL_STYLE, ooxmlTextRotationToUniver } from './edit-journal'
import { SHRINK_TO_FIT_KEY } from './shrink-to-fit-key'
import { CELL_HIDDEN_KEY, CELL_LOCKED_KEY } from './sheet-protection'
import { INDENT_STEP_PX } from './selection-format'
import { type UniverWorksheet } from './univer-state'

export type CfHighlightBuilder = ReturnType<
  ReturnType<UniverWorksheet['newConditionalFormattingRule']>['whenCellNotEmpty']
>

export function buildConditionalRule(
  worksheet: UniverWorksheet,
  dxfStyles: readonly WorkbookCellStyle[],
  rule: WorkbookRangeResult['conditionalRules'][number],
) {
  const ranges: IRange[] = rule.ranges.map((area) => ({
    startRow: area.startRow,
    startColumn: area.startColumn,
    endRow: area.endRow,
    endColumn: area.endColumn,
  }))
  const builder = worksheet.newConditionalFormattingRule()
  if (rule.ruleType === 'colorScale') {
    if (rule.colors.length < 2 || rule.cfvos.length !== rule.colors.length) return null
    return builder
      .setColorScale(
        rule.cfvos.map((cfvo, index) => ({
          index,
          color: rule.colors[index] ?? '#FFFFFF',
          value: toCfValue(cfvo),
        })),
      )
      .setRanges(ranges)
      .build()
  }
  if (rule.ruleType === 'dataBar') {
    const [min, max] = rule.cfvos
    if (!min || !max) return null
    const positive = rule.colors[0] ?? '#638EC6'
    return builder
      .setDataBar({
        min: toCfValue(min),
        max: toCfValue(max),
        positiveColor: positive,
        // Explicit x14 negative fill wins; an x14 twin flagged
        // same-as-positive reuses the positive fill; otherwise Excel's
        // default negative fill is red.
        nativeColor: rule.negativeColor ?? (rule.negativeSameAsPositive ? positive : '#FF0000'),
        isShowValue: rule.showValue,
        isGradient: rule.gradient ?? true,
      })
      .setRanges(ranges)
      .build()
  }
  if (rule.ruleType === 'iconSet') {
    if (rule.cfvos.length < 3) return null
    const iconType = rule.iconSetName ?? '3TrafficLights1'
    const count = rule.cfvos.length
    // xlsx cfvos are ascending thresholds (first one is the catch-all minimum);
    // Univer wants a descending greaterThanOrEqual chain per icon.
    const worstFirst = WORST_FIRST_ICON_SETS.has(iconType)
    const configs = []
    for (let index = count - 1; index >= 0; index -= 1) {
      const cfvo = rule.cfvos[index]
      if (!cfvo) return null
      // The file's icon order runs worst-first; Univer's iconMap runs
      // best-first except for the rating sets.
      const fileIcon = rule.iconReverse ? count - 1 - index : index
      const iconIndex = worstFirst ? fileIcon : count - 1 - fileIcon
      configs.push({
        iconType,
        iconId: String(iconIndex),
        operator: index > 0 && cfvo.gte === false ? 'greaterThan' : 'greaterThanOrEqual',
        value: index === 0 ? { type: CFValueType.min } : toCfValue(cfvo),
      })
    }
    return builder
      .setIconSet({
        iconConfigs: configs as Parameters<typeof builder.setIconSet>[0]['iconConfigs'],
        isShowValue: rule.showValue,
      })
      .setRanges(ranges)
      .build()
  }
  // Univer offsets relative CF formulas from the top-left-sorted first range,
  // not the file's sqref order.
  const first = [...ranges].sort(
    (a, b) => a.startRow - b.startRow || a.startColumn - b.startColumn,
  )[0]
  const anchor = first ? `${columnLetter(first.startColumn)}${first.startRow + 1}` : 'A1'
  const coveredCells = ranges.reduce(
    (sum, r) => sum + (r.endRow - r.startRow + 1) * (r.endColumn - r.startColumn + 1),
    0,
  )
  const highlight = buildHighlightCondition(builder, rule, anchor, coveredCells)
  if (!highlight) return null
  const built = applyDxfFormat(highlight, dxfStyles, rule.dxfIndex).setRanges(ranges).build()
  patchBuiltHighlightRule(
    built.rule as BuiltHighlightRule,
    rule.ruleType,
    rule.dxfIndex === undefined ? undefined : dxfStyles[rule.dxfIndex],
  )
  // The facade builder has no setter for cfRule/@stopIfTrue; Univer's
  // evaluator honors the flag on the installed rule (matched → skip
  // lower-priority rules). Only highlight rules carry it — Excel's UI does
  // not offer stop-if-true on color scales / data bars / icon sets.
  if (rule.stopIfTrue === true) {
    ;(built as { stopIfTrue?: boolean }).stopIfTrue = true
  }
  return built
}

export type BuiltHighlightRule = {
  operator?: string
  value?: string
  /// The builder's IStyleBase plus the patched-on keys.
  style?: {
    n?: { pattern: string }
    bd?: Partial<Record<'t' | 'b' | 'l' | 'r', ReturnType<typeof toUniverBorder>>>
    [key: string]: unknown
  }
}

/// IHighlightCell supports more than the facade builder exposes: style is an
/// IStyleBase (the paint path merges its n and bd verbatim), and the
/// error-presence text operators have no setter at all. Patch the built rule
/// for a dxf number format, dxf borders, and containsErrors /
/// notContainsErrors (built as a placeholder text condition in
/// buildHighlightCondition).
export function patchBuiltHighlightRule(
  target: BuiltHighlightRule,
  ruleType: string,
  dxf: WorkbookCellStyle | undefined,
): void {
  if (ruleType === 'containsErrors' || ruleType === 'notContainsErrors') {
    target.operator = ruleType
    delete target.value
  }
  if (!dxf) return
  // A numFmt-only dxf builds with no style object at all.
  if (dxf.numberFormat) {
    target.style = { ...target.style, n: { pattern: dxf.numberFormat } }
  }
  const bd = {
    ...(dxf.borderTop ? { t: toUniverBorder(dxf.borderTop) } : {}),
    ...(dxf.borderBottom ? { b: toUniverBorder(dxf.borderBottom) } : {}),
    ...(dxf.borderLeft ? { l: toUniverBorder(dxf.borderLeft) } : {}),
    ...(dxf.borderRight ? { r: toUniverBorder(dxf.borderRight) } : {}),
  }
  if (Object.keys(bd).length > 0) target.style = { ...target.style, bd }
}

function toCfValue(cfvo: { kind: string; value?: string | undefined }): IValueConfig {
  switch (cfvo.kind) {
    case 'min':
    case 'autoMin':
      return { type: CFValueType.min }
    case 'max':
    case 'autoMax':
      return { type: CFValueType.max }
    case 'percent':
      return { type: CFValueType.percent, value: Number(cfvo.value ?? 0) }
    case 'percentile':
      return { type: CFValueType.percentile, value: Number(cfvo.value ?? 0) }
    case 'formula':
      return { type: CFValueType.formula, value: toCfFormula(cfvo.value ?? '0') }
    default: {
      // Legacy writers put defined names / expressions into type="num" cfvos;
      // a NaN literal would collapse the scale to 0.
      const numeric = Number(cfvo.value ?? 0)
      return Number.isFinite(numeric)
        ? { type: CFValueType.num, value: numeric }
        : { type: CFValueType.formula, value: toCfFormula(cfvo.value ?? '0') }
    }
  }
}

/// Univer's CF formula service takes '='-prefixed formulas (same registry as
/// whenFormulaSatisfied); xlsx cfvo bodies come without the prefix.
function toCfFormula(body: string): string {
  return body.startsWith('=') ? body : `=${body}`
}

/// Formula CF costs one dependency tree per covered cell; above this, huge
/// (e.g. whole-column) rules keep the cheaper native condition.
export const CELLIS_FORMULA_CELL_LIMIT = 20_000
/// Above this, auto dataBar bounds fall back to Univer's raw data extremes
/// instead of paying a chunked full-range read on open.
export const CF_AUTO_BOUNDS_CELL_CAP = 512_000

/// Excel evaluates numeric cellIs rules on blank cells as 0; Univer's native
/// number conditions skip blanks (matching only notEqual/notBetween). True
/// when the two would paint blanks differently.
export function cellIsBlankDiverges(operator: string, first: number, second: number): boolean {
  let excelBlank: boolean
  switch (operator) {
    case 'greaterThan':
      excelBlank = 0 > first
      break
    case 'greaterThanOrEqual':
      excelBlank = 0 >= first
      break
    case 'lessThan':
      excelBlank = 0 < first
      break
    case 'lessThanOrEqual':
      excelBlank = 0 <= first
      break
    case 'equal':
      excelBlank = first === 0
      break
    case 'notEqual':
      excelBlank = first !== 0
      break
    case 'between':
      excelBlank = Math.min(first, second) <= 0 && 0 <= Math.max(first, second)
      break
    case 'notBetween':
      excelBlank = !(Math.min(first, second) <= 0 && 0 <= Math.max(first, second))
      break
    default:
      return false
  }
  const univerBlank = operator === 'notEqual' || operator === 'notBetween'
  return excelBlank !== univerBlank
}

function buildCellIsFormula(
  builder: ReturnType<UniverWorksheet['newConditionalFormattingRule']>,
  operator: string,
  anchor: string,
  first: number,
  second: number,
): CfHighlightBuilder | null {
  switch (operator) {
    case 'equal':
      return builder.whenFormulaSatisfied(`=${anchor}=${first}`)
    case 'notEqual':
      return builder.whenFormulaSatisfied(`=${anchor}<>${first}`)
    case 'greaterThan':
      return builder.whenFormulaSatisfied(`=${anchor}>${first}`)
    case 'greaterThanOrEqual':
      return builder.whenFormulaSatisfied(`=${anchor}>=${first}`)
    case 'lessThan':
      return builder.whenFormulaSatisfied(`=${anchor}<${first}`)
    case 'lessThanOrEqual':
      return builder.whenFormulaSatisfied(`=${anchor}<=${first}`)
    case 'between':
      return builder.whenFormulaSatisfied(
        `=AND(${anchor}>=${Math.min(first, second)},${anchor}<=${Math.max(first, second)})`,
      )
    case 'notBetween':
      return builder.whenFormulaSatisfied(
        `=NOT(AND(${anchor}>=${Math.min(first, second)},${anchor}<=${Math.max(first, second)}))`,
      )
    default:
      return null
  }
}

function buildHighlightCondition(
  builder: ReturnType<UniverWorksheet['newConditionalFormattingRule']>,
  rule: WorkbookRangeResult['conditionalRules'][number],
  anchor: string,
  coveredCells = 0,
): CfHighlightBuilder | null {
  const firstNumber = Number(rule.formulas[0])
  const secondNumber = Number(rule.formulas[1])
  switch (rule.ruleType) {
    case 'cellIs':
      if (!Number.isFinite(firstNumber)) {
        return buildCellIsNonNumeric(builder, rule, anchor, coveredCells)
      }
      if (
        (rule.operator === 'between' || rule.operator === 'notBetween') &&
        !Number.isFinite(secondNumber)
      ) {
        return null
      }
      if (
        rule.operator !== undefined &&
        coveredCells > 0 &&
        coveredCells <= CELLIS_FORMULA_CELL_LIMIT &&
        cellIsBlankDiverges(rule.operator, firstNumber, secondNumber)
      ) {
        return buildCellIsFormula(builder, rule.operator, anchor, firstNumber, secondNumber)
      }
      switch (rule.operator) {
        case 'greaterThan':
          return builder.whenNumberGreaterThan(firstNumber)
        case 'greaterThanOrEqual':
          return builder.whenNumberGreaterThanOrEqualTo(firstNumber)
        case 'lessThan':
          return builder.whenNumberLessThan(firstNumber)
        case 'lessThanOrEqual':
          return builder.whenNumberLessThanOrEqualTo(firstNumber)
        case 'equal':
          return builder.whenNumberEqualTo(firstNumber)
        case 'notEqual':
          return builder.whenNumberNotEqualTo(firstNumber)
        case 'between':
          return Number.isFinite(secondNumber)
            ? builder.whenNumberBetween(firstNumber, secondNumber)
            : null
        case 'notBetween':
          return Number.isFinite(secondNumber)
            ? builder.whenNumberNotBetween(firstNumber, secondNumber)
            : null
        default:
          return null
      }
    case 'containsText':
    case 'notContainsText':
    case 'beginsWith':
    case 'endsWith':
      return buildTextCondition(builder, rule, anchor, coveredCells)
    case 'containsBlanks':
      return builder.whenCellEmpty()
    case 'notContainsBlanks':
      return builder.whenCellNotEmpty()
    case 'containsErrors':
    case 'notContainsErrors':
      // Univer's calculate unit evaluates these operators but the facade
      // builder has no setter; build a placeholder text condition and
      // re-target the operator on the built rule (patchBuiltHighlightRule).
      return builder.whenTextContains('')
    case 'duplicateValues':
      return builder.setDuplicateValues()
    case 'uniqueValues':
      return builder.setUniqueValues()
    case 'top10':
      return rule.rank === undefined
        ? null
        : builder.setRank({
            isBottom: rule.bottom,
            isPercent: rule.percent,
            value: rule.rank,
          })
    case 'expression':
      return rule.formulas[0] ? builder.whenFormulaSatisfied(`=${rule.formulas[0]}`) : null
    default:
      return null
  }
}

/// Excel's text operators compare case-insensitively (containsText is
/// SEARCH-based), but Univer's evaluator is a raw indexOf — "SHORT" never
/// matches a containsText "Short" rule. Install the rule as a formula (the
/// file's own SEARCH/LEFT/RIGHT formula when present) and keep the cheaper
/// native condition for huge ranges, where the case gap is the lesser evil.
function buildTextCondition(
  builder: ReturnType<UniverWorksheet['newConditionalFormattingRule']>,
  rule: WorkbookRangeResult['conditionalRules'][number],
  anchor: string,
  coveredCells: number,
): CfHighlightBuilder | null {
  const text = rule.text
  if (!text) return null
  if (coveredCells > 0 && coveredCells <= CELLIS_FORMULA_CELL_LIMIT) {
    // Always synthesized: the file's own formula is anchored to its first
    // sqref cell, but Univer offsets from the top-left-sorted first range.
    const formula = synthesizeTextConditionFormula(rule.ruleType, text, anchor)
    if (formula) return builder.whenFormulaSatisfied(toCfFormula(formula))
  }
  switch (rule.ruleType) {
    case 'containsText':
      return builder.whenTextContains(text)
    case 'notContainsText':
      return builder.whenTextDoesNotContain(text)
    case 'beginsWith':
      return builder.whenTextStartsWith(text)
    case 'endsWith':
      return builder.whenTextEndsWith(text)
    default:
      return null
  }
}

function synthesizeTextConditionFormula(
  ruleType: string,
  text: string,
  anchor: string,
): string | null {
  const quoted = `"${text.replace(/"/g, '""')}"`
  switch (ruleType) {
    case 'containsText':
      return `NOT(ISERROR(SEARCH(${quoted},${anchor})))`
    case 'notContainsText':
      return `ISERROR(SEARCH(${quoted},${anchor}))`
    case 'beginsWith':
      return `LEFT(${anchor},LEN(${quoted}))=${quoted}`
    case 'endsWith':
      return `RIGHT(${anchor},LEN(${quoted}))=${quoted}`
    default:
      return null
  }
}

/// cellIs with a non-numeric operand: a quoted string or a cell reference /
/// expression, both installed as a formula rule so text compares the way
/// Excel does (case-insensitive: `"A"` highlights `a`). Univer's native text
/// rule is a strict `===`, so it is only the fallback for huge ranges where
/// a formula per cell is too expensive. `anchor` is the relative top-left of
/// the rule's first range, the cell Excel evaluates against.
function buildCellIsNonNumeric(
  builder: ReturnType<UniverWorksheet['newConditionalFormattingRule']>,
  rule: WorkbookRangeResult['conditionalRules'][number],
  anchor: string,
  coveredCells = 0,
): CfHighlightBuilder | null {
  const first = rule.formulas[0]
  const second = rule.formulas[1]
  if (!first) return null
  const quoted = /^"([\s\S]*)"$/.exec(first)
  switch (rule.operator) {
    case 'equal':
      if (!quoted) return builder.whenFormulaSatisfied(`=${anchor}=(${first})`)
      return coveredCells > 0 && coveredCells <= CELLIS_FORMULA_CELL_LIMIT
        ? builder.whenFormulaSatisfied(`=${anchor}=${escapeCfWildcards(first)}`)
        : builder.whenTextEqualTo(quoted[1]!.replace(/""/g, '"'))
    case 'notEqual':
      return builder.whenFormulaSatisfied(`=${anchor}<>${wrapOperand(first)}`)
    case 'greaterThan':
      return builder.whenFormulaSatisfied(`=${anchor}>${wrapOperand(first)}`)
    case 'greaterThanOrEqual':
      return builder.whenFormulaSatisfied(`=${anchor}>=${wrapOperand(first)}`)
    case 'lessThan':
      return builder.whenFormulaSatisfied(`=${anchor}<${wrapOperand(first)}`)
    case 'lessThanOrEqual':
      return builder.whenFormulaSatisfied(`=${anchor}<=${wrapOperand(first)}`)
    case 'between':
      return second
        ? builder.whenFormulaSatisfied(
            `=AND(${anchor}>=${wrapOperand(first)},${anchor}<=${wrapOperand(second)})`,
          )
        : null
    case 'notBetween':
      return second
        ? builder.whenFormulaSatisfied(
            `=NOT(AND(${anchor}>=${wrapOperand(first)},${anchor}<=${wrapOperand(second)}))`,
          )
        : null
    default:
      return null
  }
}

/// Univer's `=` treats `*` / `?` in a string operand as wildcards; Excel's
/// does not, so they are `~`-escaped before the literal enters the formula.
function escapeCfWildcards(quotedOperand: string): string {
  return quotedOperand.replace(/[*?]/g, '~$&')
}

/// Quoted strings must stay verbatim; anything else gets parenthesized so
/// composite expressions keep their precedence inside the comparison.
function wrapOperand(operand: string): string {
  return /^"[\s\S]*"$/.test(operand) ? operand : `(${operand})`
}

function applyDxfFormat(
  highlight: CfHighlightBuilder,
  dxfStyles: readonly WorkbookCellStyle[],
  dxfIndex: number | undefined,
): CfHighlightBuilder {
  const dxf = dxfIndex === undefined ? undefined : dxfStyles[dxfIndex]
  if (!dxf) return highlight
  let styled = highlight
  if (dxf.fillColor) styled = styled.setBackground(dxf.fillColor)
  if (dxf.fontColor) styled = styled.setFontColor(dxf.fontColor)
  if (dxf.bold) styled = styled.setBold(true)
  if (dxf.italic) styled = styled.setItalic(true)
  if (dxf.underline) styled = styled.setUnderline(true)
  if (dxf.strikethrough) styled = styled.setStrikethrough(true)
  return styled
}

export function installCustom(
  shrink: boolean,
  style: WorkbookCellStyle | undefined,
): { custom?: Record<string, boolean> } {
  const custom = {
    ...(shrink ? { [SHRINK_TO_FIT_KEY]: true } : {}),
    ...protectionCustom(style).custom,
  }
  return Object.keys(custom).length > 0 ? { custom } : {}
}

/// xf protection flags ride along in ICellData.custom (Univer has no model).
export function protectionCustom(style: WorkbookCellStyle | undefined): {
  custom?: Record<string, boolean>
} {
  if (!style || (style.locked === undefined && style.hidden !== true)) return {}
  return {
    custom: {
      ...(style.locked === undefined ? {} : { [CELL_LOCKED_KEY]: style.locked }),
      ...(style.hidden === true ? { [CELL_HIDDEN_KEY]: true } : {}),
    },
  }
}

export function toUniverStyle(style: WorkbookCellStyle): IStyleData {
  const diagonal = style.borderDiagonal ? toUniverBorder(style.borderDiagonal) : undefined
  const borders = {
    ...(style.borderTop ? { t: toUniverBorder(style.borderTop) } : {}),
    ...(style.borderBottom ? { b: toUniverBorder(style.borderBottom) } : {}),
    ...(style.borderLeft ? { l: toUniverBorder(style.borderLeft) } : {}),
    ...(style.borderRight ? { r: toUniverBorder(style.borderRight) } : {}),
    ...(diagonal && style.diagonalDown ? { tl_br: diagonal } : {}),
    ...(diagonal && style.diagonalUp ? { bl_tr: diagonal } : {}),
  }
  return {
    ...(style.fontFamily ? { ff: escapeCssLeadingDigit(style.fontFamily) } : {}),
    ...(style.fontSize ? { fs: style.fontSize } : {}),
    bl: style.bold ? BooleanNumber.TRUE : BooleanNumber.FALSE,
    it: style.italic ? BooleanNumber.TRUE : BooleanNumber.FALSE,
    ...(style.underline ? { ul: { s: BooleanNumber.TRUE } } : {}),
    ...(style.strikethrough ? { st: { s: BooleanNumber.TRUE } } : {}),
    ...(style.vertAlign ? { va: baselineOffsetOf(style.vertAlign) } : {}),
    // xf alignment is fully resolved in the file: an explicit non-wrap cell
    // must override a WRAP column/row style at compose time.
    // centerContinuous rides on the overflow machinery (center-continuous.ts),
    // so it must not wrap: Excel renders a wrapText+centerContinuous title on
    // one line across the run, not folded inside the anchor cell (prod ref).
    tb:
      style.wrapText && style.horizontalAlignment !== 'centerContinuous'
        ? WrapStrategy.WRAP
        : WrapStrategy.OVERFLOW,
    ...(style.fontColor ? { cl: { rgb: style.fontColor } } : {}),
    // Like tb above: an explicit xf without a fill must BLOCK a filled
    // column/row style at compose time — Univer merges styles by key, so a
    // missing bg lets a <col style=> fill bleed through explicitly-styled
    // cells (Excel treats each xf as complete, never a merge). See
    // NO_FILL_STYLE for why the sentinel rather than bg: null.
    ...(style.fillColor ? { bg: { rgb: style.fillColor } } : { bg: { ...NO_FILL_STYLE } }),
    ...(style.numberFormat ? { n: { pattern: style.numberFormat } } : {}),
    ...(Object.keys(borders).length > 0 ? { bd: borders } : {}),
    ...(mapHorizontalAlignment(style.horizontalAlignment) === undefined
      ? {}
      : { ht: mapHorizontalAlignment(style.horizontalAlignment) }),
    ...(mapVerticalAlignment(style.verticalAlignment) === undefined
      ? {}
      : { vt: mapVerticalAlignment(style.verticalAlignment) }),
    ...(style.indent ? { pd: { l: style.indent * INDENT_STEP_PX } } : {}),
    ...(style.textRotation
      ? { tr: ooxmlTextRotationToUniver(style.textRotation) ?? undefined }
      : {}),
  }
}

function toUniverBorder(edge: NonNullable<WorkbookCellStyle['borderTop']>): {
  s: BorderStyleTypes
  cl: { rgb: string }
} {
  return {
    s: mapBorderStyle(edge.style),
    cl: { rgb: edge.color ?? '#000000' },
  }
}

function mapHorizontalAlignment(value: string | undefined): HorizontalAlign | undefined {
  if (value === 'left') return HorizontalAlign.LEFT
  if (value === 'center') return HorizontalAlign.CENTER
  // The center-across-selection box is widened by center-continuous.ts.
  if (value === 'centerContinuous') return HorizontalAlign.CENTER
  if (value === 'right') return HorizontalAlign.RIGHT
  if (value === 'justify') return HorizontalAlign.JUSTIFIED
  if (value === 'distributed') return HorizontalAlign.DISTRIBUTED
  return undefined
}

function mapVerticalAlignment(value: string | undefined): VerticalAlign | undefined {
  if (value === 'top') return VerticalAlign.TOP
  if (value === 'center') return VerticalAlign.MIDDLE
  if (value === 'bottom') return VerticalAlign.BOTTOM
  // Excel centres a justify/distributed block; Univer has no equivalent and
  // would otherwise fall through to its bottom default.
  if (value === 'justify' || value === 'distributed') return VerticalAlign.MIDDLE
  return undefined
}

export function columnLetter(index: number): string {
  let label = ''
  for (let i = index; i >= 0; i = Math.floor(i / 26) - 1) {
    label = String.fromCharCode(65 + (i % 26)) + label
  }
  return label
}

export function baselineOffsetOf(vertAlign: 'superscript' | 'subscript'): BaselineOffset {
  return vertAlign === 'subscript' ? BaselineOffset.SUBSCRIPT : BaselineOffset.SUPERSCRIPT
}

/// Excel paints only the highest-precedence rule of a "paint-once" type
/// (colorScale / dataBar / iconSet) on a cell; stacked rules of the same
/// type do not blend. Univer keeps whichever it evaluates last, so a
/// lower-precedence duplicate could win. Drop a rule whose every range is
/// fully covered by a higher-precedence same-type rule — the common shape
/// (tdf105272 carries three identical stacked scales plus the real one).
const PAINT_ONCE_TYPES = new Set(['colorScale', 'dataBar', 'iconSet'])

export function dropShadowedPaintOnceRules(
  rules: WorkbookRangeResult['conditionalRules'],
): WorkbookRangeResult['conditionalRules'] {
  const covers = (
    outer: (typeof rules)[number]['ranges'][number],
    inner: (typeof rules)[number]['ranges'][number],
  ): boolean =>
    outer.startRow <= inner.startRow &&
    outer.endRow >= inner.endRow &&
    outer.startColumn <= inner.startColumn &&
    outer.endColumn >= inner.endColumn
  return rules.filter(
    (rule) =>
      !PAINT_ONCE_TYPES.has(rule.ruleType) ||
      !rules.some(
        (other) =>
          other !== rule &&
          other.ruleType === rule.ruleType &&
          other.priority < rule.priority &&
          rule.ranges.every((range) => other.ranges.some((cover) => covers(cover, range))),
      ),
  )
}
