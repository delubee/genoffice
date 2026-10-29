/// Declarative data-validation save: the renderer snapshots the full Univer
/// rule set of a dirty sheet and this module rewrites the worksheet's
/// `<dataValidations>` section from it (mirroring the CF/filter recipe).
/// Mappings are the exact inverse of the read-side install in App.tsx.

import { MAX_GRID_COLUMNS, MAX_GRID_ROWS } from '../shared/grid-bounds'
import type { DvRule } from '../domain/workbook-dsl'
import { parseRange } from '../domain/cell-address'
import { offsetFormulaRefs } from '../domain/formula-shift'

export class DvEditError extends Error {}

export interface DvCellArea {
  readonly startRow: number
  readonly endRow: number
  readonly startColumn: number
  readonly endColumn: number
}

/// One rule in the Univer data-validation model shape (validated structurally
/// here — unknown shapes fail the save rather than guess).
export interface DvWireRule {
  readonly ranges: readonly DvCellArea[]
  readonly rule: Record<string, unknown>
}

export function dvRuleToWireRule(rule: DvRule): Record<string, unknown> {
  const options = rule.options ?? {}
  const common = {
    allowBlank: options.allowBlank ?? true,
    showErrorMessage: options.showErrorMessage ?? true,
    showInputMessage: options.showInputMessage ?? false,
    ...(options.promptTitle === undefined ? {} : { promptTitle: options.promptTitle }),
    ...(options.prompt === undefined ? {} : { prompt: options.prompt }),
    ...(options.errorTitle === undefined ? {} : { errorTitle: options.errorTitle }),
    ...(options.error === undefined ? {} : { error: options.error }),
    ...(options.errorStyle === undefined
      ? {}
      : { errorStyle: { information: 0, stop: 1, warning: 2 }[options.errorStyle] }),
  }
  switch (rule.kind) {
    case 'any':
      return { ...common, type: 'any' }
    case 'list':
      return {
        ...common,
        type: 'list',
        formula1: JSON.stringify(rule.values),
        showDropDown: options.showDropdown ?? true,
      }
    case 'listRef':
      return {
        ...common,
        type: 'list',
        formula1: `=${rule.range}`,
        showDropDown: options.showDropdown ?? true,
      }
    case 'numberBetween':
      return {
        ...common,
        type: 'decimal',
        operator: 'between',
        formula1: String(rule.min),
        formula2: String(rule.max),
      }
    case 'dateBetween':
      return {
        ...common,
        type: 'date',
        operator: 'between',
        formula1: rule.start,
        formula2: rule.end,
      }
    case 'checkbox':
      return { ...common, type: 'checkbox' }
    case 'formula':
      return { ...common, type: 'custom', formula1: rule.formula }
    case 'criterion':
      return {
        ...common,
        type: rule.type,
        operator: rule.operator,
        formula1: rule.formula1,
        ...(rule.formula2 === undefined ? {} : { formula2: rule.formula2 }),
      }
  }
}

const DV_TYPES = new Set(['whole', 'decimal', 'list', 'date', 'time', 'textLength', 'custom'])
const DV_OPERATORS = new Set([
  'between',
  'notBetween',
  'equal',
  'notEqual',
  'greaterThan',
  'greaterThanOrEqual',
  'lessThan',
  'lessThanOrEqual',
])
/// Univer DataValidationErrorStyle: INFO=0, STOP=1 (OOXML default), WARNING=2.
const DV_ERROR_STYLE_NAMES: Record<number, string | undefined> = {
  0: 'information',
  1: undefined,
  2: 'warning',
}

export interface DvApplyOptions {
  /** keep the existing rules, replace those on the same ranges, drop `remove`, add the rest */
  readonly append?: boolean | undefined
  readonly remove?: readonly DvCellArea[] | undefined
}

export function applyDvRules(
  worksheetXml: string,
  rules: readonly DvWireRule[],
  options: DvApplyOptions = {},
): string {
  if (options.append) return appendDvRules(worksheetXml, rules, options.remove ?? [])
  if (/<x14:dataValidation\b/.test(worksheetXml)) {
    throw new DvEditError(
      'This sheet has extended (x14) data validation — editing its rules is not ' +
        'supported yet.',
    )
  }
  const xml = worksheetXml.replace(
    /<dataValidations\b[^>]*>[\s\S]*?<\/dataValidations>|<dataValidations\b[^>]*\/>/g,
    '',
  )
  if (rules.length === 0) return xml

  const body = rules.map(serializeRule).join('')
  const section = `<dataValidations count="${rules.length}">${body}</dataValidations>`
  const anchor =
    /<hyperlinks\b|<printOptions\b|<pageMargins\b|<pageSetup\b|<headerFooter\b|<rowBreaks\b|<colBreaks\b|<drawing\b|<legacyDrawing\b|<picture\b|<oleObjects\b|<tableParts\b|<extLst\b/.exec(
      xml,
    )
  if (anchor) {
    return xml.slice(0, anchor.index) + section + xml.slice(anchor.index)
  }
  const end = xml.lastIndexOf('</worksheet>')
  if (end === -1) throw new DvEditError('Worksheet has no closing element.')
  return xml.slice(0, end) + section + xml.slice(end)
}

const DV_SECTION_RE =
  /<dataValidations\b[^>]*>([\s\S]*?)<\/dataValidations>|<dataValidations\b[^>]*\/>/
const DV_ENTRY_RE = /<dataValidation\b[^>]*?\/>|<dataValidation\b[^>]*>[\s\S]*?<\/dataValidation>/g

function appendDvRules(
  worksheetXml: string,
  rules: readonly DvWireRule[],
  remove: readonly DvCellArea[],
): string {
  const section = DV_SECTION_RE.exec(worksheetXml)
  const cuts = [...rules.flatMap((rule) => rule.ranges), ...remove]
  // Split overlaps so new rules do not conflict with a surviving part of an old rule.
  const kept: string[] = []
  for (const entry of section?.[1] ? [...section[1].matchAll(DV_ENTRY_RE)].map((m) => m[0]) : []) {
    const sqref = /\bsqref="([^"]*)"/.exec(entry)?.[1] ?? ''
    const areas = sqref.split(/\s+/).filter(Boolean)
    if (areas.length === 0) throw new DvEditError('A data-validation rule has no sqref area.')
    // Areas outside the parser's grammar are document-controlled: keep them verbatim.
    const stationary: string[] = [],
      shifted: { original: DvCellArea; range: DvCellArea }[] = []
    let changed = false
    for (const ref of areas) {
      const original = tryParseSqrefArea(ref)
      if (!original) {
        stationary.push(ref)
        continue
      }
      const residuals = subtractDvAreas(original, cuts)
      if (residuals.length !== 1 || !sameArea(original, residuals[0]!)) changed = true
      for (const range of residuals)
        if (range.startRow === original.startRow && range.startColumn === original.startColumn)
          stationary.push(toRef(range))
        else shifted.push({ original, range })
    }
    if (!changed) {
      kept.push(entry)
      continue
    }
    // function replacer: surviving area text must not expand $& / $1 patterns
    if (stationary.length)
      kept.push(entry.replace(/\bsqref="[^"]*"/, () => `sqref="${stationary.join(' ')}"`))
    for (const { original, range } of shifted)
      kept.push(
        offsetValidationFormulas(
          entry.replace(/\bsqref="[^"]*"/, () => `sqref="${toRef(range)}"`),
          range.startRow - original.startRow,
          range.startColumn - original.startColumn,
        ),
      )
  }
  const entries = [...kept, ...rules.map(serializeRule)]
  const xml = section ? worksheetXml.replace(section[0], '') : worksheetXml
  if (entries.length === 0) return xml
  const body = `<dataValidations count="${entries.length}">${entries.join('')}</dataValidations>`
  if (section) return xml.slice(0, section.index) + body + xml.slice(section.index)
  return insertBeforeTail(xml, body)
}

function tryParseSqrefArea(ref: string): DvCellArea | null {
  try {
    const area = parseRange(normalizeRef(ref))
    toRef(area)
    return area
  } catch {
    return null
  }
}

function sameArea(left: DvCellArea, right: DvCellArea): boolean {
  return (
    left.startRow === right.startRow &&
    left.endRow === right.endRow &&
    left.startColumn === right.startColumn &&
    left.endColumn === right.endColumn
  )
}

export function subtractDvAreas(area: DvCellArea, cuts: readonly DvCellArea[]): DvCellArea[] {
  let remaining = [area]
  for (const cut of cuts) remaining = remaining.flatMap((part) => subtractArea(part, cut))
  return remaining
}

/** Remove overlap from pending rules, keeping relative formulas anchored to each residual. */
export function subtractDvRules(
  rules: readonly DvWireRule[],
  cuts: readonly DvCellArea[],
): DvWireRule[] {
  return rules.flatMap((entry) =>
    entry.ranges.flatMap((original) =>
      subtractDvAreas(original, cuts).map((range) => {
        const row = range.startRow - original.startRow,
          column = range.startColumn - original.startColumn
        const rule = { ...entry.rule }
        if (row || column)
          for (const key of ['formula1', 'formula2']) {
            const value = rule[key]
            if (typeof value === 'string' && !(rule.type === 'list' && !value.startsWith('=')))
              rule[key] = offsetFormulaRefs(value, row, column)
          }
        return { ranges: [range], rule }
      }),
    ),
  )
}

/** Native list options are JSON arrays; legacy imports may still use comma text. */
export function nativeDvListValues(text: string): string[] {
  try {
    const values: unknown = JSON.parse(text)
    if (Array.isArray(values) && values.every((value) => typeof value === 'string'))
      return values.filter(Boolean)
  } catch {
    /* Legacy comma lists are still supported. */
  }
  return text.split(',').filter(Boolean)
}

function subtractArea(area: DvCellArea, cut: DvCellArea): DvCellArea[] {
  const top = Math.max(area.startRow, cut.startRow),
    bottom = Math.min(area.endRow, cut.endRow),
    left = Math.max(area.startColumn, cut.startColumn),
    right = Math.min(area.endColumn, cut.endColumn)
  if (top > bottom || left > right) return [area]
  const pieces: DvCellArea[] = []
  if (area.startRow < top) pieces.push({ ...area, endRow: top - 1 })
  if (bottom < area.endRow) pieces.push({ ...area, startRow: bottom + 1 })
  if (area.startColumn < left)
    pieces.push({
      startRow: top,
      endRow: bottom,
      startColumn: area.startColumn,
      endColumn: left - 1,
    })
  if (right < area.endColumn)
    pieces.push({
      startRow: top,
      endRow: bottom,
      startColumn: right + 1,
      endColumn: area.endColumn,
    })
  return pieces
}

function offsetValidationFormulas(entry: string, rowDelta: number, columnDelta: number): string {
  return entry.replace(
    /(<formula([12])>)([\s\S]*?)(<\/formula\2>)/g,
    (_match, open: string, _index: string, encoded: string, close: string) =>
      `${open}${escapeXmlText(offsetFormulaRefs(unescapeFormulaXml(encoded), rowDelta, columnDelta))}${close}`,
  )
}

function unescapeFormulaXml(value: string): string {
  return value.replace(/&(#x[0-9a-f]+|#\d+|quot|apos|lt|gt|amp);/gi, (entity, code: string) => {
    if (code === 'quot') return '"'
    if (code === 'apos') return "'"
    if (code === 'lt') return '<'
    if (code === 'gt') return '>'
    if (code === 'amp') return '&'
    const numeric =
      code[1]?.toLowerCase() === 'x' ? Number.parseInt(code.slice(2), 16) : Number(code.slice(1))
    return Number.isInteger(numeric) && numeric >= 0 && numeric <= 0x10ffff
      ? String.fromCodePoint(numeric)
      : entity
  })
}

/** `A1:A1` and `$A$1` name the same cell as `A1`; `A:A` / `1:1` expand to the full grid. */
function normalizeRef(ref: string): string {
  const [a, b] = ref.replace(/\$/g, '').toUpperCase().split(':')
  if (b !== undefined) {
    if (/^[A-Z]+$/.test(a!) && /^[A-Z]+$/.test(b)) return `${a}1:${b}${MAX_GRID_ROWS}`
    if (/^\d+$/.test(a!) && /^\d+$/.test(b))
      return `A${a}:${columnToLetters(MAX_GRID_COLUMNS - 1)}${b}`
  }
  return b === undefined || b === a ? a! : `${a}:${b}`
}

function insertBeforeTail(xml: string, section: string): string {
  const anchor =
    /<hyperlinks\b|<printOptions\b|<pageMargins\b|<pageSetup\b|<headerFooter\b|<rowBreaks\b|<colBreaks\b|<drawing\b|<legacyDrawing\b|<picture\b|<oleObjects\b|<tableParts\b|<extLst\b/.exec(
      xml,
    )
  if (anchor) return xml.slice(0, anchor.index) + section + xml.slice(anchor.index)
  const end = xml.lastIndexOf('</worksheet>')
  if (end === -1) throw new DvEditError('Worksheet has no closing element.')
  return xml.slice(0, end) + section + xml.slice(end)
}

function serializeRule(wireRule: DvWireRule): string {
  if (wireRule.ranges.length === 0) {
    throw new DvEditError('A data-validation rule has no ranges.')
  }
  let rule = wireRule.rule
  let rawType = String(rule.type ?? '')
  if (rawType === 'listMultiple') {
    throw new DvEditError(
      'Multi-select list rules are Univer-only and cannot be saved to xlsx — ' +
        'delete the rule before saving.',
    )
  }
  if (rawType === 'checkbox') {
    // Checkbox is Univer-only; degrade to a two-value list so Excel keeps the
    // constraint. The default 1/0 pair round-trips back to a checkbox on load.
    const checked =
      rule.formula1 === undefined || rule.formula1 === '' ? '1' : String(rule.formula1)
    const unchecked =
      rule.formula2 === undefined || rule.formula2 === '' ? '0' : String(rule.formula2)
    rule = {
      ...rule,
      type: 'list',
      operator: undefined,
      formula1: `${checked},${unchecked}`,
      formula2: undefined,
    }
    rawType = 'list'
  }
  // 'any' is the read-side mapping of OOXML type="none" (no constraint, just
  // messages); it round-trips back to the default, attribute-less form.
  const type = rawType === 'any' || rawType === 'none' ? undefined : rawType
  if (type !== undefined && !DV_TYPES.has(type)) {
    throw new DvEditError(`Unsupported data-validation type "${rawType}".`)
  }

  const attrs: string[] = []
  if (type !== undefined) attrs.push(`type="${type}"`)
  const operator = rule.operator === undefined ? undefined : String(rule.operator)
  if (operator !== undefined && operator !== '') {
    if (!DV_OPERATORS.has(operator)) {
      throw new DvEditError(`Unsupported data-validation operator "${operator}".`)
    }
    // "between" is the OOXML default; only operator-carrying types keep it.
    if (operator !== 'between' && type !== undefined && type !== 'list' && type !== 'custom') {
      attrs.push(`operator="${operator}"`)
    }
  }
  if (rule.allowBlank === true) attrs.push('allowBlank="1"')
  // OOXML's showDropDown="1" SUPPRESSES the in-cell dropdown (inverted name);
  // Univer's showDropDown means what it says.
  if (type === 'list' && rule.showDropDown === false) attrs.push('showDropDown="1"')
  if (rule.showInputMessage === true) attrs.push('showInputMessage="1"')
  if (rule.showErrorMessage === true) attrs.push('showErrorMessage="1"')
  const errorStyle = errorStyleName(rule.errorStyle)
  if (errorStyle !== undefined) attrs.push(`errorStyle="${errorStyle}"`)
  for (const [key, attribute] of [
    ['errorTitle', 'errorTitle'],
    ['error', 'error'],
    ['promptTitle', 'promptTitle'],
    ['prompt', 'prompt'],
  ] as const) {
    const value = rule[key]
    if (typeof value === 'string' && value.length > 0) {
      attrs.push(`${attribute}="${escapeXmlAttribute(value)}"`)
    }
  }
  attrs.push(`sqref="${wireRule.ranges.map(toRef).join(' ')}"`)

  const formulas = serializeFormulas(type, rule)
  return formulas === ''
    ? `<dataValidation ${attrs.join(' ')}/>`
    : `<dataValidation ${attrs.join(' ')}>${formulas}</dataValidation>`
}

function errorStyleName(value: unknown): string | undefined {
  if (value === undefined || value === null) return undefined
  const style = Number(value)
  if (!(style in DV_ERROR_STYLE_NAMES)) {
    throw new DvEditError(`Unsupported data-validation error style "${String(value)}".`)
  }
  return DV_ERROR_STYLE_NAMES[style]
}

function serializeFormulas(type: string | undefined, rule: Record<string, unknown>): string {
  const formula1 = formulaText(type, rule.formula1)
  const formula2 = formulaText(type, rule.formula2)
  return (
    (formula1 === undefined ? '' : `<formula1>${escapeXmlText(formula1)}</formula1>`) +
    (formula2 === undefined ? '' : `<formula2>${escapeXmlText(formula2)}</formula2>`)
  )
}

/// Inverse of the install-side formula transforms: list literals regain their
/// quotes, `=`-prefixed references/formulas lose the prefix, and panel-edited
/// date/time strings become Excel serial numbers.
function formulaText(type: string | undefined, raw: unknown): string | undefined {
  if (raw === undefined || raw === null) return undefined
  const text = String(raw)
  if (text === '') return undefined
  if (type === 'list') {
    if (text.startsWith('=')) return text.slice(1)
    const values = nativeDvListValues(text)
    if (values.some((value) => value.includes(',')))
      throw new DvEditError(
        'An inline Excel validation option cannot contain a comma. Use a source range instead.',
      )
    return `"${values.join(',').replaceAll('"', '""')}"`
  }
  if (type === 'custom') {
    return text.startsWith('=') ? text.slice(1) : text
  }
  if (type === 'date') {
    const serial = dateToSerial(text)
    if (serial !== undefined) return String(serial)
  }
  if (type === 'time') {
    const fraction = timeToFraction(text)
    if (fraction !== undefined) return String(fraction)
  }
  return text.startsWith('=') ? text.slice(1) : text
}

/// 'YYYY-MM-DD[ HH:mm[:ss]]' (or slashes) → Excel serial (days since
/// 1899-12-30). Plain numbers and references pass through untouched.
/// Impossible calendar dates or clock times return undefined so the caller
/// keeps the original text instead of writing a silently wrong serial; a
/// pre-1900 year throws DvEditError, because it has no serial to write.
/// Serials 1-59 sit one day before the linear rule (Excel's phantom 1900-02-29);
/// formatSerial in @genoffice/file-parse reverses the same shift.
function dateToSerial(text: string): number | undefined {
  const match =
    /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[T ](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/.exec(
      text.trim(),
    )
  if (!match) return undefined
  const [, year, month, day, hour, minute, second] = match
  const yearNum = Number(year)
  const monthNum = Number(month)
  const dayNum = Number(day)
  // Excel's 1900 date system has no serial before 1900-01-01, so such a date
  // cannot be written at all. Date.UTC also maps a 0..99 year onto 19xx
  // (0099 -> 1999), which would otherwise store a plausible, silently wrong
  // serial; reject the year outright instead.
  if (yearNum < 1900) {
    throw new DvEditError(
      `Data-validation date "${text.trim()}" is before 1900, which Excel cannot store.`,
    )
  }
  if (monthNum < 1 || monthNum > 12) return undefined
  if (dayNum < 1 || dayNum > daysInMonth(yearNum, monthNum)) return undefined
  let seconds = 0
  if (hour !== undefined) {
    const hourNum = Number(hour)
    const minuteNum = Number(minute)
    const secondNum = Number(second ?? 0)
    if (hourNum < 0 || hourNum > 23) return undefined
    if (minuteNum < 0 || minuteNum > 59) return undefined
    if (secondNum < 0 || secondNum > 59) return undefined
    seconds = hourNum * 3600 + minuteNum * 60 + secondNum
  }
  const days = (Date.UTC(yearNum, monthNum - 1, dayNum) - Date.UTC(1899, 11, 30)) / 86_400_000
  const serial = days < 61 ? days - 1 : days
  return seconds === 0 ? serial : serial + seconds / 86_400
}

/// Days in a 1-based month, with the Gregorian leap-year rule for February.
function daysInMonth(year: number, month: number): number {
  if (month === 2) return isLeapYear(year) ? 29 : 28
  if (month === 4 || month === 6 || month === 9 || month === 11) return 30
  return 31
}

/// Gregorian leap-year rule: divisible by 4, except centuries not by 400.
function isLeapYear(year: number): boolean {
  return (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0
}

function timeToFraction(text: string): number | undefined {
  const match = /^(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/.exec(text.trim())
  if (!match) return undefined
  const [, hour, minute, second] = match
  const hourNum = Number(hour)
  const minuteNum = Number(minute)
  const secondNum = Number(second ?? 0)
  if (hourNum < 0 || hourNum > 23) return undefined
  if (minuteNum < 0 || minuteNum > 59) return undefined
  if (secondNum < 0 || secondNum > 59) return undefined
  return (hourNum * 3600 + minuteNum * 60 + secondNum) / 86_400
}

function toRef(range: DvCellArea): string {
  if (
    range.startRow < 0 ||
    range.startColumn < 0 ||
    range.endRow < range.startRow ||
    range.endColumn < range.startColumn ||
    range.endRow >= MAX_GRID_ROWS ||
    range.endColumn >= MAX_GRID_COLUMNS
  ) {
    throw new DvEditError('A data-validation range lies outside the worksheet grid.')
  }
  const start = `${columnToLetters(range.startColumn)}${range.startRow + 1}`
  return range.startRow === range.endRow && range.startColumn === range.endColumn
    ? start
    : `${start}:${columnToLetters(range.endColumn)}${range.endRow + 1}`
}

function columnToLetters(column: number): string {
  let letters = ''
  let remaining = column + 1
  while (remaining > 0) {
    remaining -= 1
    letters = String.fromCharCode(65 + (remaining % 26)) + letters
    remaining = Math.floor(remaining / 26)
  }
  return letters
}

function escapeXmlText(input: string): string {
  return input.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')
}

function escapeXmlAttribute(input: string): string {
  return input
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll('\n', '&#10;')
}
