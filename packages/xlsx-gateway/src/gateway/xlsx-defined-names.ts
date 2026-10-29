/// Declarative defined-names save: rewrites workbook.xml's `<definedNames>`
/// from the editor's model. Entries the editor never models — `_xlnm.*`
/// built-ins, hidden names, and names it failed to install (`preserveNames`)
/// — stay byte-verbatim.

import { rewriteNameReferences, type NameReferencePlan } from '../domain/defined-name-references'
import { withFutureFunctionMarkers } from './future-functions'

export class DefinedNameError extends Error {}

export interface DefinedNameEntry {
  readonly name: string
  readonly formula: string
  /// localSheetId (position in workbook sheet order) for sheet-scoped names.
  readonly sheetIndex?: number | undefined
}

export interface DefinedNamesState {
  readonly names: readonly DefinedNameEntry[]
  readonly preserveNames: readonly string[]
}

/// Excel name rules (simplified): starts with a letter, `_`, or `\`;
/// continues with letters, digits, `_`, `.`, or `\`; must not look like an
/// A1 or R1C1 cell reference. Letters are Unicode — Excel accepts CJK names,
/// and Create from Selection builds them from localized headers, so this
/// must accept everything definedNameFromLabel produces.
const NAME_PATTERN = /^[\p{L}_\\][\p{L}\p{N}_.\\]*$/u
const CELL_REF_PATTERN = /^(?:[A-Za-z]{1,3}[0-9]+|[Rr][0-9]*[Cc][0-9]*)$/

function definitionInfo(element: string) {
  const head = element.slice(0, element.indexOf('>') + 1)
  const name = unescapeXml(/\bname="([^"]*)"/.exec(head)?.[1] ?? '')
  const body = head.endsWith('/>')
    ? ''
    : element.slice(head.length, element.lastIndexOf('</definedName>'))
  const text = unescapeXml(body.replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, '$1'))
  return {
    head,
    name,
    fileOnly:
      name.toLowerCase().startsWith('_xlnm') ||
      /\bhidden="(?:1|true)"/.test(head) ||
      text.trim() === '',
  }
}
const definitionElements = /<definedName\b[^>]*(?:\/>|>[\s\S]*?<\/definedName>)/g

export function applyDefinedNamesState(workbookXml: string, state: DefinedNamesState): string {
  const preserved = new Set(state.preserveNames)
  const preservedKeys = new Set<string>()
  for (const match of workbookXml.matchAll(definitionElements)) {
    const { head, name, fileOnly } = definitionInfo(match[0])
    if (!fileOnly && !preserved.has(name)) continue
    const scope = /\blocalSheetId="(\d+)"/.exec(head)?.[1]
    preservedKeys.add(`${name.toLowerCase()}\u0000${scope === undefined ? -1 : Number(scope)}`)
  }
  const seen = new Set<string>()
  for (const entry of state.names) {
    validateName(entry.name)
    const key = `${entry.name.toLowerCase()}\u0000${entry.sheetIndex ?? -1}`
    if (preservedKeys.has(key)) {
      throw new DefinedNameError(
        `The name "${entry.name}" also exists in a form the editor cannot model — ` +
          'saving would duplicate it.',
      )
    }
    if (seen.has(key)) {
      throw new DefinedNameError(`The name "${entry.name}" is defined twice.`)
    }
    seen.add(key)
  }

  // Drop every modeled entry from the existing section, keeping built-ins,
  // hidden names, and preserve-listed ones in place.
  const xml = workbookXml.replace(definitionElements, (element) => {
    const { name, fileOnly } = definitionInfo(element)
    return fileOnly || preserved.has(name) ? element : ''
  })

  const additions = state.names
    .map(
      (entry) =>
        `<definedName name="${escapeXmlAttribute(entry.name)}"` +
        (entry.sheetIndex === undefined ? '' : ` localSheetId="${entry.sheetIndex}"`) +
        `>${escapeXmlText(withFutureFunctionMarkers(entry.formula.replace(/^=/, '')))}</definedName>`,
    )
    .join('')

  return appendDefinedNames(xml, additions)
}

/**
 * Append entries to the workbook's `<definedNames>` section, creating the
 * section where there is none.
 *
 * Shared with the redaction mirror (`applyRedactionNameMirror`), which writes
 * into the same section with a different idea of what an entry means. The
 * schema order is the fiddly part — `<definedNames>` follows `<sheets>`, but
 * only after `<functionGroups>` and `<externalReferences>` when those are
 * present — so it is done once here rather than twice.
 */
export function appendDefinedNames(workbookXml: string, additions: string): string {
  const section = /<definedNames\b[^>]*>([\s\S]*?)<\/definedNames>|<definedNames\b[^>]*\/>/.exec(
    workbookXml,
  )
  if (section) {
    const inner = (section[1] ?? '') + additions
    const replacement = inner === '' ? '' : `<definedNames>${inner}</definedNames>`
    return (
      workbookXml.slice(0, section.index) +
      replacement +
      workbookXml.slice(section.index + section[0].length)
    )
  }
  if (additions === '') return workbookXml
  // Schema order: definedNames follows sheets (and functionGroups/externalReferences).
  const anchor = /<\/sheets>|<sheets\b[^>]*\/>/.exec(workbookXml)
  if (!anchor) throw new DefinedNameError('workbook.xml has no sheets element.')
  const groups = /<functionGroups\b[^>]*>[\s\S]*?<\/functionGroups>|<functionGroups\b[^>]*\/>/.exec(
    workbookXml,
  )
  const externals =
    /<externalReferences\b[^>]*>[\s\S]*?<\/externalReferences>|<externalReferences\b[^>]*\/>/.exec(
      workbookXml,
    )
  let at = externals ? externals.index + externals[0].length : anchor.index + anchor[0].length
  // functionGroups only counts when externalReferences is absent: the later
  // element is the one the section must follow
  if (!externals && groups) at = groups.index + groups[0].length
  return `${workbookXml.slice(0, at)}<definedNames>${additions}</definedNames>${workbookXml.slice(at)}`
}

/** Shared by table names, which follow the same rules (minus the `_xlnm` reservation). */
export function isValidDefinedName(name: string): boolean {
  return (
    name.length > 0 &&
    name.length <= 255 &&
    NAME_PATTERN.test(name) &&
    !CELL_REF_PATTERN.test(name) &&
    name.toLowerCase() !== 'true' &&
    name.toLowerCase() !== 'false'
  )
}

function validateName(name: string): void {
  if (!isValidDefinedName(name)) {
    throw new DefinedNameError(`"${name}" is not a valid defined name.`)
  }
  if (name.toLowerCase().startsWith('_xlnm')) {
    throw new DefinedNameError('Names starting with "_xlnm" are reserved by Excel.')
  }
}

export function unescapeXml(input: string): string {
  return input
    .replaceAll('&quot;', '"')
    .replaceAll('&lt;', '<')
    .replaceAll('&gt;', '>')
    .replaceAll('&amp;', '&')
}

export function escapeXmlText(input: string): string {
  return input.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')
}

function escapeXmlAttribute(input: string): string {
  return input
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
}

/** Decode text and CDATA separately: entity-looking CDATA is literal content. */
export function decodeDefinedNameFormula(body: string): string {
  return body
    .split(/(<!\[CDATA\[[\s\S]*?\]\]>)/g)
    .map((part) => {
      if (part.startsWith('<![CDATA[')) return part.slice(9, -3)
      return part
        .replace(/<!--[\s\S]*?-->/g, '')
        .replace(
          /&(?:#x([0-9a-f]+)|#([0-9]+)|(quot|apos|lt|gt|amp));/gi,
          (match, hex: string | undefined, dec: string | undefined, named: string | undefined) => {
            if (named)
              return (
                ({ quot: '"', apos: "'", lt: '<', gt: '>', amp: '&' } as Record<string, string>)[
                  named
                ] ?? match
              )
            const code = hex === undefined ? Number(dec) : Number.parseInt(hex, 16)
            return code <= 0x10ffff && !(code >= 0xd800 && code <= 0xdfff)
              ? String.fromCodePoint(code)
              : match
          },
        )
    })
    .join('')
}

/** Replay source definition dependencies BEFORE replacing modeled definitions.
 * Replacement journal bodies already incorporate action-time edits. */
export function rewriteSourceDefinedNameReferences(xml: string, plan: NameReferencePlan): string {
  return xml.replace(definitionElements, (element) => {
    const openEnd = element.indexOf('>') + 1
    const head = element.slice(0, openEnd)
    if (head.endsWith('/>')) return element
    const scope = /\blocalSheetId="([^"]*)"/.exec(head)?.[1]
    const owner = scope === undefined ? '' : plan.sheets[Number(scope)]?.id
    if (owner === undefined || (scope !== undefined && !/^\d+$/.test(scope)))
      throw new DefinedNameError('Unknown defined-name reference scope.')
    const body = element.slice(openEnd, element.lastIndexOf('</definedName>'))
    const before = decodeDefinedNameFormula(body)
    const after = rewriteNameReferences(before, owner, plan)
    return after === before ? element : head + escapeXmlText(after) + '</definedName>'
  })
}
