import type { EditorJSON } from '../conversion'
import { docxSchema } from './schema'

export const DOCX_LIMITS = {
  file: 16 * 1024 ** 2,
  root: 32 * 1024 ** 2,
  steps: 1000,
  stepBytes: 1024 ** 2,
  history: 256 * 1024 ** 2,
} as const
/** Blocks with a host identity: the session anchors each to its source XML. */
const tracked = new Set(['docParagraph', 'docHeading', 'docListItem', 'docTable', 'docProtected'])
const textBlocks = new Set(['docParagraph', 'docHeading', 'docListItem'])
/**
 * Attributes the serializer writes verbatim as XML, or that anchor content to
 * source parts (comments, revisions, fields, relationships). The view may keep
 * a value it was given but never author one: a new or split owner drops them.
 */
const rawAttrs = new Set([
  'docxIndex',
  'rawPPr',
  'rawRPr',
  'rawTrPr',
  'rawTcPr',
  'xml',
  'omml',
  'model',
  'sdtShell',
  'sdtPr',
  'beginXml',
  'instr',
  'fieldId',
  'fieldPart',
  'rId',
  'ids',
  'originalStructure',
  'originalFormatting',
  'blockRevision',
  'rowRevision',
  'cellRevision',
  'pPrChange',
  'paraMarkIns',
  'paraMarkDel',
  'tblFloatSource',
  'bookmarks',
  'hiddenBookmarks',
  'commentStarts',
  'commentEnds',
  'genXml',
  'genImage',
  'genChart',
])
/** Marks whose meaning lives in another package part (only kept, never created). */
const sourceMarks = new Set(['refField', 'instrField', 'ctrlCheckbox', 'redaction'])
const mutableAttrs = new Set([
  'align',
  'indentLeft',
  'indentRight',
  'firstLineIndent',
  'hangingIndent',
  'spaceBefore',
  'spaceAfter',
  'spaceBeforeAuto',
  'spaceAfterAuto',
  'lineFactor',
  'level',
  'numId',
  'kind',
  'ilvl',
  'styleId',
  'colWidthsPct',
  'widthPx',
  'widthPct',
  'indentFirstLine',
  'indentFirstLineChars',
  'lineSpacing',
  'lineRule',
  'lineRawTwips',
  'widowControl',
  'keepNext',
  'keepLines',
  'suppressLineNumbers',
  'contextualSpacing',
  'pageBreakBefore',
  'outlineOnly',
])
const nestedChecked = new Set(['colwidth', 'heightTwips', 'heightRule'])
const textStyleAttrs = new Set([
  'color',
  'sizeHalfPoints',
  'font',
  'eastAsiaFont',
  'fontAscii',
  'csFont',
  'eaSlotEmpty',
  'vertAlign',
  'dstrike',
  'caps',
  'vanish',
  'vanishOwn',
  'charScalePct',
  'charScaleEm',
  'charSpacingTwips',
  'positionHalfPoints',
  'kernHalfPoints',
  'kern',
])
const fontNumberBounds: Record<string, readonly [number, number]> = {
  charScalePct: [1, 600],
  charScaleEm: [-100, 100],
  charSpacingTwips: [-31680, 31680],
  positionHalfPoints: [-3168, 3168],
  kernHalfPoints: [0, 1638],
}
// Strict: a non-finite number would otherwise encode as null and pass as unchanged.
const encode = (value: unknown) =>
  JSON.stringify(value, (_key, item: unknown) => {
    if (typeof item === 'number' && !Number.isFinite(item)) throw new Error('Invalid number')
    return item
  })
export function jsonBytes(value: unknown): number {
  return new TextEncoder().encode(encode(value)).length
}

function walk(node: EditorJSON, visit: (node: EditorJSON) => void) {
  visit(node)
  for (const child of node.content ?? []) walk(child, visit)
}

/** Private runs never reach the view; host-protected objects stay placeholders. */
function needsProtection(node: EditorJSON): boolean {
  if (node.type === 'docProtected') return true
  let redacted = false
  walk(node, (child) => {
    if (child.marks?.some((mark) => mark.type === 'redaction')) redacted = true
  })
  return redacted
}

export function prepareDocument(input: EditorJSON, makeId: () => string): EditorJSON {
  const doc = structuredClone(input)
  doc.content = (doc.content ?? []).map((node) =>
    needsProtection(node) && node.type !== 'docProtected'
      ? {
          type: 'docProtected',
          attrs: {
            docxIndex: node.attrs?.docxIndex ?? null,
            blockType: 'preserved',
            label: 'Preserved content',
            previewText: '',
          },
        }
      : node,
  )
  for (const node of doc.content)
    if (
      node.type === 'docProtected' &&
      node.attrs?.docxIndex == null &&
      !node.attrs?.genXml &&
      !node.attrs?.genImage &&
      !node.attrs?.genChart
    )
      throw new Error('Unsupported content cannot be safely anchored to the original DOCX')
  walk(doc, (node) => {
    if (tracked.has(node.type))
      node.attrs = { ...node.attrs, officeId: makeId(), officeLocked: node.type === 'docProtected' }
  })
  return docxSchema.nodeFromJSON(doc).toJSON() as EditorJSON
}

function assertAttrValue(key: string, value: unknown) {
  if (value == null) return
  if (typeof value === 'number' && (!Number.isFinite(value) || Math.abs(value) > 100000))
    throw new Error(`Invalid ${key}`)
  if (key === 'color' && (typeof value !== 'string' || !/^[\da-f]{6}$/i.test(value)))
    throw new Error('Invalid color')
  if (key === 'sizeHalfPoints' && (typeof value !== 'number' || value < 2 || value > 800))
    throw new Error('Invalid font size')
  if (
    ['eaSlotEmpty', 'dstrike', 'vanish', 'vanishOwn', 'kern'].includes(key) &&
    typeof value !== 'boolean'
  )
    throw new Error(`Invalid ${key}`)
  if (key === 'vertAlign' && !['superscript', 'subscript'].includes(String(value)))
    throw new Error('Invalid baseline')
  if (key === 'caps' && !['all', 'small', 'none'].includes(String(value)))
    throw new Error('Invalid capitalization')
  if (
    [
      'widowControl',
      'keepNext',
      'keepLines',
      'suppressLineNumbers',
      'contextualSpacing',
      'pageBreakBefore',
      'outlineOnly',
      'spaceBeforeAuto',
      'spaceAfterAuto',
    ].includes(key) &&
    typeof value !== 'boolean'
  )
    throw new Error(`Invalid ${key}`)
  if (key === 'lineRule' && !['auto', 'exact', 'atLeast'].includes(String(value)))
    throw new Error('Invalid line spacing rule')
  if (key === 'lineSpacing' && (typeof value !== 'number' || value <= 0 || value > 1000))
    throw new Error('Invalid line spacing')
  if (
    [
      'indentLeft',
      'indentRight',
      'indentFirstLine',
      'indentFirstLineChars',
      'spaceBefore',
      'spaceAfter',
      'lineRawTwips',
    ].includes(key) &&
    (typeof value !== 'number' || !Number.isInteger(value))
  )
    throw new Error(`Invalid ${key}`)
  if (key === 'lineRawTwips' && Number(value) <= 0) throw new Error('Invalid line spacing height')
  const bounds = fontNumberBounds[key]
  if (
    bounds &&
    (typeof value !== 'number' ||
      !Number.isFinite(value) ||
      value < bounds[0] ||
      value > bounds[1] ||
      (key !== 'charScaleEm' && !Number.isInteger(value)))
  )
    throw new Error(`Invalid ${key}`)
  if (
    ['font', 'eastAsiaFont', 'fontAscii', 'csFont'].includes(key) &&
    // eslint-disable-next-line no-control-regex -- Font names cannot contain C0 control characters.
    (typeof value !== 'string' || value.length > 200 || /[<>\x00-\x1f]/.test(value))
  )
    throw new Error('Invalid font')
  if (key === 'align' && !['left', 'right', 'center', 'justify'].includes(String(value)))
    throw new Error('Invalid alignment')
  if (
    ['level', 'ilvl', 'lineFactor', 'heightTwips', 'widthPx', 'widthPct'].includes(key) &&
    typeof value !== 'number'
  )
    throw new Error(`Invalid ${key}`)
  if (key === 'level' && (Number(value) < 1 || Number(value) > 9 || !Number.isInteger(value)))
    throw new Error('Invalid heading level')
  if (key === 'ilvl' && (Number(value) < 0 || Number(value) > 8 || !Number.isInteger(value)))
    throw new Error('Invalid list level')
  if (key === 'kind' && !['bullet', 'ordered'].includes(String(value)))
    throw new Error('Invalid list kind')
  if (key === 'heightRule' && !['atLeast', 'exact'].includes(String(value)))
    throw new Error('Invalid row height rule')
  if (
    ['colwidth', 'colWidthsPct'].includes(key) &&
    (!Array.isArray(value) ||
      value.length > 200 ||
      value.some(
        (width) =>
          typeof width !== 'number' || !Number.isFinite(width) || width <= 0 || width > 10000,
      ))
  )
    throw new Error('Invalid column widths')
}

/** Authored values are plain JSON of bounded shape; links and pictures are checked by kind. */
function assertShape(key: string, value: unknown, depth = 0): void {
  if (value == null || typeof value === 'boolean') return
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) throw new Error(`Invalid ${key}`)
    return
  }
  if (typeof value === 'string') {
    // eslint-disable-next-line no-control-regex -- C0 controls other than tab/newline never serialize.
    if (/[\x00-\x08\x0b\x0c\x0e-\x1f]/.test(value)) throw new Error(`Invalid ${key}`)
    if ((key === 'dataUrl' || key === 'imageDataUrl') && depth === 0)
      if (!/^data:image\/(png|jpeg|gif|bmp|webp|tiff);base64,[A-Za-z0-9+/=]*$/.test(value))
        throw new Error(`Invalid ${key}`)
    if (key === 'href' && /^\s*(javascript|vbscript|data|file):/i.test(value))
      throw new Error('Invalid link')
    return
  }
  if (depth >= 6) throw new Error(`Invalid ${key}`)
  if (Array.isArray(value)) {
    if (value.length > 4096) throw new Error(`Invalid ${key}`)
    for (const item of value) assertShape(key, item, depth + 1)
    return
  }
  if (typeof value === 'object') {
    const entries = Object.entries(value)
    if (entries.length > 256) throw new Error(`Invalid ${key}`)
    for (const [, item] of entries) assertShape(key, item, depth + 1)
    return
  }
  throw new Error(`Invalid ${key}`)
}

/** Insert → Pictures: a new picture block whose bytes the save writes as a fresh media part. */
function newPicture(attrs: Record<string, unknown>): boolean {
  const image = attrs.genImage as
    { base64?: unknown; mime?: unknown; widthPx?: unknown; heightPx?: unknown } | null | undefined
  return (
    attrs.blockType === 'image' &&
    attrs.docxIndex == null &&
    !!image &&
    typeof image.base64 === 'string' &&
    /^[A-Za-z0-9+/=]+$/.test(image.base64) &&
    ['image/png', 'image/jpeg', 'image/gif'].includes(String(image.mime)) &&
    typeof image.widthPx === 'number' &&
    typeof image.heightPx === 'number' &&
    image.widthPx > 0 &&
    image.heightPx > 0 &&
    image.widthPx <= 20000 &&
    image.heightPx <= 20000 &&
    attrs.imageDataUrl === `data:${String(image.mime)};base64,${image.base64}`
  )
}

/**
 * Source metadata and raw XML are immutable; every other schema value may be
 * authored within its shape bounds, as the desktop editor writes it.
 */
export function validateAndNormalize(
  before: EditorJSON,
  candidate: EditorJSON,
  makeId: () => string,
): EditorJSON {
  if (jsonBytes(candidate) > DOCX_LIMITS.root) throw new Error('Document exceeds size limit')
  docxSchema.nodeFromJSON(candidate).check()
  const originals = new Map<string, EditorJSON>()
  walk(before, (node) => {
    if (typeof node.attrs?.officeId === 'string') originals.set(node.attrs.officeId, node)
  })
  const seen = new Set<string>()
  // Raw values present anywhere in a source block, per node or mark type.
  const sourceValues = new WeakMap<EditorJSON, Map<string, Set<string>>>()
  const hasSourceValue = (source: EditorJSON, type: string, key: string, value: unknown) => {
    let values = sourceValues.get(source)
    if (!values) {
      values = new Map()
      const add = (owner: string, attrs: Record<string, unknown> | undefined) => {
        for (const [attr, original] of Object.entries(attrs ?? {})) {
          const name = `${owner}:${attr}`
          let known = values!.get(name)
          if (!known) values!.set(name, (known = new Set()))
          known.add(encode(original))
        }
      }
      walk(source, (node) => {
        add(node.type, node.attrs)
        for (const mark of node.marks ?? []) add(mark.type, mark.attrs)
      })
      sourceValues.set(source, values)
    }
    return values.get(`${type}:${key}`)?.has(encode(value)) ?? false
  }
  const defaultsOf = (type: string) => docxSchema.nodes[type]!.create().attrs
  const result = structuredClone(candidate)
  const verify = (node: EditorJSON, originalBlock?: EditorJSON) => {
    let original = originalBlock
    if (tracked.has(node.type)) {
      // A mid-paragraph split copies every attribute, the identity included
      // (keepOnSplit only filters end splits): the second owner is the new block.
      const repeated = typeof node.attrs?.officeId === 'string' && seen.has(node.attrs.officeId)
      if (repeated && originals.get(node.attrs!.officeId as string)?.attrs?.officeLocked)
        throw new Error('Invalid or duplicate block identity')
      const id = repeated ? null : node.attrs?.officeId
      if (id != null && (typeof id !== 'string' || !originals.has(id)))
        throw new Error('Invalid or duplicate block identity')
      if (typeof id === 'string') {
        seen.add(id)
        original = originals.get(id)!
        if (original.attrs?.officeLocked) {
          if (encode(node) !== encode(original))
            throw new Error('Protected content cannot be changed')
          return
        }
        const defaults = defaultsOf(node.type)
        for (const [key, value] of Object.entries(node.attrs ?? {})) {
          if (key === 'officeId' || key === 'officeLocked') continue
          const unchanged = encode(value) === encode(original.attrs?.[key])
          if (rawAttrs.has(key)) {
            if (
              !unchanged &&
              !(original.attrs?.[key] === undefined && encode(value) === encode(defaults[key]))
            )
              throw new Error(`Immutable attribute: ${key}`)
            continue
          }
          if (unchanged) continue
          assertShape(key, value)
          if (mutableAttrs.has(key) && textBlocks.has(node.type)) assertAttrValue(key, value)
        }
        if (
          node.type !== original.type &&
          !(textBlocks.has(node.type) && textBlocks.has(original.type))
        )
          throw new Error('Unsupported block conversion')
      } else {
        const picture = node.type === 'docProtected' && newPicture(node.attrs ?? {})
        if (node.type === 'docProtected' && !picture)
          throw new Error('Cannot create protected objects')
        const inheritedIndex = node.attrs?.docxIndex
        const inherited = [...originals.values()].find(
          (source) =>
            !source.attrs?.officeLocked &&
            inheritedIndex != null &&
            source.attrs?.docxIndex === inheritedIndex,
        )
        // A new owner never takes over source XML: anchors and raw values reset.
        const defaults = defaultsOf(node.type)
        node.attrs = { ...node.attrs, officeId: makeId(), officeLocked: false }
        for (const key of rawAttrs)
          if (key in node.attrs && !(picture && key === 'genImage'))
            node.attrs[key] = defaults[key] ?? null
        if (!picture && needsProtection(node)) throw new Error('Unsupported new content')
        for (const [key, value] of Object.entries(node.attrs)) {
          if (rawAttrs.has(key) || key === 'officeId' || key === 'officeLocked') continue
          assertShape(key, value)
          if (mutableAttrs.has(key) && textBlocks.has(node.type)) assertAttrValue(key, value)
        }
        original = inherited
      }
    } else if (node.type !== 'doc' && node.type !== 'text') {
      const defaults = defaultsOf(node.type)
      for (const [key, value] of Object.entries(node.attrs ?? {})) {
        if (rawAttrs.has(key)) {
          if (
            encode(value) !== encode(defaults[key]) &&
            !(original && hasSourceValue(original, node.type, key, value))
          )
            throw new Error(`Untrusted nested attribute: ${key}`)
          continue
        }
        assertShape(key, value)
        if (nestedChecked.has(key)) assertAttrValue(key, value)
      }
    }
    for (const mark of node.marks ?? []) {
      if (!docxSchema.marks[mark.type]) throw new Error('Unsupported mark')
      const fromSource = (key: string, value: unknown) =>
        !!original && value != null && hasSourceValue(original, mark.type, key, value)
      if (
        sourceMarks.has(mark.type) &&
        !Object.keys(mark.attrs ?? {}).every(
          (key) => fromSource(key, mark.attrs![key]) || mark.attrs![key] == null,
        )
      )
        throw new Error('Unsupported mark')
      if (sourceMarks.has(mark.type) && !original) throw new Error('Unsupported mark')
      for (const [key, value] of Object.entries(mark.attrs ?? {})) {
        if (value == null || fromSource(key, value)) continue
        // Comment anchors name comment ids; the session admits only ids of its comment list.
        if (mark.type === 'comment' && key === 'ids') {
          if (typeof value !== 'string' || !/^[0-9]{1,9}( [0-9]{1,9})*$/.test(value))
            throw new Error('Invalid comment anchor')
          continue
        }
        if (rawAttrs.has(key)) throw new Error(`Untrusted mark attribute: ${key}`)
        assertShape(key, value)
        if (textStyleAttrs.has(key) && mark.type === 'docTextStyle') assertAttrValue(key, value)
      }
    }
    for (const child of node.content ?? []) verify(child, original)
  }
  verify(result)
  for (const [id, node] of originals)
    if (node.attrs?.officeLocked && !seen.has(id))
      throw new Error('Protected content cannot be removed')
  return result
}
