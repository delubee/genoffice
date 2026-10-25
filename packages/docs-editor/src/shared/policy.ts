import type { EditorJSON } from '../conversion'
import { docxSchema } from './schema'

export const DOCX_LIMITS = {
  file: 16 * 1024 ** 2,
  root: 32 * 1024 ** 2,
  steps: 1000,
  stepBytes: 1024 ** 2,
  history: 256 * 1024 ** 2,
} as const
const editable = new Set(['docParagraph', 'docHeading', 'docListItem', 'docTable'])
const simpleMarks = new Set(['bold', 'italic', 'underline', 'strike', 'docTextStyle'])
const simpleInline = new Set(['text', 'hardBreak'])
const mutableAttrs = new Set([
  'align',
  'indentLeft',
  'indentRight',
  'firstLineIndent',
  'hangingIndent',
  'spaceBefore',
  'spaceAfter',
  'lineFactor',
  'level',
  'numId',
  'kind',
  'ilvl',
  'styleId',
  'colWidthsPct',
  'widthPx',
  'widthPct',
])
const nestedMutable = new Set(['colwidth', 'heightTwips', 'heightRule'])
const sourceOnlyAttrs = (attrs: Record<string, unknown>) =>
  Object.fromEntries(Object.entries(attrs).filter(([key]) => !nestedMutable.has(key)))
const sourceAttrs = new Set([
  'docxIndex',
  'rawPPr',
  'sdtShell',
  'bookmarks',
  'hiddenBookmarks',
  'commentStarts',
  'commentEnds',
  'pPrChange',
  'paraMarkIns',
  'paraMarkDel',
  'blockRevision',
  'originalStructure',
  'originalFormatting',
])
const textStyleAttrs = new Set([
  'color',
  'sizeHalfPoints',
  'font',
  'eastAsiaFont',
  'fontAscii',
  'csFont',
])
const encode = (value: unknown) => JSON.stringify(value)
export function jsonBytes(value: unknown): number {
  return new TextEncoder().encode(encode(value)).length
}

function walk(node: EditorJSON, visit: (node: EditorJSON) => void) {
  visit(node)
  for (const child of node.content ?? []) walk(child, visit)
}

/** Conservatively preserve blocks requiring features outside the first editor slice. */
function needsProtection(node: EditorJSON): boolean {
  if (node.type === 'docProtected') return true
  if (!editable.has(node.type)) return true
  let complex = Boolean(node.attrs?.sdtShell || node.attrs?.blockRevision)
  walk(node, (child) => {
    if (
      child.attrs?.sdtShell ||
      child.attrs?.blockRevision ||
      child.attrs?.rowRevision ||
      child.attrs?.paraMarkIns ||
      child.attrs?.paraMarkDel
    )
      complex = true
    if (
      child !== node &&
      !simpleInline.has(child.type) &&
      !['docTableRow', 'docTableCell', 'docTableHeader', ...editable].includes(child.type)
    )
      complex = true
    if (child.marks?.some((mark) => !simpleMarks.has(mark.type))) complex = true
    if (child.attrs?.colspan && child.attrs.colspan !== 1) complex = true
    if (child.attrs?.rowspan && child.attrs.rowspan !== 1) complex = true
  })
  return complex
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
    if (editable.has(node.type) || node.type === 'docProtected') {
      node.attrs = { ...node.attrs, officeId: makeId(), officeLocked: node.type === 'docProtected' }
    }
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

/** Source metadata is immutable; only explicitly supported presentation values may change. */
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
  const result = structuredClone(candidate)
  const verify = (node: EditorJSON, originalBlock?: EditorJSON) => {
    const tracked = editable.has(node.type) || node.type === 'docProtected'
    let original = originalBlock
    if (tracked) {
      const id = node.attrs?.officeId
      if (id != null && (typeof id !== 'string' || !originals.has(id) || seen.has(id)))
        throw new Error('Invalid or duplicate block identity')
      if (typeof id === 'string') {
        seen.add(id)
        original = originals.get(id)!
        if (original.attrs?.officeLocked) {
          if (encode(node) !== encode(original))
            throw new Error('Protected content cannot be changed')
          return
        }
        const defaults = docxSchema.nodes[node.type]!.create().attrs
        for (const [key, value] of Object.entries(node.attrs ?? {})) {
          const introducedDefault =
            original.attrs?.[key] === undefined && encode(value) === encode(defaults[key])
          if (
            !mutableAttrs.has(key) &&
            !introducedDefault &&
            encode(value) !== encode(original.attrs?.[key])
          )
            throw new Error(`Immutable attribute: ${key}`)
          if (mutableAttrs.has(key)) assertAttrValue(key, value)
        }
        if (
          node.type !== original.type &&
          !['docParagraph', 'docHeading', 'docListItem'].includes(node.type)
        )
          throw new Error('Unsupported block conversion')
      } else {
        if (!editable.has(node.type)) throw new Error('Cannot create protected objects')
        const inheritedIndex = node.attrs?.docxIndex
        const inherited = [...originals.values()].find(
          (source) =>
            !source.attrs?.officeLocked &&
            inheritedIndex != null &&
            source.attrs?.docxIndex === inheritedIndex,
        )
        node.attrs = { ...node.attrs, officeId: makeId(), officeLocked: false }
        // A split inherits source attributes in upstream schema. Retain its text/format,
        // but never create a second owner of the original XML anchor.
        for (const key of sourceAttrs) if (key in node.attrs) node.attrs[key] = null
        if (needsProtection(node)) throw new Error('Unsupported new content')
        const defaults = docxSchema.nodes[node.type]!.create().attrs
        for (const [key, value] of Object.entries(node.attrs)) {
          if (sourceAttrs.has(key) || key === 'officeId' || key === 'officeLocked') continue
          if (!mutableAttrs.has(key) && encode(value) !== encode(defaults[key]))
            throw new Error(`Unsupported new attribute: ${key}`)
          assertAttrValue(key, value)
        }
        original = inherited
      }
    }
    if (!tracked && node.type !== 'doc' && node.type !== 'text') {
      const defaults = docxSchema.nodes[node.type]!.create().attrs
      for (const [key, value] of Object.entries(node.attrs ?? {}))
        if (nestedMutable.has(key)) assertAttrValue(key, value)
      let trusted = encode(sourceOnlyAttrs(node.attrs ?? {})) === encode(sourceOnlyAttrs(defaults))
      if (original)
        walk(original, (source) => {
          if (
            source.type === node.type &&
            encode(sourceOnlyAttrs(source.attrs ?? {})) ===
              encode(sourceOnlyAttrs(node.attrs ?? {}))
          )
            trusted = true
        })
      if (!trusted) throw new Error('Untrusted nested attributes')
    }
    for (const mark of node.marks ?? []) {
      if (!simpleMarks.has(mark.type)) throw new Error('Unsupported mark')
      for (const [key, value] of Object.entries(mark.attrs ?? {})) {
        if (textStyleAttrs.has(key) && mark.type === 'docTextStyle') assertAttrValue(key, value)
        else if (value != null) {
          let trusted = false
          if (original)
            walk(original, (source) => {
              if (
                source.marks?.some(
                  (m) => m.type === mark.type && encode(m.attrs?.[key]) === encode(value),
                )
              )
                trusted = true
            })
          if (!trusted) throw new Error(`Untrusted mark attribute: ${key}`)
        }
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
