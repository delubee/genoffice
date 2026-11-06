import { describe, expect, it } from 'vitest'
import { prepareDocument, validateAndNormalize } from '@genoffice/docs-editor/shared/policy'
import type { EditorJSON } from '@genoffice/docs-editor/conversion'

let sequence = 0
const id = () => `block-${++sequence}`
const fixture = () =>
  prepareDocument(
    {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: { docxIndex: 0 },
          content: [{ type: 'text', text: 'Editable' }],
        },
        { type: 'docProtected', attrs: { docxIndex: 1, blockType: 'chart', label: 'Chart' } },
      ],
    },
    id,
  )

describe('DOCX host policy', () => {
  it('allows text edits but rejects protected modification and deletion atomically', () => {
    const before = fixture()
    const next = structuredClone(before)
    next.content![0]!.content![0]!.text = 'Changed'
    expect(validateAndNormalize(before, next, id)).toEqual(next)
    next.content![1]!.attrs!.label = 'Forged'
    expect(() => validateAndNormalize(before, next, id)).toThrow('Protected')
    next.content!.pop()
    expect(() => validateAndNormalize(before, next, id)).toThrow('removed')
  })
  it('rejects XML injection and forged or repeated identities', () => {
    const before = fixture()
    const forged = structuredClone(before)
    forged.content![0]!.attrs!.rawPPr = '<w:pPr><evil/></w:pPr>'
    expect(() => validateAndNormalize(before, forged, id)).toThrow('Immutable')
    forged.content![0] = structuredClone(before.content![0]!)
    forged.content!.push(structuredClone(forged.content![1]!))
    expect(() => validateAndNormalize(before, forged, id)).toThrow('duplicate')
  })
  it('treats a repeated identity from a mid-paragraph split as a new block without source anchors', () => {
    const before = fixture()
    const next = structuredClone(before)
    const second = structuredClone(next.content![0]!)
    next.content![0]!.content![0]!.text = 'Edi'
    second.content![0]!.text = 'table'
    next.content!.splice(1, 0, second)
    const normalized = validateAndNormalize(before, next, id)
    expect(normalized.content![0]!.attrs!.officeId).toBe(before.content![0]!.attrs!.officeId)
    expect(normalized.content![1]!.attrs!.officeId).not.toBe(before.content![0]!.attrs!.officeId)
    expect(normalized.content![1]!.attrs!.docxIndex).toBeNull()
  })
  it('gives split paragraphs fresh identities and clears inherited source anchors', () => {
    const before = fixture()
    const next = structuredClone(before)
    const split = structuredClone(next.content![0]!)
    split.attrs!.officeId = null
    split.content![0]!.text = 'Split'
    next.content!.splice(1, 0, split)
    const normalized = validateAndNormalize(before, next, id)
    expect(normalized.content![1]!.attrs!.docxIndex).toBeNull()
    expect(normalized.content![1]!.attrs!.officeId).not.toBe(before.content![0]!.attrs!.officeId)
  })
  it('protects private runs before exposing the editable projection', () => {
    const doc: EditorJSON = {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          attrs: { docxIndex: 0 },
          content: [{ type: 'text', text: 'Secret', marks: [{ type: 'redaction' }] }],
        },
      ],
    }
    const prepared = prepareDocument(doc, id)
    expect(prepared.content![0]!.type).toBe('docProtected')
    expect(JSON.stringify(prepared)).not.toContain('Secret')
  })
  it('accepts desktop-authored tables, marks and paragraph formats within value bounds', () => {
    const before = fixture()
    const next = structuredClone(before)
    next.content![0]!.attrs!.shadingFill = 'FFFF00'
    next.content![0]!.attrs!.borders = { top: { style: 'single', sz: 4, color: '000000' } }
    next.content![0]!.content![0]!.marks = [
      { type: 'docTextStyle', attrs: { highlight: 'yellow' } },
      { type: 'link', attrs: { href: 'https://example.com' } },
    ]
    next.content!.splice(1, 0, {
      type: 'docTable',
      attrs: { borders: { insideH: { style: 'single', sz: 4 } }, widthPct: 100 },
      content: [
        {
          type: 'docTableRow',
          content: [
            {
              type: 'docTableCell',
              attrs: { colwidth: [120] },
              content: [{ type: 'docParagraph' }],
            },
          ],
        },
      ],
    })
    const normalized = validateAndNormalize(before, next, id)
    expect(normalized.content![1]!.attrs!.officeId).toBeTruthy()
    expect(normalized.content![0]!.attrs!.shadingFill).toBe('FFFF00')
  })
  it('never lets the view author raw XML, source anchors or executable links', () => {
    const before = fixture()
    const table = (cellAttrs: Record<string, unknown>) => ({
      type: 'docTable',
      content: [
        {
          type: 'docTableRow',
          content: [
            { type: 'docTableCell', attrs: cellAttrs, content: [{ type: 'docParagraph' }] },
          ],
        },
      ],
    })
    const withTable = structuredClone(before)
    withTable.content!.splice(1, 0, table({ rawTcPr: '<w:tcPr><evil/></w:tcPr>' }))
    expect(() => validateAndNormalize(before, withTable, id)).toThrow('Untrusted nested attribute')
    const link = structuredClone(before)
    link.content![0]!.content![0]!.marks = [
      { type: 'link', attrs: { href: 'javascript:alert(1)' } },
    ]
    expect(() => validateAndNormalize(before, link, id)).toThrow('Invalid link')
    const field = structuredClone(before)
    field.content![0]!.content![0]!.marks = [
      { type: 'instrField', attrs: { instr: 'INCLUDETEXT "c:/secret"' } },
    ]
    expect(() => validateAndNormalize(before, field, id)).toThrow()
    const anchor = structuredClone(before)
    anchor.content![0]!.attrs!.docxIndex = 1
    expect(() => validateAndNormalize(before, anchor, id)).toThrow('Immutable')
    const fresh = structuredClone(before)
    fresh.content!.push({
      type: 'docParagraph',
      attrs: { docxIndex: 0, rawPPr: '<w:pPr/>' },
    } as never)
    const normalized = validateAndNormalize(before, fresh, id)
    expect(normalized.content![2]!.attrs!.docxIndex).toBeNull()
  })
  it('lets Picture Format resize a file picture but never move its source anchor', () => {
    const before = prepareDocument(
      {
        type: 'doc',
        content: [
          {
            type: 'docProtected',
            attrs: { docxIndex: 0, blockType: 'image', label: 'Picture', imageWidthPx: 100 },
          },
        ],
      },
      id,
    )
    expect(before.content![0]!.attrs!.officeLocked).toBe(false)
    const resized = structuredClone(before)
    resized.content![0]!.attrs!.imageWidthPx = 240
    resized.content![0]!.attrs!.imageRotDeg = 90
    expect(validateAndNormalize(before, resized, id).content![0]!.attrs!.imageWidthPx).toBe(240)
    const moved = structuredClone(before)
    moved.content![0]!.attrs!.docxIndex = 7
    expect(() => validateAndNormalize(before, moved, id)).toThrow('Immutable')
  })
})
