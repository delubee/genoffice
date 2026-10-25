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
        { type: 'docProtected', attrs: { docxIndex: 1, blockType: 'image', label: 'Picture' } },
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
    forged.content!.push(structuredClone(forged.content![0]!))
    expect(() => validateAndNormalize(before, forged, id)).toThrow('duplicate')
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
})
