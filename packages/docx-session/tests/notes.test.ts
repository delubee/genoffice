import { it, expect } from 'vitest'
import { Transform } from '@tiptap/pm/transform'
import JSZip from 'jszip'
import { docxSchema } from '@genoffice/docs-editor/shared/schema'
import { buildDocx } from '../../docx-engine/tests/helpers/build-docx'
import { DocxSession } from '../src/session'

const submit = (session: DocxSession, tr: Transform, lists: Record<string, unknown> = {}) => {
  const descriptor = session.descriptor()
  return session.submit({
    sessionId: session.sessionId,
    epoch: descriptor.epoch,
    expectedRevision: descriptor.revision,
    baseCollabVersion: descriptor.collabVersion,
    clientId: 'browser',
    transactionId: crypto.randomUUID(),
    steps: tr.steps.map((step) => step.toJSON()),
    ...lists,
  })
}

it('inserts a footnote reference with its note in one entry and saves footnotes.xml', async () => {
  const session = await DocxSession.open(
    await buildDocx({ bodyXml: '<w:p><w:r><w:t>Claim</w:t></w:r></w:p>' }),
  )
  const doc = docxSchema.nodeFromJSON(session.content())
  const reference = () =>
    new Transform(doc).insert(
      6,
      docxSchema.nodes.docNoteRef!.create({ kind: 'footnote', id: '1', num: 1 }),
    )
  expect(() => submit(session, reference())).toThrow('Unknown note reference')
  const before = session.content()
  submit(session, reference(), { footnotes: [{ id: '1', text: 'Source: annual report' }] })
  expect(session.presentation().settings!.footnotes).toEqual([
    { id: '1', text: 'Source: annual report' },
  ])
  session.history('undo')
  expect(session.content()).toEqual(before)
  session.history('redo')
  const zip = await JSZip.loadAsync(await session.save())
  expect(await zip.file('word/footnotes.xml')!.async('string')).toContain('Source: annual report')
  expect(await zip.file('word/document.xml')!.async('string')).toContain('w:footnoteReference')
  expect(() =>
    submit(session, new Transform(docxSchema.nodeFromJSON(session.content())), {
      footnotes: [{ id: '1', text: 'x', styleId: '<evil/>' }],
    }),
  ).toThrow()
})
