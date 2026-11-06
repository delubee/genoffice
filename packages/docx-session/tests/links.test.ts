import { it, expect } from 'vitest'
import { Transform } from '@tiptap/pm/transform'
import JSZip from 'jszip'
import { docxSchema } from '@genoffice/docs-editor/shared/schema'
import { buildDocx } from '../../docx-engine/tests/helpers/build-docx'
import { DocxSession } from '../src/session'

const submit = (session: DocxSession, tr: Transform) => {
  const descriptor = session.descriptor()
  return session.submit({
    sessionId: session.sessionId,
    epoch: descriptor.epoch,
    expectedRevision: descriptor.revision,
    baseCollabVersion: descriptor.collabVersion,
    clientId: 'browser',
    transactionId: crypto.randomUUID(),
    steps: tr.steps.map((step) => step.toJSON()),
  })
}

it('saves a link authored in the view as a hyperlink relationship and refuses script links', async () => {
  const session = await DocxSession.open(
    await buildDocx({ bodyXml: '<w:p><w:r><w:t>Visit example</w:t></w:r></w:p>' }),
  )
  const doc = docxSchema.nodeFromJSON(session.content())
  const link = (href: string) =>
    new Transform(doc).addMark(7, 14, docxSchema.marks.link!.create({ href }))
  expect(() => submit(session, link('javascript:alert(1)'))).toThrow('Invalid link')
  submit(session, link('https://example.com/docs'))
  const zip = await JSZip.loadAsync(await session.save())
  const documentXml = await zip.file('word/document.xml')!.async('string')
  expect(documentXml).toMatch(/<w:hyperlink\b[^>]*r:id="/)
  const rels = await zip.file('word/_rels/document.xml.rels')!.async('string')
  expect(rels).toContain('Target="https://example.com/docs"')
  expect(rels).toContain('TargetMode="External"')
})
