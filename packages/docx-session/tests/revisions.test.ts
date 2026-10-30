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

it('records a tracked insertion as w:ins and accepting it clears the revision', async () => {
  const session = await DocxSession.open(
    await buildDocx({ bodyXml: '<w:p><w:r><w:t>Draft</w:t></w:r></w:p>' }),
  )
  let doc = docxSchema.nodeFromJSON(session.content())
  const ins = docxSchema.marks.ins!.create({
    author: 'Reviewer',
    date: '2026-10-30T10:00:00Z',
    id: '1',
  })
  submit(session, new Transform(doc).insert(6, docxSchema.text(' text', [ins])))
  let zip = await JSZip.loadAsync(await session.save())
  expect(await zip.file('word/document.xml')!.async('string')).toMatch(
    /<w:ins\b[^>]*w:author="Reviewer"/,
  )
  doc = docxSchema.nodeFromJSON(session.content())
  submit(session, new Transform(doc).removeMark(1, doc.content.size - 1, docxSchema.marks.ins))
  zip = await JSZip.loadAsync(await session.save())
  const xml = await zip.file('word/document.xml')!.async('string')
  expect(xml).not.toContain('<w:ins')
  expect(xml).toContain('Draft text')
})
