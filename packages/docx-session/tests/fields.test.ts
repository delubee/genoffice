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

it('authors whitelisted fields from Insert → Field and refuses fields that reach outside', async () => {
  const session = await DocxSession.open(
    await buildDocx({ bodyXml: '<w:p><w:r><w:t>Page </w:t></w:r></w:p>' }),
  )
  const field = (instr: string) => {
    const doc = docxSchema.nodeFromJSON(session.content())
    return new Transform(doc).insert(
      6,
      docxSchema.text('1', [docxSchema.marks.instrField!.create({ instr })]),
    )
  }
  for (const instr of ['INCLUDETEXT "c:/secret.docx"', 'DDE excel "x"', 'PAGE <w:evil/>'])
    expect(() => submit(session, field(instr)), instr).toThrow()
  submit(session, field('PAGE \\* MERGEFORMAT'))
  const zip = await JSZip.loadAsync(await session.save())
  expect(await zip.file('word/document.xml')!.async('string')).toMatch(/PAGE/)
})
