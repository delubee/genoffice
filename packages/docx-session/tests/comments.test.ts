import { it, expect } from 'vitest'
import { Transform } from '@tiptap/pm/transform'
import JSZip from 'jszip'
import { docxSchema } from '@genoffice/docs-editor/shared/schema'
import { buildDocx } from '../../docx-engine/tests/helpers/build-docx'
import { DocxSession } from '../src/session'

const submit = (session: DocxSession, tr: Transform, comments?: unknown) => {
  const descriptor = session.descriptor()
  return session.submit({
    sessionId: session.sessionId,
    epoch: descriptor.epoch,
    expectedRevision: descriptor.revision,
    baseCollabVersion: descriptor.collabVersion,
    clientId: 'browser',
    transactionId: crypto.randomUUID(),
    steps: tr.steps.map((step) => step.toJSON()),
    ...(comments ? { comments: comments as never } : {}),
  })
}

it('anchors a new comment with its list in one entry, resolves it and saves comments.xml', async () => {
  const session = await DocxSession.open(
    await buildDocx({ bodyXml: '<w:p><w:r><w:t>Review this sentence</w:t></w:r></w:p>' }),
  )
  const doc = docxSchema.nodeFromJSON(session.content())
  const anchor = () =>
    new Transform(doc).addMark(1, 7, docxSchema.marks.comment!.create({ ids: '1' }))
  // An anchor without its comment is refused.
  expect(() => submit(session, anchor())).toThrow('Unknown comment anchor')
  const comment = {
    id: '1',
    author: 'Reviewer',
    date: '2026-10-30T10:00:00Z',
    text: 'Check the figure',
  }
  const before = session.content()
  submit(session, anchor(), [comment])
  expect(session.presentation().settings!.comments).toEqual([comment])
  session.applySettings({
    transactionId: crypto.randomUUID(),
    expectedRevision: session.descriptor().revision,
    settings: { comments: [{ ...comment, done: true }] },
  })
  expect(session.presentation().settings!.comments![0]!.done).toBe(true)
  session.history('undo')
  session.history('undo')
  expect(session.content()).toEqual(before)
  expect(session.presentation().settings!.comments).toEqual([])
  session.history('redo')
  session.history('redo')
  const zip = await JSZip.loadAsync(await session.save())
  expect(await zip.file('word/comments.xml')!.async('string')).toContain('Check the figure')
  expect(await zip.file('word/document.xml')!.async('string')).toContain('w:commentRangeStart')
  const reopened = await DocxSession.open(await session.save())
  expect(reopened.presentation().settings!.comments![0]).toMatchObject({
    text: 'Check the figure',
    done: true,
  })
  expect(() =>
    submit(session, new Transform(docxSchema.nodeFromJSON(session.content())), [
      { id: '1', author: 'x', text: 'a' },
      { id: '1', author: 'y', text: 'b' },
    ]),
  ).toThrow()
})
