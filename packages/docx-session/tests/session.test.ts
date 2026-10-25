import { expect, it } from 'vitest'
import { Transform } from '@tiptap/pm/transform'
import { docxSchema } from '@genoffice/docs-editor/shared/schema'
import { parseDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../docx-engine/tests/helpers/build-docx'
import { DocxSession, type StepRequest } from '../src/session'

async function open() {
  return DocxSession.open(await buildDocx({ bodyXml: '<w:p><w:r><w:t>Start</w:t></w:r></w:p>' }))
}

it('tracks saved roots across undo/redo and marks recovered content unsaved', async () => {
  const session = await open()
  expect(session.descriptor().dirty).toBe(false)
  session.submit(request(session, 'Saved '))
  session.markSaved()
  expect(session.descriptor().dirty).toBe(false)
  session.submit(request(session, 'Unsaved '))
  expect(session.descriptor().dirty).toBe(true)
  session.history('undo')
  expect(session.descriptor().dirty).toBe(false)
  session.history('undo')
  expect(session.descriptor().dirty).toBe(true)
  session.history('redo')
  expect(session.descriptor().dirty).toBe(false)
  session.markUnsaved()
  expect(session.descriptor().dirty).toBe(true)
})
function request(session: DocxSession, text: string): StepRequest {
  const state = session.descriptor()
  const doc = docxSchema.nodeFromJSON(session.content())
  const tr = new Transform(doc).insert(1, docxSchema.text(text))
  return {
    sessionId: state.sessionId,
    epoch: state.epoch,
    transactionId: `edit-${state.revision}`,
    expectedRevision: state.revision,
    baseCollabVersion: state.collabVersion,
    clientId: 'browser',
    steps: tr.steps.map((step) => step.toJSON()),
  }
}

it('retries idempotently, rejects reused IDs and stale writes without changing state', async () => {
  const session = await open()
  const edit = request(session, 'One')
  const first = session.submit(edit)
  expect(session.submit(edit)).toEqual(first)
  expect(session.descriptor().revision).toBe(1)
  expect(() => session.submit({ ...edit, clientId: 'other' })).toThrow('reused')
  expect(() => session.submit({ ...edit, transactionId: 'stale' })).toThrow('Stale')
  expect(docxSchema.nodeFromJSON(session.content()).textContent).toBe('OneStart')
})

it('keeps twenty edits undoable across saves and resets collaboration epochs on history', async () => {
  const session = await open()
  const start = session.content()
  for (let index = 0; index < 20; index++) {
    if (index % 2 === 0) session.submit(request(session, `${index},`))
    else
      session.applyOperations({
        transactionId: `tool-${index}`,
        expectedRevision: session.descriptor().revision,
        operations: [
          {
            type: 'replaceText',
            blockId: session.read().blocks[0]!.blockId,
            start: 0,
            end: 0,
            text: `${index},`,
          },
        ],
      })
    if (index === 9) await session.save()
  }
  const end = session.content()
  const saved = await parseDocx(await session.save())
  expect(saved.blocks[0]!.runs!.map((run) => run.text).join('')).toContain('19,18,17,')
  const stale = request(session, 'Late')
  for (let index = 0; index < 20; index++) session.history('undo')
  expect(session.content()).toEqual(start)
  expect(session.descriptor().epoch).toBe(20)
  expect(() => session.submit(stale)).toThrow('Stale')
  for (let index = 0; index < 20; index++) session.history('redo')
  expect(session.content()).toEqual(end)
})

it('acknowledges paragraph splits with additional host identity-normalization steps', async () => {
  const session = await open()
  const state = session.descriptor()
  const doc = docxSchema.nodeFromJSON(session.content())
  const tr = new Transform(doc).split(3, 1, [
    { type: doc.firstChild!.type, attrs: { ...doc.firstChild!.attrs, officeId: null } },
  ])
  const result = session.submit({
    sessionId: state.sessionId,
    epoch: state.epoch,
    transactionId: 'split',
    expectedRevision: state.revision,
    baseCollabVersion: 0,
    clientId: 'browser',
    steps: tr.steps.map((step) => step.toJSON()),
  })
  expect(result.clientIds).toContain('host')
  const blocks = session.content().content!
  expect(blocks[1]!.attrs!.officeId).not.toBe(blocks[0]!.attrs!.officeId)
  expect(blocks[1]!.attrs!.docxIndex).toBeNull()
  expect(
    (await parseDocx(await session.save())).blocks
      .filter((block) => block.type === 'paragraph')
      .map((block) => block.runs!.map((run) => run.text).join('')),
  ).toEqual(['St', 'art'])
})
