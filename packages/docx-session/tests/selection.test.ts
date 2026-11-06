import { expect, it } from 'vitest'
import { DocxSession } from '../src/session'
import { DocumentBindings } from '../src/bindings'
import { buildDocx } from '../../docx-engine/tests/helpers/build-docx'

it('projects host-validated UTF-16 ranges and rejects stale or invalid selection claims', async () => {
  const session = await DocxSession.open(
    await buildDocx({ bodyXml: '<w:p><w:r><w:t>A😀BC</w:t></w:r></w:p>' }),
  )
  const descriptor = session.descriptor()
  const selection = session.selection({ ...descriptor, from: 2, to: 4 })
  expect(selection.blocks[0]!.text).toBe('😀')
  expect(selection.blocks[0]).toMatchObject({ start: 1, end: 3 })
  expect(() => session.selection({ ...descriptor, from: 0, to: 999 })).toThrow('positions')
  const bindings = new DocumentBindings()
  const id = bindings.bind('file:///a.docx', session.sessionId, selection)
  selection.blocks[0]!.text = 'Caller mutation'
  expect(bindings.resolve(id).selection!.blocks[0]!.text).toBe('😀')
  session.applyOperations({
    transactionId: 'change',
    expectedRevision: 0,
    operations: [
      {
        type: 'replaceText',
        blockId: session.read().blocks[0]!.blockId,
        start: 0,
        end: 1,
        text: 'Z',
      },
    ],
  })
  expect(() => session.selection({ ...descriptor, from: 2, to: 4 })).toThrow('Stale')
  bindings.release('file:///a.docx')
  expect(() => bindings.resolve(id)).toThrow('Unknown')
})
