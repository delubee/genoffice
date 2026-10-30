import { it, expect } from 'vitest'
import { Transform } from '@tiptap/pm/transform'
import JSZip from 'jszip'
import { parseDocx } from '@genoffice/docx-engine'
import { docxSchema } from '@genoffice/docs-editor/shared/schema'
import { buildDocx } from '../../docx-engine/tests/helpers/build-docx'
import { DocxSession } from '../src/session'

// 1×1 transparent PNG
const PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

const picture = (attrs: Record<string, unknown> = {}) =>
  docxSchema.nodes.docProtected!.create({
    blockType: 'image',
    label: 'Picture',
    imageDataUrl: `data:image/png;base64,${PNG}`,
    imageWidthPx: 40,
    imageHeightPx: 30,
    genImage: { base64: PNG, mime: 'image/png', widthPx: 40, heightPx: 30 },
    ...attrs,
  })

const submitter = (session: DocxSession) => (tr: Transform) => {
  const descriptor = session.descriptor()
  return session.submit({
    sessionId: session.sessionId,
    epoch: descriptor.epoch,
    expectedRevision: descriptor.revision,
    baseCollabVersion: descriptor.collabVersion,
    clientId: 'browser',
    transactionId: `picture-${descriptor.revision}`,
    steps: tr.steps.map((step) => step.toJSON()),
  })
}

it('inserts a picture from Insert → Pictures and saves it as a new media part', async () => {
  const session = await DocxSession.open(
    await buildDocx({ bodyXml: '<w:p><w:r><w:t>Before picture</w:t></w:r></w:p>' }),
  )
  const doc = docxSchema.nodeFromJSON(session.content())
  submitter(session)(new Transform(doc).insert(doc.content.size, picture()))
  const saved = await session.save()
  const zip = await JSZip.loadAsync(saved)
  expect(Object.keys(zip.files).some((name) => /^word\/media\/.+\.png$/.test(name))).toBe(true)
  const reopened = await parseDocx(saved)
  expect(reopened.blocks.some((block) => JSON.stringify(block).includes('image'))).toBe(true)
})

it('refuses forged picture blocks that claim source XML or mismatched bytes', async () => {
  const session = await DocxSession.open(
    await buildDocx({ bodyXml: '<w:p><w:r><w:t>Before picture</w:t></w:r></w:p>' }),
  )
  const doc = docxSchema.nodeFromJSON(session.content())
  const submit = submitter(session)
  for (const forged of [
    picture({ imageDataUrl: 'data:image/png;base64,AAAA' }),
    picture({ genImage: { base64: PNG, mime: 'image/svg+xml', widthPx: 40, heightPx: 30 } }),
    picture({ blockType: 'chart' }),
  ])
    expect(() => submit(new Transform(doc).insert(doc.content.size, forged))).toThrow()
})
