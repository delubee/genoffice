import { it, expect } from 'vitest'
import { Transform } from '@tiptap/pm/transform'
import {
  generateCaptionXml,
  generateIndexFieldXml,
  generateTocFieldXml,
} from '@genoffice/docx-engine'
import { docxSchema } from '@genoffice/docs-editor/shared/schema'
import { buildDocx } from '../../docx-engine/tests/helpers/build-docx'
import { DocxSession } from '../src/session'

const block = (genXml: string) =>
  docxSchema.nodes.docProtected!.create({ blockType: 'passthrough', label: 'Field', genXml })

const submit = async (genXml: string) => {
  const session = await DocxSession.open(
    await buildDocx({ bodyXml: '<w:p><w:r><w:t>Body</w:t></w:r></w:p>' }),
  )
  const doc = docxSchema.nodeFromJSON(session.content())
  const descriptor = session.descriptor()
  session.submit({
    sessionId: session.sessionId,
    epoch: descriptor.epoch,
    expectedRevision: descriptor.revision,
    baseCollabVersion: descriptor.collabVersion,
    clientId: 'browser',
    transactionId: crypto.randomUUID(),
    steps: new Transform(doc).insert(0, block(genXml)).steps.map((step) => step.toJSON()),
  })
  return session
}

it('accepts generated TOC, index and caption field paragraphs and refuses anything else', async () => {
  const toc = generateTocFieldXml([
    { text: 'Intro & <scope>', level: 1, pageNo: 2 },
    { text: 'Detail', level: 2 },
  ] as never)
  for (const xml of [
    ...toc,
    ...generateIndexFieldXml(['Alpha', 'Beta & co']),
    generateCaptionXml('Figure', 3, 'Architecture'),
  ])
    expect((await submit(xml)).save()).resolves.toBeInstanceOf(Uint8Array)
  for (const forged of [
    toc[0]!.replace('TOC \\o', 'INCLUDETEXT "x" \\o'),
    toc[1]!.replace('Detail', 'Detail<w:evil/>'),
    '<w:p><w:r><w:t>raw</w:t></w:r></w:p>',
    generateCaptionXml('Figure', 3, 'x').replace('SEQ Figure', 'SEQ Table'),
  ])
    await expect(submit(forged)).rejects.toThrow()
})

it('accepts equations whose OMML is derived from their LaTeX and refuses forged math', async () => {
  const { equationBlockJson, inlineEquationNodeJson } =
    await import('../../../apps/docs/src/renderer/editor/equation')
  const session = await DocxSession.open(
    await buildDocx({ bodyXml: '<w:p><w:r><w:t>Inline </w:t></w:r></w:p>' }),
  )
  const insert = (node: unknown) => {
    const doc = docxSchema.nodeFromJSON(session.content())
    const descriptor = session.descriptor()
    return session.submit({
      sessionId: session.sessionId,
      epoch: descriptor.epoch,
      expectedRevision: descriptor.revision,
      baseCollabVersion: descriptor.collabVersion,
      clientId: 'browser',
      transactionId: crypto.randomUUID(),
      steps: new Transform(doc)
        .insert(
          (node as { type: string }).type === 'docProtected' ? 0 : 8,
          docxSchema.nodeFromJSON(node as never),
        )
        .steps.map((step) => step.toJSON()),
    })
  }
  insert(equationBlockJson('x^2 + y^2 = z^2'))
  const forgedBlock = equationBlockJson('x') as { attrs: Record<string, unknown> }
  expect(() =>
    insert({ ...forgedBlock, attrs: { ...forgedBlock.attrs, genXml: '<w:p><w:evil/></w:p>' } }),
  ).toThrow()
  const inline = inlineEquationNodeJson('a/b') as { attrs: Record<string, unknown> }
  expect(() =>
    insert({ ...inline, attrs: { ...inline.attrs, omml: '<m:oMath><w:evil/></m:oMath>' } }),
  ).toThrow()
  insert(inline)
  const xml = await (
    await (await import('jszip')).default.loadAsync(await session.save())
  )
    .file('word/document.xml')!
    .async('string')
  expect(xml).toContain('<m:oMath>')
})
