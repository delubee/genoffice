import { expect, it } from 'vitest'
import JSZip from 'jszip'
import { parseDocx, saveDocx } from '@genoffice/docx-engine'
import { buildDocx, IMAGE_PARAGRAPH_XML } from '../../docx-engine/tests/helpers/build-docx'
import { DocxSession } from '../src/session'
import { docxSchema } from '@genoffice/docs-editor/shared/schema'
import { documentToEditor, editorToSavePlan } from '@genoffice/docs-editor/conversion'

it('uses the rich schema without DOM and preserves untouched package entries on a paragraph edit', async () => {
  const fixture = await buildDocx({
    bodyXml:
      '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Original</w:t></w:r></w:p><w:p><w:r><w:t>Untouched</w:t></w:r></w:p>',
  })
  const zip = await JSZip.loadAsync(fixture)
  zip.file('customXml/sentinel.xml', '<sentinel>preserved</sentinel>')
  zip.file('word/header1.xml', '<header>original</header>')
  const bytes = await zip.generateAsync({ type: 'uint8array' })
  const parsed = await parseDocx(bytes)
  const initial = documentToEditor(parsed)
  docxSchema.nodeFromJSON(initial).check()
  const unchanged = await saveDocx(parsed, editorToSavePlan(initial, parsed.blocks).saveBlocks)
  expect(Buffer.from(unchanged)).toEqual(Buffer.from(bytes))
  const edited = structuredClone(initial)
  edited.content![0]!.content![0]!.text = 'Edited'
  const plan = editorToSavePlan(edited, parsed.blocks)
  const result = await JSZip.loadAsync(await saveDocx(parsed, plan.saveBlocks))
  expect(await result.file('customXml/sentinel.xml')!.async('string')).toBe(
    '<sentinel>preserved</sentinel>',
  )
  expect(await result.file('word/header1.xml')!.async('string')).toBe('<header>original</header>')
  const output = await result.file('word/document.xml')!.async('string')
  expect(output).toContain('Edited')
  expect(output).toContain('Untouched')
  expect(output).toContain('<w:b')
})

it('keeps real image bytes, header/footer parts and private runs when editing a different paragraph', async () => {
  const bytes = await buildDocx({
    bodyXml:
      '<w:p><w:r><w:t>Editable</w:t></w:r></w:p>' +
      IMAGE_PARAGRAPH_XML +
      '<w:p><w:r><w:rPr><go:redact w:label="private"/></w:rPr><w:t>PRIVATE_SENTINEL</w:t></w:r></w:p>',
    withImage: true,
    docRootExtraAttrs: 'xmlns:go="urn:genoffice:redaction"',
    sectPrExtra:
      '<w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/>',
    extraRels:
      '<Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rIdFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>',
    extraParts: [
      {
        path: 'word/header1.xml',
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
        xml: '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>HEADER_SENTINEL</w:t></w:r></w:p></w:hdr>',
      },
      {
        path: 'word/footer1.xml',
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
        xml: '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>FOOTER_SENTINEL</w:t></w:r></w:p></w:ftr>',
      },
    ],
  })
  const session = await DocxSession.open(bytes)
  const read = session.read()
  expect(JSON.stringify(read)).not.toContain('PRIVATE_SENTINEL')
  expect(read.blocks.filter((block) => block.protected)).toHaveLength(2)
  session.applyOperations({
    transactionId: 'edit-safe-paragraph',
    expectedRevision: 0,
    operations: [
      { type: 'replaceText', blockId: read.blocks[0]!.blockId, start: 0, end: 8, text: 'Changed' },
    ],
  })
  const before = await JSZip.loadAsync(bytes),
    after = await JSZip.loadAsync(await session.save())
  for (const part of ['word/media/image1.png', 'word/header1.xml', 'word/footer1.xml'])
    expect(await after.file(part)!.async('uint8array')).toEqual(
      await before.file(part)!.async('uint8array'),
    )
  expect(await after.file('word/document.xml')!.async('string')).toContain('PRIVATE_SENTINEL')
})

it('honors enforced document protection in both the host descriptor and write path', async () => {
  const bytes = await buildDocx({
    bodyXml: '<w:p><w:r><w:t>Protected</w:t></w:r></w:p>',
    extraParts: [
      {
        path: 'word/settings.xml',
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
        xml: '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:documentProtection w:edit="readOnly" w:enforcement="1"/></w:settings>',
      },
    ],
  })
  const session = await DocxSession.open(bytes)
  expect(session.descriptor().readonly).toBe(true)
  expect(() =>
    session.applyOperations({
      transactionId: 'blocked',
      expectedRevision: 0,
      operations: [
        {
          type: 'replaceText',
          blockId: session.read().blocks[0]!.blockId,
          start: 0,
          end: 9,
          text: 'Attempt',
        },
      ],
    }),
  ).toThrow('read-only')
  expect(session.descriptor().revision).toBe(0)
})
