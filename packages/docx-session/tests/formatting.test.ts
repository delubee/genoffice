import { it, expect } from 'vitest'
import JSZip from 'jszip'
import { parseDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../docx-engine/tests/helpers/build-docx'
import { DocxSession } from '../src/session'

it('creates real heading styles and numbering definitions, and keeps both overlays in history', async () => {
  const session = await DocxSession.open(
    await buildDocx({
      bodyXml: '<w:p><w:r><w:t>Heading</w:t></w:r></w:p><w:p><w:r><w:t>Item</w:t></w:r></w:p>',
      stylesXml:
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>',
    }),
  )
  const blocks = session.read().blocks
  session.applyOperations({
    transactionId: 'heading',
    expectedRevision: 0,
    operations: [{ type: 'setHeading', blockId: blocks[0]!.blockId, level: 1 }],
  })
  session.applyOperations({
    transactionId: 'list',
    expectedRevision: 1,
    operations: [{ type: 'setList', blockId: blocks[1]!.blockId, kind: 'ordered' }],
  })
  const saved = await session.save()
  const reopened = await parseDocx(saved)
  expect(reopened.blocks[0]!.type).toBe('heading')
  expect(reopened.blocks[0]!.level).toBe(1)
  expect(reopened.blocks[1]!.type).toBe('listItem')
  expect(reopened.numbering.has(reopened.blocks[1]!.list!.numId)).toBe(true)
  const zip = await JSZip.loadAsync(saved)
  expect(await zip.file('word/styles.xml')!.async('string')).toContain('GenOfficeHeading1')
  session.history('undo')
  expect((await JSZip.loadAsync(await session.save())).file('word/numbering.xml')).toBeNull()
  session.history('redo')
  expect((await parseDocx(await session.save())).blocks[1]!.type).toBe('listItem')
})
