import { expect, it } from 'vitest'
import { Transform } from '@tiptap/pm/transform'
import { docxSchema } from '@genoffice/docs-editor/shared/schema'
import { parseDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../docx-engine/tests/helpers/build-docx'
import { DocxSession } from '../src/session'

function paragraph(session: DocxSession, attrs: Record<string, unknown>, outline?: number) {
  const doc = docxSchema.nodeFromJSON(session.content())
  const transform = new Transform(doc).setNodeMarkup(
    0,
    outline ? docxSchema.nodes.docHeading : undefined,
    {
      ...doc.firstChild!.attrs,
      ...(outline ? { level: outline, outlineOnly: true } : {}),
      ...attrs,
    },
  )
  const descriptor = session.descriptor()
  return session.submit({
    sessionId: session.sessionId,
    epoch: descriptor.epoch,
    expectedRevision: descriptor.revision,
    baseCollabVersion: descriptor.collabVersion,
    clientId: 'paragraph-dialog',
    transactionId: crypto.randomUUID(),
    steps: transform.steps.map((step) => step.toJSON()),
  })
}

it('persists paragraph dialog spacing, indents, pagination flags and outline without restyling', async () => {
  const session = await DocxSession.open(
    await buildDocx({
      bodyXml:
        '<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:t>Paragraph</w:t></w:r></w:p>' +
        '<w:p><w:r><w:t>Untouched</w:t></w:r></w:p>',
    }),
  )
  const initial = session.content()
  paragraph(
    session,
    {
      align: 'right',
      indentLeft: 720,
      indentRight: 360,
      indentFirstLine: 440,
      indentFirstLineChars: 200,
      spaceBefore: 120,
      spaceAfter: 240,
      spaceBeforeAuto: false,
      spaceAfterAuto: false,
      lineSpacing: null,
      lineRule: 'exact',
      lineRawTwips: 360,
      widowControl: false,
      keepNext: true,
      keepLines: true,
      suppressLineNumbers: true,
      contextualSpacing: true,
      pageBreakBefore: true,
    },
    3,
  )
  const edited = session.content()
  const reopened = await parseDocx(await session.save())
  expect(reopened.blocks[0]!.format).toMatchObject({
    align: 'right',
    indentLeft: 720,
    indentRight: 360,
    directCharIndents: { firstLine: 200 },
    spaceBefore: 120,
    spaceAfter: 240,
    lineRule: 'exact',
    lineRawTwips: 360,
    widowControl: false,
    keepNext: true,
    keepLines: true,
    suppressLineNumbers: true,
    contextualSpacing: true,
    pageBreakBefore: true,
  })
  expect(reopened.blocks[0]!.styleId).toBe('Normal')
  expect(reopened.blocks[0]!.level).toBe(3)
  expect(session.content().content![1]).toEqual(initial.content![1])
  session.history('undo')
  expect(session.content()).toEqual(initial)
  session.history('redo')
  expect(session.content()).toEqual(edited)
  paragraph(session, {
    keepNext: false,
    keepLines: false,
    pageBreakBefore: false,
    contextualSpacing: false,
    widowControl: true,
    lineRule: null,
    lineRawTwips: null,
    lineSpacing: 1.5,
    indentFirstLine: null,
    indentFirstLineChars: null,
  })
  const cleared = (await parseDocx(await session.save())).blocks[0]!.format!
  expect(cleared).toMatchObject({
    keepNext: false,
    keepLines: false,
    contextualSpacing: false,
    widowControl: true,
    lineSpacing: 1.5,
  })
  expect(cleared.pageBreakBefore).not.toBe(true)
  expect(cleared.lineRule).not.toBe('exact')
  expect(cleared.directCharIndents?.firstLine).toBeUndefined()
})

it('rejects malformed paragraph fields before changing the authoritative document', async () => {
  const session = await DocxSession.open(
    await buildDocx({
      bodyXml: '<w:p><w:r><w:t>Paragraph</w:t></w:r></w:p>',
    }),
  )
  const before = session.content()
  for (const attrs of [
    { lineSpacing: 0 },
    { lineRule: 'injected' },
    { lineRawTwips: -10 },
    { indentFirstLine: '720' },
    { indentFirstLineChars: 1.5 },
    { keepNext: 'true' },
    { pageBreakBefore: 1 },
    { outlineOnly: 'true' },
  ]) {
    expect(
      () => paragraph(session, attrs, 'outlineOnly' in attrs ? 3 : undefined),
      JSON.stringify(attrs),
    ).toThrow()
    expect(session.content()).toEqual(before)
    expect(session.descriptor().revision).toBe(0)
  }
})
