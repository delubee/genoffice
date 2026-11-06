import { expect, it } from 'vitest'
import JSZip from 'jszip'
import { Transform } from '@tiptap/pm/transform'
import { docxSchema } from '@genoffice/docs-editor/shared/schema'
import { parseDocx } from '@genoffice/docx-engine'
import { buildDocx } from '../../docx-engine/tests/helpers/build-docx'
import { DocxSession } from '../src/session'

function format(session: DocxSession, attrs: Record<string, unknown>, from = 1, to = 6) {
  const document = docxSchema.nodeFromJSON(session.content())
  const style = document.nodeAt(from)!.marks.find((mark) => mark.type.name === 'docTextStyle')
  const transform = new Transform(document).addMark(
    from,
    to,
    docxSchema.marks.docTextStyle!.create({ ...style?.attrs, ...attrs }),
  )
  const descriptor = session.descriptor()
  return session.submit({
    sessionId: session.sessionId,
    epoch: descriptor.epoch,
    expectedRevision: descriptor.revision,
    baseCollabVersion: descriptor.collabVersion,
    clientId: 'font-dialog',
    transactionId: crypto.randomUUID(),
    steps: transform.steps.map((step) => step.toJSON()),
  })
}

it('saves all desktop Font dialog fields through manual steps and shared history', async () => {
  const session = await DocxSession.open(
    await buildDocx({
      bodyXml:
        '<w:p><w:r><w:t>Alpha</w:t></w:r><w:r><w:rPr><w:color w:val="008800"/></w:rPr><w:t>Beta</w:t></w:r></w:p>',
    }),
  )
  const before = session.content()
  format(session, {
    fontAscii: 'Georgia',
    font: 'SimSun',
    eastAsiaFont: 'SimSun',
    eaSlotEmpty: false,
    sizeHalfPoints: 28,
    color: '123456',
    vertAlign: 'superscript',
    dstrike: true,
    caps: 'small',
    vanish: true,
    vanishOwn: true,
    charScalePct: 80,
    charScaleEm: -0.2,
    charSpacingTwips: 30,
    positionHalfPoints: 6,
    kernHalfPoints: 24,
    kern: true,
  })
  const edited = session.content()
  const bytes = await session.save()
  const zip = await JSZip.loadAsync(bytes)
  const xml = await zip.file('word/document.xml')!.async('string')
  for (const tag of [
    'rFonts',
    'sz',
    'vertAlign',
    'dstrike',
    'smallCaps',
    'vanish',
    'w',
    'spacing',
    'position',
    'kern',
  ])
    expect(xml).toContain(`<w:${tag}`)
  const reopened = await parseDocx(bytes)
  expect(reopened.blocks[0]!.runs![0]).toMatchObject({
    fontAscii: 'Georgia',
    eastAsiaFont: 'SimSun',
    sizeHalfPoints: 28,
    color: '123456',
    vertAlign: 'superscript',
    dstrike: true,
    caps: 'small',
    vanish: true,
    charScalePct: 80,
    charSpacingTwips: 30,
    positionHalfPoints: 6,
    kernHalfPoints: 24,
  })
  expect(reopened.blocks[0]!.runs![1]).toMatchObject({ text: 'Beta', color: '008800' })
  session.history('undo')
  expect(session.content()).toEqual(before)
  session.history('redo')
  expect(session.content()).toEqual(edited)
  // Explicit off values must override the previously authored effects.
  format(session, {
    dstrike: false,
    caps: 'none',
    vanish: false,
    vanishOwn: false,
    vertAlign: null,
    charScalePct: null,
    charScaleEm: null,
    charSpacingTwips: null,
    positionHalfPoints: null,
    kernHalfPoints: 0,
    kern: false,
  })
  const cleared = await parseDocx(await session.save())
  expect(cleared.blocks[0]!.runs![0]!.dstrike).not.toBe(true)
  expect(cleared.blocks[0]!.runs![0]!.vanish).not.toBe(true)
  expect(cleared.blocks[0]!.runs![0]!.caps).not.toBe('small')
  expect(cleared.blocks[0]!.runs![0]!.charScalePct).toBeUndefined()
})

it('rejects malformed font fields atomically while preserving unchanged foreign formatting', async () => {
  const session = await DocxSession.open(
    await buildDocx({
      bodyXml: '<w:p><w:r><w:rPr><w:w w:val="700"/></w:rPr><w:t>Alpha</w:t></w:r></w:p>',
    }),
  )
  format(session, { color: 'AABBCC' })
  const before = session.content()
  const revision = session.descriptor().revision
  for (const attrs of [
    { charScalePct: 0 },
    { charScalePct: 601 },
    { charScaleEm: 'url(bad)' },
    { charSpacingTwips: '20' },
    { positionHalfPoints: Infinity },
    { kernHalfPoints: -1 },
    { kern: 'normal' },
    { caps: '<xml/>' },
    { vanishOwn: 'true' },
    { vertAlign: 'url(bad)' },
    { eaSlotEmpty: 1 },
  ]) {
    expect(() => format(session, attrs), JSON.stringify(attrs)).toThrow()
    expect(session.content()).toEqual(before)
    expect(session.descriptor().revision).toBe(revision)
  }
})
