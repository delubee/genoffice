import { it, expect } from 'vitest'
import { Transform } from '@tiptap/pm/transform'
import JSZip from 'jszip'
import { docxSchema } from '@genoffice/docs-editor/shared/schema'
import { buildDocx } from '../../docx-engine/tests/helpers/build-docx'
import { DocxSession } from '../src/session'

it('creates a style from the Styles pane, applies it in the same entry and saves styles.xml', async () => {
  const session = await DocxSession.open(
    await buildDocx({ bodyXml: '<w:p><w:r><w:t>Callout</w:t></w:r></w:p>' }),
  )
  const doc = docxSchema.nodeFromJSON(session.content())
  const node = doc.firstChild!
  const style = {
    styleId: 'Callout',
    type: 'paragraph' as const,
    name: 'Callout',
    rPr: { bold: true, color: 'C00000' },
    pPr: { spaceAfterTwips: 240 },
  }
  const tr = new Transform(doc).setNodeMarkup(0, undefined, { ...node.attrs, styleId: 'Callout' })
  const request = (styles?: unknown) => {
    const descriptor = session.descriptor()
    return {
      sessionId: session.sessionId,
      epoch: descriptor.epoch,
      expectedRevision: descriptor.revision,
      baseCollabVersion: descriptor.collabVersion,
      clientId: 'browser',
      transactionId: crypto.randomUUID(),
      steps: tr.steps.map((step) => step.toJSON()),
      ...(styles ? { styles: styles as never } : {}),
    }
  }
  expect(() => session.submit(request())).toThrow('Unknown paragraph style')
  session.submit(request([style]))
  await session.stylesReady()
  expect(new Map(session.presentation().styles).get('Callout')?.name).toBe('Callout')
  // A list-only edit modifies the definition as its own entry.
  session.applySettings({
    transactionId: crypto.randomUUID(),
    expectedRevision: session.descriptor().revision,
    settings: { styles: [{ styleId: 'Callout', rPr: { italic: true } }] },
  })
  const zip = await JSZip.loadAsync(await session.save())
  const stylesXml = await zip.file('word/styles.xml')!.async('string')
  expect(stylesXml).toMatch(/w:styleId="Callout"/)
  expect(stylesXml).toMatch(/<w:i\/>|<w:i /)
  expect(() =>
    session.applySettings({
      transactionId: crypto.randomUUID(),
      expectedRevision: session.descriptor().revision,
      settings: { styles: [{ styleId: 'Evil', name: '"><w:evil/>' }] },
    }),
  ).toThrow()
})
