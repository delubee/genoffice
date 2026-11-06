import { it, expect } from 'vitest'
import { Transform } from '@tiptap/pm/transform'
import JSZip from 'jszip'
import { docxSchema } from '@genoffice/docs-editor/shared/schema'
import type { PendingListDefinitions } from '@genoffice/docs-editor/shared/protocol'
import { buildDocx } from '../../docx-engine/tests/helpers/build-docx'
import { DocxSession } from '../src/session'

const submitWith = (session: DocxSession, tr: Transform, numbering?: unknown) => {
  const descriptor = session.descriptor()
  return session.submit({
    sessionId: session.sessionId,
    epoch: descriptor.epoch,
    expectedRevision: descriptor.revision,
    baseCollabVersion: descriptor.collabVersion,
    clientId: 'browser',
    transactionId: `list-${descriptor.revision}-${Math.random()}`,
    steps: tr.steps.map((step) => step.toJSON()),
    ...(numbering ? { numbering: numbering as PendingListDefinitions } : {}),
  })
}

const listOf = (session: DocxSession, numId: string) => {
  const doc = docxSchema.nodeFromJSON(session.content())
  const node = doc.firstChild!
  return new Transform(doc).setNodeMarkup(0, docxSchema.nodes.docListItem, {
    ...node.attrs,
    numId,
    kind: 'ordered',
    ilvl: 0,
  })
}

it('commits Ribbon list definitions with the steps that use them and saves them to numbering.xml', async () => {
  const session = await DocxSession.open(
    await buildDocx({ bodyXml: '<w:p><w:r><w:t>First item</w:t></w:r></w:p>' }),
  )
  const before = session.content()
  // Without the definition the host refuses an unknown numId.
  expect(() => submitWith(session, listOf(session, '7'))).toThrow('Unknown numbering definition')
  const definitions = {
    newDefs: [
      {
        numId: '7',
        kind: 'ordered',
        levels: [{ numFmt: 'upperRoman', lvlText: '%1)', indentLeft: 720, hanging: 360 }],
      },
    ],
    restartNums: [],
  }
  const commit = submitWith(session, listOf(session, '7'), definitions)
  expect(new Map(commit.presentation!.numbering).get('7')?.levels[0]?.numFmt).toBe('upperRoman')
  session.history('undo')
  expect(session.content()).toEqual(before)
  expect(new Map(session.presentation().numbering).has('7')).toBe(false)
  session.history('redo')
  const zip = await JSZip.loadAsync(await session.save())
  const numberingXml = await zip.file('word/numbering.xml')!.async('string')
  expect(numberingXml).toContain('w:numId="7"')
  expect(numberingXml).toContain('upperRoman')
})

it('refuses list definitions that would inject XML or take an existing identity', async () => {
  const session = await DocxSession.open(
    await buildDocx({ bodyXml: '<w:p><w:r><w:t>Item</w:t></w:r></w:p>' }),
  )
  for (const numbering of [
    {
      newDefs: [
        {
          numId: '9',
          kind: 'bullet',
          levels: [{ numFmt: 'bullet', lvlText: '"/><evil/>', indentLeft: 720 }],
        },
      ],
      restartNums: [],
    },
    { newDefs: [{ numId: '0', kind: 'bullet' }], restartNums: [] },
    { newDefs: [], restartNums: [{ numId: '11', abstractNumId: '99', startOverrides: {} }] },
  ])
    expect(() => submitWith(session, listOf(session, '9'), numbering)).toThrow()
})

it('adjusts list indents of a session list as its own entry and refuses unknown lists', async () => {
  const session = await DocxSession.open(
    await buildDocx({ bodyXml: '<w:p><w:r><w:t>Item</w:t></w:r></w:p>' }),
  )
  submitWith(session, listOf(session, '7'), {
    newDefs: [
      {
        numId: '7',
        kind: 'ordered',
        levels: [{ numFmt: 'decimal', lvlText: '%1.', indentLeft: 720, hanging: 360 }],
      },
    ],
    restartNums: [],
  })
  const edit = (abstractNumId: string) =>
    session.applySettings({
      transactionId: crypto.randomUUID(),
      expectedRevision: session.descriptor().revision,
      settings: {
        numbering: {
          newDefs: [],
          restartNums: [],
          levelEdits: [
            {
              abstractNumId,
              ilvl: 0,
              level: { numFmt: 'decimal', lvlText: '%1)', indentLeft: 1440, hanging: 360 },
            },
          ],
        },
      },
    })
  expect(() => edit('99')).toThrow('Unknown list definition')
  edit('pending-7')
  expect(new Map(session.presentation().numbering).get('7')?.levels[0]?.lvlText).toBe('%1)')
  const zip = await JSZip.loadAsync(await session.save())
  expect(await zip.file('word/numbering.xml')!.async('string')).toContain('w:val="%1)"')
})
