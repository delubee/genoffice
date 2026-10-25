import { it, expect } from 'vitest'
import { EditorState, TextSelection } from '@tiptap/pm/state'
import { Transform } from '@tiptap/pm/transform'
import { addRowAfter, addColumnAfter } from '@tiptap/pm/tables'
import { parseDocx } from '@genoffice/docx-engine'
import { docxSchema } from '@genoffice/docs-editor/shared/schema'
import { buildDocx } from '../../docx-engine/tests/helpers/build-docx'
import { DocxSession } from '../src/session'

it('creates a rectangular table and saves row/column edits as real OOXML tables', async () => {
  const session = await DocxSession.open(
    await buildDocx({ bodyXml: '<w:p><w:r><w:t>Before table</w:t></w:r></w:p>' }),
  )
  const nodes = docxSchema.nodes
  const cell = () =>
    nodes.docTableCell!.create(null, nodes.docParagraph!.create(null, docxSchema.text('Cell')))
  const row = () => nodes.docTableRow!.create(null, [cell(), cell()])
  const table = nodes.docTable!.create(null, [row(), row()])
  const submit = (tr: Transform) => {
    const descriptor = session.descriptor()
    session.submit({
      sessionId: session.sessionId,
      epoch: descriptor.epoch,
      expectedRevision: descriptor.revision,
      baseCollabVersion: descriptor.collabVersion,
      clientId: 'browser',
      transactionId: `table-${descriptor.revision}`,
      steps: tr.steps.map((step) => step.toJSON()),
    })
  }
  const doc = docxSchema.nodeFromJSON(session.content())
  submit(new Transform(doc).insert(doc.content.size, table))
  for (const command of [addRowAfter, addColumnAfter]) {
    const current = docxSchema.nodeFromJSON(session.content())
    let firstCell = -1
    current.descendants((node, pos) => {
      if (node.type.name === 'docTableCell' && firstCell < 0) firstCell = pos
    })
    const state = EditorState.create({
      schema: docxSchema,
      doc: current,
      selection: TextSelection.near(current.resolve(firstCell + 2)),
    })
    expect(command(state, submit)).toBe(true)
  }
  const reopened = await parseDocx(await session.save())
  const savedTable = reopened.blocks.find((block) => block.type === 'table')!.table!
  expect(savedTable.rows).toHaveLength(3)
  expect(savedTable.rows[0]!).toHaveLength(3)
})
