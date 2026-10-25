import { Transform } from '@tiptap/pm/transform'
import type { Node } from '@tiptap/pm/model'
import { operationsSchema } from '@genoffice/docs-editor/shared/operations'
export { operationsSchema, type DocxOperation } from '@genoffice/docs-editor/shared/operations'

export function operationSteps(
  doc: Node,
  input: unknown,
  context: {
    heading(level: number): string
    list(kind: 'bullet' | 'ordered'): string
  },
) {
  const operations = operationsSchema.parse(input)
  const tr = new Transform(doc)
  for (const op of operations) {
    let found: { node: Node; pos: number } | undefined
    tr.doc.descendants((node, pos) => {
      if (node.attrs.officeId === op.blockId) found = { node, pos }
    })
    if (!found) throw new Error('Unknown block identity')
    const { node, pos } = found
    if (
      node.attrs.officeLocked ||
      !['docParagraph', 'docHeading', 'docListItem'].includes(node.type.name)
    )
      throw new Error('Target is not an editable paragraph')
    if (op.type === 'formatParagraph')
      tr.setNodeMarkup(pos, undefined, { ...node.attrs, align: op.align })
    else if (op.type === 'insertParagraph') {
      tr.insert(
        pos + node.nodeSize,
        doc.type.schema.nodes.docParagraph!.create(
          null,
          op.text ? doc.type.schema.text(op.text) : undefined,
        ),
      )
    } else if (op.type === 'deleteParagraph') tr.delete(pos, pos + node.nodeSize)
    else if (op.type === 'formatText') {
      if (op.end <= op.start || op.end > node.content.size)
        throw new Error('Invalid formatting range')
      const from = pos + 1 + op.start,
        to = pos + 1 + op.end
      for (const name of ['bold', 'italic', 'underline', 'strike'] as const) {
        if (op.format[name] === true) tr.addMark(from, to, doc.type.schema.marks[name]!.create())
        else if (op.format[name] === false) tr.removeMark(from, to, doc.type.schema.marks[name])
      }
      const { font, color, sizeHalfPoints } = op.format
      const attrs = {
        ...(font ? { font } : {}),
        ...(color ? { color } : {}),
        ...(sizeHalfPoints ? { sizeHalfPoints } : {}),
      }
      if (Object.keys(attrs).length)
        tr.doc.nodesBetween(from, to, (text, offset) => {
          if (!text.isText) return
          const style = text.marks.find((mark) => mark.type.name === 'docTextStyle')
          tr.addMark(
            Math.max(from, offset),
            Math.min(to, offset + text.nodeSize),
            doc.type.schema.marks.docTextStyle!.create({ ...style?.attrs, ...attrs }),
          )
        })
    } else if (op.type === 'setHeading') {
      tr.setNodeMarkup(pos, doc.type.schema.nodes[op.level ? 'docHeading' : 'docParagraph'], {
        ...node.attrs,
        level: op.level || null,
        styleId: op.level ? context.heading(op.level) : null,
        numId: null,
        ilvl: null,
        kind: null,
      })
    } else if (op.type === 'setList') {
      tr.setNodeMarkup(pos, doc.type.schema.nodes.docListItem, {
        ...node.attrs,
        numId: context.list(op.kind),
        kind: op.kind,
        ilvl: 0,
      })
    } else {
      // These offsets are UTF-16 text offsets in a paragraph containing text only.
      // Rich inline objects require a separate typed operation, never guessed positions.
      let plain = true
      node.forEach((child) => {
        if (!child.isText) plain = false
      })
      if (!plain || op.end < op.start || op.end > node.content.size)
        throw new Error('Invalid text range')
      const from = pos + 1 + op.start,
        to = pos + 1 + op.end
      if (op.text)
        tr.replaceWith(from, to, doc.type.schema.text(op.text, node.resolve(op.start).marks()))
      else tr.delete(from, to)
    }
  }
  return tr.steps.map((step) => step.toJSON())
}
