import { Extension, getSchema, type AnyExtension, type NodeViewProps } from '@tiptap/core'
import { DOMSerializer } from '@tiptap/pm/model'
import { editorExtensions } from '../../../../apps/docs/src/renderer/editor/extensions'

export const blockTypes = ['docParagraph', 'docHeading', 'docListItem', 'docTable', 'docProtected']
const identity = Extension.create({
  name: 'officeIdentity',
  addGlobalAttributes() {
    return [
      {
        types: blockTypes,
        attributes: {
          officeId: { default: null, keepOnSplit: false, rendered: false },
          officeLocked: { default: false, keepOnSplit: false, rendered: false },
        },
      },
    ]
  },
})
const excluded = new Set([
  'undoRedo',
  'trackChanges',
  'tableHandle',
  'multiSelect',
  'aiQueueAnchors',
  'paginationGaps',
  'pageGapNav',
])
export const documentExtensions: AnyExtension[] = [
  ...(editorExtensions as AnyExtension[])
    .filter((extension) => !excluded.has(extension.name))
    .map((extension) =>
      extension.name === 'docProtected'
        ? extension.extend({
            draggable: false,
            addNodeView() {
              return ({ node }: NodeViewProps) => {
                const spec = node.type.spec.toDOM!(node)
                const dom = DOMSerializer.renderSpec(document, spec).dom as HTMLElement
                dom.contentEditable = 'false'
                dom.setAttribute('data-office-protected', 'true')
                return { dom, stopEvent: () => true }
              }
            },
          })
        : extension,
    ),
  identity,
]
export const docxSchema = getSchema(documentExtensions)
