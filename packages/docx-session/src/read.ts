import type { EditorJSON } from '@genoffice/docs-editor/conversion'
import { jsonBytes } from '@genoffice/docs-editor/shared/policy'

export interface BlockView {
  blockId: string
  type: string
  protected: boolean
  text?: string
}
export interface SelectionScope {
  stateId: string
  revision: number
  from: number
  to: number
  blocks: Array<BlockView & { start?: number; end?: number; length?: number }>
}
export function readBlocks(doc: EditorJSON, offset = 0, limit = 100): BlockView[] {
  if (
    !Number.isInteger(offset) ||
    offset < 0 ||
    !Number.isInteger(limit) ||
    limit < 1 ||
    limit > 100
  )
    throw new Error('Invalid read range')
  const textOf = (node: EditorJSON): string =>
    node.text ?? (node.content ?? []).map(textOf).join(node.type === 'docTableRow' ? '\t' : '')
  const result = (doc.content ?? []).slice(offset, offset + limit).map((block): BlockView => {
    const protectedContent = Boolean(block.attrs?.officeLocked)
    return {
      blockId: String(block.attrs?.officeId),
      type: block.type,
      protected: protectedContent,
      ...(protectedContent ? {} : { text: textOf(block) }),
    }
  })
  if (jsonBytes(result) > 256 * 1024) throw new Error('Read exceeds 256 KiB; request fewer blocks')
  return result
}
