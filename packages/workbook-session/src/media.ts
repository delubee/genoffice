import { workbookMediaResultSchema } from '@genoffice/xlsx-gateway/shared/workbook-protocol'
import { z } from 'zod'
import type { WorkbookSource } from './ports'
import type { WorkbookState } from './domain'
import { WorkbookError } from './domain'

export const currentMediaRequestSchema = z.strictObject({
  stateId: z.string().min(1),
  visualId: z.string().min(1).max(300),
})
/** Media identities resolve only against the immutable source of the current root. */
export async function readCurrentMedia(input: {
  request: unknown
  getState(): Readonly<WorkbookState>
  source: WorkbookSource
  sourceSessionId: string
}) {
  const request = currentMediaRequestSchema.parse(input.request),
    state = input.getState()
  const current = () => input.getState().stateId === request.stateId
  if (!current())
    throw new WorkbookError('revision_conflict', 'The requested state is no longer current.')
  const metadata = state.features?.originalMetadata as
    { visuals?: { id: string; sheetId: string; kind: string }[] } | undefined
  const matches = metadata?.visuals?.filter((v) => v.id === request.visualId) ?? []
  const visual = matches[0],
    owner = state.sheets.find((s) => s.originalSheetId === visual?.sheetId)
  if (
    matches.length !== 1 ||
    !visual ||
    !owner ||
    !['image', 'shape'].includes(visual.kind) ||
    (owner.features?.visualEdits as Record<string, { remove?: true }> | undefined)?.[visual.id]
      ?.remove
  )
    throw new WorkbookError('unknown_visual', 'The media is not in the current workbook.')
  if (!input.source.readMedia)
    throw new WorkbookError('unsupported_operation', 'The workbook source cannot read media.')
  const result = workbookMediaResultSchema
    .extend({
      base64: z
        .string()
        .min(1)
        .max(48 * 1024 * 1024),
    })
    .parse(await input.source.readMedia(input.sourceSessionId, visual.id))
  if (!current())
    throw new WorkbookError('revision_conflict', 'The document changed while reading media.')
  return result
}
