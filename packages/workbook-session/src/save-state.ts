import { LIMITS, WorkbookError, type HistoryEntry, type WorkbookState } from './domain'
import type { ResourceStore } from './ports'
import type { OperationRecord } from './receipts'
import { historyReferences } from './history-records'
import { appliedTimeline } from './history'

/** Stages all references before the single atomic authority-root publication. */
export async function stageDocumentRoot(input: {
  store: ResourceStore
  sessionId: string
  revision: number
  state: WorkbookState
  savedStateId: string
  timeline: readonly HistoryEntry[]
  cursor: number
  operations: Iterable<[string, OperationRecord]>
}) {
  appliedTimeline(input.timeline, input.cursor)
  const bytes = Buffer.from(
    JSON.stringify({
      schemaVersion: 1,
      sessionId: input.sessionId,
      revision: input.revision,
      state: input.state,
      savedStateId: input.savedStateId,
      timeline: input.timeline,
      cursor: input.cursor,
      operations: [...input.operations],
    }),
  )
  if (bytes.length > LIMITS.historyBytes)
    throw new WorkbookError(
      'undo_budget_exceeded',
      'The document root exceeds its persistence budget.',
    )
  const chunks: string[] = []
  for (let offset = 0; offset < bytes.length; offset += LIMITS.chunkBytes)
    chunks.push(await input.store.writeChunk(bytes.subarray(offset, offset + LIMITS.chunkBytes)))
  return {
    schemaVersion: 1 as const,
    references: [
      ...new Set([...chunks, ...(await historyReferences(input.store, input.timeline))]),
    ],
    payload: {
      kind: 'document-root',
      sessionId: input.sessionId,
      revision: input.revision,
      stateId: input.state.stateId,
      serializedBytes: bytes.length,
      chunks,
    },
  }
}
