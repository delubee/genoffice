import { WorkbookError, type HistoryEntry, type WorkbookState } from './domain'
import type { ResourceStore } from './ports'
import { loadHistoryState } from './history-records'
import { immutable } from './state'

export function appliedTimeline(entries: readonly HistoryEntry[], cursor: number) {
  if (!Number.isSafeInteger(cursor) || cursor < 0 || cursor > entries.length)
    throw new WorkbookError('invalid_document', 'The history cursor is invalid.')
  return entries.slice(0, cursor)
}

export async function prepareHistoryMove(input: {
  store: ResourceStore
  entries: readonly HistoryEntry[]
  cursor: number
  current: Readonly<WorkbookState>
  direction: 'undo' | 'redo'
  entryId?: string
}) {
  appliedTimeline(input.entries, input.cursor)
  const index = input.direction === 'undo' ? input.cursor - 1 : input.cursor,
    entry = input.entries[index]
  if (!entry || (input.entryId !== undefined && entry.entryId !== input.entryId))
    throw new WorkbookError(
      'revision_conflict',
      'This history entry is no longer the next undo/redo operation.',
    )
  const expected = input.direction === 'undo' ? entry.afterStateId : entry.beforeStateId
  if (input.current.stateId !== expected)
    throw new WorkbookError(
      'revision_conflict',
      'The current root does not match this history entry.',
    )
  const state = immutable(
    await loadHistoryState(input.store, entry, input.direction === 'undo' ? 'inverse' : 'forward'),
  )
  return { entry, state, cursor: input.direction === 'undo' ? input.cursor - 1 : input.cursor + 1 }
}
