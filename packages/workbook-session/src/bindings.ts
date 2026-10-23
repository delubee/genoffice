import { randomUUID } from 'node:crypto'
import { parseRange } from '@genoffice/xlsx-gateway/domain/cell-address'
import { MAX_GRID_COLUMNS, MAX_GRID_ROWS } from '@genoffice/xlsx-gateway/shared/grid-bounds'
import {
  WorkbookError,
  type DocumentDescriptor,
  type TaskBinding,
  type WorkbookState,
} from './domain'
import { bindingSchema } from './schemas'
import { assertFreshRevision } from './freshness'

interface SessionView {
  descriptor: DocumentDescriptor
  state: Readonly<WorkbookState>
  structure: string
}
interface BoundTask {
  binding: TaskBinding
  readRevision: number
}

/** One instance belongs to one extension window; bindings never follow focus. */
export class WorkbookBindings {
  private sessions = new Map<string, SessionView>()
  private tasks = new Map<string, BoundTask>()
  observe(descriptor: DocumentDescriptor, state: Readonly<WorkbookState>) {
    const previous = this.sessions.get(descriptor.sessionId)
    if (
      state.stateId !== descriptor.stateId ||
      (previous && previous.descriptor.documentUri !== descriptor.documentUri)
    )
      throw new WorkbookError(
        'invalid_document',
        'The observed document identity does not match its session.',
      )
    if (previous && descriptor.revision < previous.descriptor.revision)
      throw new WorkbookError(
        'revision_conflict',
        'Cannot replace a session with an older revision.',
      )
    if (
      previous &&
      descriptor.revision === previous.descriptor.revision &&
      state.stateId !== previous.state.stateId
    )
      throw new WorkbookError(
        'revision_conflict',
        'A content root cannot change without a new revision.',
      )
    const structure = JSON.stringify([
      state.structuralLog,
      state.sheets.map((sheet) => [sheet.id, sheet.name]),
    ])
    this.sessions.set(descriptor.sessionId, { descriptor: { ...descriptor }, state, structure })
    if (!previous || previous.descriptor.revision === descriptor.revision) return
    for (const task of this.tasks.values()) {
      if (task.binding.sessionId !== descriptor.sessionId || task.binding.status === 'invalidated')
        continue
      task.binding.status = 'stale'
      task.binding.selectionStatus =
        previous.structure !== structure ||
        task.binding.selectionStatus === 'selection_stale_structure'
          ? 'selection_stale_structure'
          : 'stale'
    }
  }
  bind(sessionId: string, logicalSheetId: string, frozenRange: string | null = null): TaskBinding {
    const view = this.sessions.get(sessionId)
    if (!view) throw new WorkbookError('session_closed', 'The requested session is closed.')
    assertFreshRevision(view.descriptor, sessionId)
    if (!view.state.sheets.some((sheet) => sheet.id === logicalSheetId))
      throw new WorkbookError('invalid_document', 'Choose an existing logical worksheet ID.')
    if (frozenRange !== null) {
      const bounds = parseRange(frozenRange)
      if (bounds.endRow >= MAX_GRID_ROWS || bounds.endColumn >= MAX_GRID_COLUMNS)
        throw new WorkbookError('invalid_document', 'The frozen range exceeds the Excel grid.')
    }
    const binding = bindingSchema.parse({
      bindingId: randomUUID(),
      sessionId,
      documentUri: view.descriptor.documentUri,
      logicalSheetId,
      frozenRange,
      boundRevision: view.descriptor.revision,
      createdAt: Date.now(),
      status: 'active',
      selectionStatus: 'current',
    })
    this.tasks.set(binding.bindingId, { binding, readRevision: view.descriptor.revision })
    return structuredClone(binding)
  }
  private task(bindingId: string) {
    const task = this.tasks.get(bindingId)
    if (!task)
      throw new WorkbookError(
        'invalid_document',
        'This binding is unknown in the current window; bind the document explicitly.',
      )
    return task
  }
  get(bindingId: string): TaskBinding {
    return structuredClone(this.task(bindingId).binding)
  }
  private active(bindingId: string) {
    const task = this.task(bindingId),
      view = this.sessions.get(task.binding.sessionId)
    if (task.binding.status === 'invalidated' || !view)
      throw new WorkbookError(
        'session_closed',
        'The bound session is closed; create a new binding.',
      )
    return { task, view }
  }
  acknowledgeRead(bindingId: string, revision: number) {
    const { task, view } = this.active(bindingId)
    assertFreshRevision(view.descriptor, task.binding.sessionId, revision)
    task.readRevision = revision
    task.binding.status = 'active'
    if (task.binding.selectionStatus !== 'selection_stale_structure')
      task.binding.selectionStatus = 'current'
  }
  authorizeWrite(
    bindingId: string,
    expectedRevision: number,
    usesFrozenSelection = true,
  ): TaskBinding {
    const { task, view } = this.active(bindingId)
    assertFreshRevision(view.descriptor, task.binding.sessionId, expectedRevision)
    if (task.readRevision !== expectedRevision)
      throw new WorkbookError(
        'revision_conflict',
        'Read the current document before applying this stale binding.',
        {},
        true,
      )
    if (usesFrozenSelection && task.binding.selectionStatus === 'selection_stale_structure')
      throw new WorkbookError(
        'selection_stale_structure',
        'The original selection changed structurally; bind the selection again.',
      )
    return structuredClone(task.binding)
  }
  invalidateSession(sessionId: string) {
    this.sessions.delete(sessionId)
    for (const task of this.tasks.values())
      if (task.binding.sessionId === sessionId) task.binding.status = 'invalidated'
  }
}
