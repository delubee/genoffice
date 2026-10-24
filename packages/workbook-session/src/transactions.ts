import { createHash } from 'node:crypto'
import {
  WorkbookError,
  assertNotCancelled,
  type CommitReceipt,
  type DocumentDescriptor,
  type HistoryEntry,
  type WorkbookState,
} from './domain'
import type { CandidateInterpreter, HostLifecycle, ResourceStore } from './ports'
import { transactionInputSchema, workbookStateSchema } from './schemas'
import { immutable } from './state'
import { prepareHistoryRecord } from './history-records'
import { appliedTimeline, prepareHistoryMove } from './history'
import { stageDocumentRoot } from './save-state'
import { assertFreshRevision } from './freshness'
import { importBatchSchema, type ImportBatch } from './exchange/schemas'
import { importValueOperations } from './exchange/import-candidate'
import {
  intentFingerprint,
  operationRecordSchema,
  resolveIdempotent,
  type OperationRecord,
} from './receipts'

interface Options {
  transactionByteLimit?: number
  historyByteLimit?: number
  sessionId: string
  documentUri: string
  state: WorkbookState
  savedStateId?: string
  store: ResourceStore
  interpreter: CandidateInterpreter
  host: HostLifecycle
}
const hash = (value: string) => createHash('sha256').update(value).digest('hex')

/** One authority and queue per CustomDocument; commit point is root publication. */
export class WorkbookTransactions {
  private current: WorkbookState
  private revision = 0
  private savedStateId: string
  private lifecycle: DocumentDescriptor['lifecycle'] = 'ready'
  private timeline: HistoryEntry[] = []
  private cursor = 0
  private operations = new Map<string, OperationRecord>()
  private queue: Promise<unknown> = Promise.resolve()
  private closed = false
  private externalConflict: string | undefined
  projectionError: unknown
  maintenanceError: unknown
  constructor(private readonly options: Options) {
    this.current = immutable(workbookStateSchema.parse(options.state))
    this.savedStateId = options.savedStateId ?? this.current.stateId
  }
  get state(): WorkbookState {
    return this.current
  }
  get descriptor(): DocumentDescriptor {
    return {
      documentUri: this.options.documentUri,
      sessionId: this.options.sessionId,
      revision: this.revision,
      stateId: this.current.stateId,
      savedStateId: this.savedStateId,
      dirty: this.current.stateId !== this.savedStateId,
      lifecycle:
        this.lifecycle !== 'disposed' && this.externalConflict ? 'conflict' : this.lifecycle,
    }
  }
  get history(): readonly HistoryEntry[] {
    return structuredClone(this.timeline)
  }
  get historyPosition() {
    return { cursor: this.cursor, length: this.timeline.length }
  }
  noteExternalConflict(reason: string) {
    if (this.closed) return
    this.externalConflict = reason
    this.options.host.conflict(this.options.documentUri, reason)
  }
  submit(raw: unknown, signal?: AbortSignal): Promise<CommitReceipt> {
    const result = this.queue.then(() => this.execute(raw, signal))
    this.queue = result.catch(() => {})
    return result
  }
  submitImport(
    raw: Record<string, unknown>,
    batchInput: unknown,
    signal?: AbortSignal,
  ): Promise<CommitReceipt> {
    const batch = immutable(importBatchSchema.parse(batchInput))
    const operations =
      'newSheetName' in batch.target
        ? [{ op: 'add_sheet', name: batch.target.newSheetName }]
        : importValueOperations(batch, batch.target.sheetId)
    const intent = { ...raw, origin: 'import', operations }
    const result = this.queue.then(() => this.execute(intent, signal, batch))
    this.queue = result.catch(() => {})
    return result
  }
  private assertActive() {
    if (this.closed) throw new WorkbookError('session_closed', 'The workbook session has closed.')
  }
  private async record(operation: OperationRecord) {
    this.operations.set(operation.transactionId, operation)
    await this.options.store.publishManifest(`operation-${hash(operation.transactionId)}`, {
      schemaVersion: 1,
      references: [],
      payload: JSON.parse(JSON.stringify(operation)),
    })
  }
  private async execute(
    raw: unknown,
    signal?: AbortSignal,
    batch?: ImportBatch,
  ): Promise<CommitReceipt> {
    this.assertActive()
    const intent = transactionInputSchema.parse(raw)
    assertFreshRevision(this.descriptor, intent.sessionId)
    const payloadHash = batch
        ? hash(`${intentFingerprint(intent)}:${JSON.stringify(batch)}`)
        : intentFingerprint(intent),
      previous = this.operations.get(intent.transactionId)
    // Retried committed requests carry their original revision. Check identity first.
    if (previous) return resolveIdempotent(previous, payloadHash)
    const operation: OperationRecord = {
      status: 'received',
      transactionId: intent.transactionId,
      payloadHash,
    }
    try {
      assertNotCancelled(signal)
      assertFreshRevision(this.descriptor, intent.sessionId, intent.expectedRevision)
      await this.record(operation)
      this.lifecycle = 'preparing'
      await this.record({ ...operation, status: 'preparing' })
      const before = this.current,
        previousRevision = this.revision
      const prepared = batch
        ? await (this.options.interpreter.prepareImport?.(before, intent, batch, signal) ??
            Promise.reject(
              new WorkbookError('unsupported_operation', 'This host cannot prepare imports.'),
            ))
        : await this.options.interpreter.prepare(before, intent, signal)
      this.assertActive()
      assertNotCancelled(signal)
      if (
        this.current !== before ||
        this.revision !== previousRevision ||
        prepared.beforeStateId !== before.stateId
      )
        throw new WorkbookError(
          'revision_conflict',
          'The candidate was prepared against an obsolete root.',
        )
      const nextState = immutable(workbookStateSchema.parse(prepared.nextState))
      if (nextState.stateId === before.stateId)
        throw new WorkbookError('invalid_changeset', 'A new commit must have a new state ID.')
      const entry = await prepareHistoryRecord({
        store: this.options.store,
        before,
        prepared: { ...prepared, nextState },
        origin: intent.origin,
        label: intent.summary,
        retainedEntries: appliedTimeline(this.timeline, this.cursor),
        transactionByteLimit: this.options.transactionByteLimit,
        historyByteLimit: this.options.historyByteLimit,
      })
      await this.record({ ...operation, status: 'prepared' })
      this.assertActive()
      assertNotCancelled(signal)
      if (!Number.isSafeInteger(previousRevision + 1))
        throw new WorkbookError(
          'revision_exhausted',
          'Reopen this workbook to start a new revision epoch.',
        )
      const receipt: CommitReceipt = {
        sessionId: this.options.sessionId,
        transactionId: intent.transactionId,
        status: 'committed',
        previousRevision,
        revision: previousRevision + 1,
        stateId: nextState.stateId,
        affectedTargets: [...prepared.affectedTargets],
        summary: intent.summary,
        calculationState: prepared.calculationState,
      }
      const timeline = [...appliedTimeline(this.timeline, this.cursor), entry],
        operations = new Map(this.operations)
      operations.set(intent.transactionId, { ...operation, status: 'committed', receipt })
      const manifest = await stageDocumentRoot({
        store: this.options.store,
        sessionId: this.options.sessionId,
        revision: receipt.revision,
        state: nextState,
        savedStateId: this.savedStateId,
        timeline,
        cursor: timeline.length,
        operations,
      })
      this.assertActive()
      assertNotCancelled(signal)
      this.lifecycle = 'committing'
      await this.options.store.publishManifest('current', manifest)
      // No cancellation check or fallible content mutation after this point.
      this.current = nextState
      this.revision = receipt.revision
      this.timeline = timeline
      this.cursor = timeline.length
      this.operations = operations
      this.lifecycle = this.closed ? 'disposed' : 'ready'
      await this.notify(receipt)
      await this.collect()
      return structuredClone(receipt)
    } catch (error) {
      const failure =
        error instanceof WorkbookError
          ? error
          : new WorkbookError(
              'transaction_rejected',
              error instanceof Error ? error.message : String(error),
            )
      const terminal: OperationRecord = {
        ...operation,
        status: failure.code === 'operation_cancelled' ? 'cancelled' : 'rejected',
        error: { code: failure.code, message: failure.message },
      }
      // Failure to persist a diagnostic must not hide the original failure.
      try {
        await this.record(terminal)
      } catch {
        this.operations.set(intent.transactionId, terminal)
      }
      this.lifecycle = this.closed ? 'disposed' : 'ready'
      await this.collect()
      throw failure
    }
  }
  async operationStatus(transactionId: string) {
    const operation = this.operations.get(transactionId)
    if (operation) return structuredClone(operation)
    try {
      const stored = operationRecordSchema.parse(
        (
          (await this.options.store.readManifest(`operation-${hash(transactionId)}`)) as {
            payload: unknown
          }
        ).payload,
      )
      if (stored.transactionId !== transactionId)
        throw new WorkbookError(
          'invalid_document',
          'The operation record has an unexpected transaction ID.',
        )
      return stored
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error
    }
    return { status: 'unknown', transactionId }
  }
  private async notify(receipt?: CommitReceipt) {
    this.projectionError = undefined
    try {
      if (receipt) await this.options.host.committed(receipt, this.descriptor)
      else await this.options.host.historyChanged?.(this.descriptor)
    } catch (error) {
      this.projectionError = error
    }
    try {
      await this.options.host.projectionChanged(this.current.stateId, this.revision)
    } catch (error) {
      this.projectionError = error
    }
  }
  private async collect() {
    if (!this.options.store.collect) return
    try {
      let retained = ['current']
      try {
        await (this.options.store.readManifestMetadata?.('current') ??
          this.options.store.readManifest('current'))
      } catch (error) {
        if ((error as NodeJS.ErrnoException).code === 'ENOENT') retained = []
        else throw error
      }
      if (retained.length)
        retained.push(...this.timeline.flatMap((entry) => [entry.forwardRef, entry.inverseRef]))
      await this.options.store.collect(retained)
      this.maintenanceError = undefined
    } catch (error) {
      this.maintenanceError = error
    }
  }
  undo(entryId?: string, signal?: AbortSignal) {
    return this.move('undo', entryId, signal)
  }
  redo(entryId?: string, signal?: AbortSignal) {
    return this.move('redo', entryId, signal)
  }
  private move(
    direction: 'undo' | 'redo',
    entryId?: string,
    signal?: AbortSignal,
  ): Promise<DocumentDescriptor> {
    const task = this.queue.then(async () => {
      this.assertActive()
      assertNotCancelled(signal)
      try {
        const before = this.current,
          revision = this.revision,
          prepared = await prepareHistoryMove({
            store: this.options.store,
            entries: this.timeline,
            cursor: this.cursor,
            current: before,
            direction,
            ...(entryId ? { entryId } : {}),
          })
        if (!Number.isSafeInteger(revision + 1))
          throw new WorkbookError(
            'revision_exhausted',
            'Reopen this workbook to start a new revision epoch.',
          )
        const manifest = await stageDocumentRoot({
          store: this.options.store,
          sessionId: this.options.sessionId,
          revision: revision + 1,
          state: prepared.state,
          savedStateId: this.savedStateId,
          timeline: this.timeline,
          cursor: prepared.cursor,
          operations: this.operations,
        })
        this.assertActive()
        assertNotCancelled(signal)
        if (this.current !== before || this.revision !== revision)
          throw new WorkbookError(
            'revision_conflict',
            'The workbook changed while preparing history.',
          )
        this.lifecycle = 'committing'
        await this.options.store.publishManifest('current', manifest)
        this.current = prepared.state
        this.revision = revision + 1
        this.cursor = prepared.cursor
        this.lifecycle = this.closed ? 'disposed' : 'ready'
        await this.notify()
        await this.collect()
        return this.descriptor
      } catch (error) {
        this.lifecycle = this.closed ? 'disposed' : 'ready'
        await this.collect()
        throw error
      }
    })
    this.queue = task.catch(() => {})
    return task
  }
  saveWith(write: (state: Readonly<WorkbookState>) => Promise<void>, markSaved = true) {
    const task = this.queue.then(async () => {
      this.assertActive()
      if (markSaved && this.externalConflict)
        throw new WorkbookError('external_change_conflict', this.externalConflict, {}, true)
      this.lifecycle = 'saving'
      const state = this.current
      try {
        const manifest = markSaved
          ? await stageDocumentRoot({
              store: this.options.store,
              sessionId: this.options.sessionId,
              revision: this.revision,
              state,
              savedStateId: state.stateId,
              timeline: this.timeline,
              cursor: this.cursor,
              operations: this.operations,
            })
          : undefined
        this.assertActive()
        await write(state)
        if (manifest) {
          try {
            await this.options.store.publishManifest('current', manifest)
          } catch (error) {
            throw new WorkbookError(
              'save_failed',
              'The file was written, but its saved marker could not be persisted; content and history are retained.',
              { fileWritten: true, reason: error instanceof Error ? error.message : String(error) },
            )
          }
          this.savedStateId = state.stateId
          await this.collect()
        }
      } finally {
        this.lifecycle = this.closed ? 'disposed' : 'ready'
        await this.collect()
      }
    })
    this.queue = task.catch(() => {})
    return task
  }
  async whenIdle() {
    await this.queue
  }
  async dispose() {
    this.closed = true
    await this.queue
    await this.options.interpreter.dispose()
    this.lifecycle = 'disposed'
  }
}
