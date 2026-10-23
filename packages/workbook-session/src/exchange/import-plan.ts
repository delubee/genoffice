import { randomUUID } from 'node:crypto'
import { parseAddress, formatAddress } from '@genoffice/xlsx-gateway/domain/cell-address'
import {
  LIMITS,
  WorkbookError,
  assertNotCancelled,
  type CommitReceipt,
  type DocumentDescriptor,
} from '../domain'
import { assertFreshRevision } from '../freshness'
import type { WorkbookTransactions } from '../transactions'
import { prepareResultCopy } from './prepare-import'
import { readPrivateBatch } from './ndjson'
import type { WorkbookSnapshots, ExchangeReadView } from './snapshots'
import { preparedImportSchema } from './schemas'

type Status = 'prepared' | 'applying' | 'completed' | 'failed' | 'cancelled' | 'stale'
interface Plan {
  id: string
  snapshotId: string
  bindingId: string
  sessionId: string
  initialExpectedRevision: number
  chainExpectedRevision: number
  expiresAt: number
  status: Status
  nextCursor: number
  targetSheetId?: string
  overwriteCellCount: number
  copy: Awaited<ReturnType<typeof prepareResultCopy>>
  receipts: Map<number, { transactionId: string; expectedRevision: number; receipt: CommitReceipt }>
  queue: Promise<unknown>
  cancellation: AbortController
  maintenanceError?: string
}

export class WorkbookImports {
  private plans = new Map<string, Plan>()
  constructor(private readonly snapshots: WorkbookSnapshots) {}
  private progress(plan: Plan) {
    const completedCount = Math.min(
      plan.nextCursor * LIMITS.importBatchCells,
      plan.copy.manifest.shape.cellCount,
    )
    return preparedImportSchema.parse({
      planId: plan.id,
      snapshotId: plan.snapshotId,
      bindingId: plan.bindingId,
      sessionId: plan.sessionId,
      initialExpectedRevision: plan.initialExpectedRevision,
      chainExpectedRevision: plan.chainExpectedRevision,
      expiresAt: plan.expiresAt,
      status: plan.status,
      nextCursor: plan.nextCursor,
      completedBatchCount: plan.nextCursor,
      completedCount,
      remainingCount: plan.copy.manifest.shape.cellCount - completedCount,
      batchCount: plan.copy.batches.length,
      target: plan.targetSheetId
        ? { sheetId: plan.targetSheetId, startCell: plan.copy.manifest.target.startCell }
        : plan.copy.manifest.target,
      shape: plan.copy.manifest.shape,
      mode: plan.copy.manifest.mode,
      resultHash: plan.copy.manifest.output.sha256,
      overwriteCellCount: plan.overwriteCellCount,
      maxBatchCells: LIMITS.importBatchCells,
      historyPreflight:
        'Every batch must pass the actual 64 MiB transaction and 256 MiB retained-history budgets.',
      maintenanceError: plan.maintenanceError ?? null,
    })
  }
  private get(planId: string) {
    const plan = this.plans.get(planId)
    if (!plan)
      throw new WorkbookError(
        'plan_expired',
        'This prepared import is unknown, closed or expired; prepare a new result.',
      )
    return plan
  }
  status(planId: string) {
    return this.progress(this.get(planId))
  }
  async prepare(input: {
    snapshotId: string
    bindingId: string
    resultManifestPath: string
    privateStorage: string
    view: ExchangeReadView
    expectedRevision: number
    currentDescriptor(): DocumentDescriptor
    signal?: AbortSignal
  }) {
    await this.sweep()
    if (this.plans.size >= 128)
      throw new WorkbookError(
        'plan_limit',
        'Too many retained import plans; close old workbooks or wait for expiry.',
      )
    assertFreshRevision(
      input.currentDescriptor(),
      input.view.descriptor.sessionId,
      input.expectedRevision,
    )
    const source = this.snapshots.get(input.snapshotId)
    const copy = await prepareResultCopy({
      source,
      resultManifestPath: input.resultManifestPath,
      privateStorage: input.privateStorage,
      state: input.view.state,
      sessionId: input.view.descriptor.sessionId,
      revision: input.expectedRevision,
      signal: input.signal,
    })
    try {
      let overwriteCellCount = 0
      const target = copy.manifest.target,
        start = parseAddress(target.startCell),
        shape = copy.manifest.shape
      if ('sheetId' in target) {
        for (let row = 0; row < shape.rows;) {
          const rowEnd = Math.min(
            shape.rows - 1,
            row + Math.max(1, Math.floor(LIMITS.readCells / shape.columns)) - 1,
          )
          for (let column = 0; column < shape.columns; column += LIMITS.readCells) {
            const endColumn = Math.min(shape.columns - 1, column + LIMITS.readCells - 1)
            const page = await input.view.read(
              target.sheetId,
              `${formatAddress(start.row + row, start.column + column)}:${formatAddress(start.row + rowEnd, start.column + endColumn)}`,
              input.signal,
            )
            overwriteCellCount += page.cells.filter(
              (cell) => cell.kind !== 'blank' || !!cell.formula,
            ).length
          }
          row = rowEnd + 1
        }
      }
      assertNotCancelled(input.signal)
      assertFreshRevision(
        input.currentDescriptor(),
        input.view.descriptor.sessionId,
        input.expectedRevision,
      )
      const plan: Plan = {
        id: randomUUID(),
        snapshotId: input.snapshotId,
        bindingId: input.bindingId,
        sessionId: input.view.descriptor.sessionId,
        initialExpectedRevision: input.expectedRevision,
        chainExpectedRevision: input.expectedRevision,
        expiresAt: Date.now() + LIMITS.importLifetimeMs,
        status: 'prepared',
        nextCursor: 0,
        ...('sheetId' in target ? { targetSheetId: target.sheetId } : {}),
        overwriteCellCount,
        copy,
        receipts: new Map(),
        queue: Promise.resolve(),
        cancellation: new AbortController(),
      }
      this.plans.set(plan.id, plan)
      return this.progress(plan)
    } catch (error) {
      await copy.dispose()
      throw error
    }
  }
  apply(input: {
    planId: string
    bindingId: string
    cursor: number
    expectedRevision: number
    transactionId: string
    engine: WorkbookTransactions
    signal?: AbortSignal
  }) {
    const plan = this.get(input.planId),
      result = plan.queue.then(() => this.applyNext(plan, input))
    plan.queue = result.catch(() => {})
    return result
  }
  private async cleanCopy(plan: Plan) {
    await plan.copy.dispose().catch((error) => {
      plan.maintenanceError = String(error)
    })
  }
  private async applyNext(
    plan: Plan,
    input: {
      bindingId: string
      cursor: number
      expectedRevision: number
      transactionId: string
      engine: WorkbookTransactions
      signal?: AbortSignal
    },
  ) {
    const { engine } = input
    if (
      plan.bindingId !== input.bindingId ||
      plan.sessionId !== engine.descriptor.sessionId ||
      engine.descriptor.lifecycle === 'disposed'
    )
      throw new WorkbookError(
        'session_closed',
        'This plan belongs to a different or closed bound session.',
      )
    const prior = plan.receipts.get(input.cursor)
    if (prior) {
      if (prior.transactionId !== input.transactionId)
        throw new WorkbookError(
          'cursor_invalid',
          'This batch token was already consumed by another transaction.',
        )
      if (prior.expectedRevision !== input.expectedRevision)
        throw new WorkbookError(
          'idempotency_conflict',
          'The retried batch must retain its original expectedRevision.',
        )
      return { ...this.progress(plan), receipt: structuredClone(prior.receipt) }
    }
    if (Date.now() >= plan.expiresAt) {
      plan.status = 'failed'
      await this.cleanCopy(plan)
      throw new WorkbookError('plan_expired', 'Prepared imports expire after 30 minutes.', {
        progress: this.progress(plan),
      })
    }
    if (plan.status !== 'prepared' && plan.status !== 'applying')
      throw new WorkbookError(
        plan.status === 'stale' ? 'snapshot_stale' : 'cursor_invalid',
        'This plan cannot continue; export and prepare a new result.',
        { progress: this.progress(plan) },
      )
    if (
      !Number.isSafeInteger(input.cursor) ||
      input.cursor !== plan.nextCursor ||
      input.cursor >= plan.copy.batches.length
    )
      throw new WorkbookError(
        'cursor_invalid',
        'Apply the exact nextCursor once; do not skip or replay batch tokens.',
      )
    if (engine.descriptor.revision !== plan.chainExpectedRevision) {
      plan.status = 'stale'
      await this.cleanCopy(plan)
      throw new WorkbookError(
        'snapshot_stale',
        'Another edit or history move interrupted this import; export and recalculate.',
        { progress: this.progress(plan) },
        true,
      )
    }
    if (input.expectedRevision !== plan.chainExpectedRevision)
      throw new WorkbookError(
        'revision_conflict',
        'Use this plan’s chainExpectedRevision.',
        { progress: this.progress(plan) },
        true,
      )
    const signal = AbortSignal.any([
      plan.cancellation.signal,
      ...(input.signal ? [input.signal] : []),
    ])
    try {
      assertNotCancelled(signal)
      const count = Math.min(
        LIMITS.importBatchCells,
        plan.copy.manifest.shape.cellCount - input.cursor * LIMITS.importBatchCells,
      )
      const cells = await readPrivateBatch({
        path: plan.copy.dataPath,
        ...plan.copy.batches[input.cursor]!,
        count,
        signal,
      })
      const start = parseAddress(plan.copy.manifest.target.startCell),
        width = plan.copy.manifest.shape.columns
      const batch = {
        planId: plan.id,
        resultHash: plan.copy.manifest.output.sha256,
        mode: plan.copy.manifest.mode,
        target: plan.targetSheetId
          ? { sheetId: plan.targetSheetId }
          : {
              newSheetName:
                'newSheetName' in plan.copy.manifest.target
                  ? plan.copy.manifest.target.newSheetName
                  : '',
            },
        cells: cells.map((cell, offset) => {
          const index = input.cursor * LIMITS.importBatchCells + offset
          return {
            address: formatAddress(
              start.row + Math.floor(index / width),
              start.column + (index % width),
            ),
            cell,
          }
        }),
      }
      const receipt = await engine.submitImport(
        {
          schemaVersion: 1,
          sessionId: plan.sessionId,
          transactionId: input.transactionId,
          expectedRevision: input.expectedRevision,
          bindingId: input.bindingId,
          summary: `Import ${plan.id} batch ${input.cursor + 1}/${plan.copy.batches.length}`,
        },
        batch,
        signal,
      )
      // No cancellation/error may erase the committed fact after this point.
      plan.receipts.set(input.cursor, {
        transactionId: input.transactionId,
        expectedRevision: input.expectedRevision,
        receipt,
      })
      plan.nextCursor++
      plan.chainExpectedRevision = receipt.revision
      if (!plan.targetSheetId) plan.targetSheetId = receipt.affectedTargets[0]!.split('!')[0]!
      plan.status = signal.aborted
        ? 'cancelled'
        : plan.nextCursor === plan.copy.batches.length
          ? 'completed'
          : 'applying'
      if (plan.status === 'completed' || plan.status === 'cancelled') await this.cleanCopy(plan)
      return { ...this.progress(plan), receipt }
    } catch (error) {
      const failure =
        error instanceof WorkbookError
          ? error
          : new WorkbookError(
              'import_failed',
              error instanceof Error ? error.message : String(error),
            )
      plan.status =
        failure.code === 'operation_cancelled'
          ? 'cancelled'
          : failure.code === 'revision_conflict'
            ? 'stale'
            : 'failed'
      await this.cleanCopy(plan)
      throw new WorkbookError(
        failure.code,
        failure.message,
        { ...failure.details, progress: this.progress(plan) },
        failure.retryable,
      )
    }
  }
  async sweep() {
    for (const [id, plan] of this.plans)
      if (Date.now() >= plan.expiresAt) {
        plan.cancellation.abort()
        await plan.queue
        await this.cleanCopy(plan)
        this.plans.delete(id)
      }
  }
  async invalidateSession(sessionId: string) {
    for (const [id, plan] of this.plans)
      if (plan.sessionId === sessionId) {
        plan.cancellation.abort()
        plan.status = 'cancelled'
        await plan.queue
        await this.cleanCopy(plan)
        this.plans.delete(id)
      }
    this.snapshots.invalidateSession(sessionId)
  }
  async dispose() {
    for (const plan of this.plans.values()) plan.cancellation.abort()
    await Promise.all(
      [...this.plans.values()].map(async (plan) => {
        await plan.queue
        await this.cleanCopy(plan)
      }),
    )
    this.plans.clear()
  }
}
