import { randomUUID } from 'node:crypto'
import { goalSeekRequestSchema } from '@genoffice/xlsx-gateway/shared/session-goal-seek'
import {
  WorkbookError,
  assertNotCancelled,
  type WorkbookState,
  type DocumentDescriptor,
  type TransactionInput,
} from './domain'
import type { WorkbookSource } from './ports'
import { prepareCandidateMaterial } from './candidate-material'
import { assertMergeOperation } from '@genoffice/xlsx-gateway/domain/session-merges'
import { loadCalculationSource } from './calculation/source'
import { seekCalculationSource } from './calculation/goal-seek'

export class WorkbookGoalSeekController {
  private lifetime = new AbortController()
  private busy = false
  constructor(
    private readonly ports: {
      sessionId: string
      sourceSessionId: string
      source: WorkbookSource
      getState(): WorkbookState
      getDescriptor(): DocumentDescriptor
      submit(intent: TransactionInput): Promise<unknown>
    },
  ) {}
  async handle(raw: unknown) {
    const request = goalSeekRequestSchema.parse(raw)
    if (this.busy) throw new WorkbookError('operation_busy', 'Goal Seek is already running.')
    const signal = AbortSignal.any([this.lifetime.signal, AbortSignal.timeout(20_000)])
    const deadline = Date.now() + 20_000
    const current = () => {
      const descriptor = this.ports.getDescriptor()
      return (
        this.ports.getState().stateId === request.stateId &&
        descriptor.revision === request.revision &&
        descriptor.stateId === request.stateId
      )
    }
    const check = () => {
      assertNotCancelled(signal)
      if (Date.now() >= deadline)
        throw new WorkbookError('operation_cancelled', 'Goal Seek exceeded its time limit.')
      if (!current())
        throw new WorkbookError('revision_conflict', 'The document changed during Goal Seek.')
    }
    check()
    this.busy = true
    try {
      const state = structuredClone(this.ports.getState())
      const probe = {
        op: 'set_cell' as const,
        sheetId: request.sheetId,
        address: request.byCell,
        value: 0,
      }
      const material = await prepareCandidateMaterial({
        state,
        operations: [probe],
        source: this.ports.source,
        sourceSessionId: this.ports.sourceSessionId,
        signal,
      })
      check()
      const sheet = material.state.sheets.find((item) => item.id === request.sheetId)
      const metadata = state.features?.originalMetadata as { readOnly?: boolean } | undefined
      if (
        !sheet ||
        sheet.readOnly ||
        sheet.protected ||
        metadata?.readOnly ||
        !material.writableSheetIds.includes(request.sheetId)
      )
        throw new WorkbookError('protected_target', 'The changing cell is not writable.')
      assertMergeOperation(material.state, probe)
      const source = await loadCalculationSource({
        state,
        source: this.ports.source,
        sourceSessionId: this.ports.sourceSessionId,
        isCurrent: current,
        signal,
      })
      const result = await seekCalculationSource(source, request, { isCurrent: current, signal })
      check()
      if (result.found && sheet.cells[request.byCell]?.rawValue !== result.solution) {
        await this.ports.submit({
          schemaVersion: 1,
          sessionId: this.ports.sessionId,
          transactionId: randomUUID(),
          expectedRevision: request.revision,
          origin: 'manual',
          summary: 'Goal Seek',
          operations: [{ ...probe, value: result.solution }],
        })
      }
      return {
        stateId: this.ports.getState().stateId,
        revision: this.ports.getDescriptor().revision,
        found: result.found,
        solution: result.solution,
        reached: Number.isFinite(result.reached) ? result.reached : null,
        iterations: result.iterations,
      }
    } finally {
      this.busy = false
    }
  }
  dispose() {
    this.lifetime.abort()
  }
}
