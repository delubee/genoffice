import type { z } from 'zod'
import type * as schemas from './schemas'

export const LIMITS = Object.freeze({
  readCells: 2_000,
  formatCells: 200,
  aggregateCells: 1_000_000,
  expandedWrites: 2_000,
  rangeCells: 200_000,
  importBatchCells: 1_000,
  artifactCells: 1_000_000,
  artifactBytes: 64 * 1024 * 1024,
  manifestBytes: 256 * 1024,
  recordBytes: 256 * 1024,
  messageBytes: 1024 * 1024,
  chunkBytes: 256 * 1024,
  transactionBytes: 64 * 1024 * 1024,
  historyBytes: 256 * 1024 * 1024,
  importLifetimeMs: 30 * 60 * 1000,
})

export class WorkbookError extends Error {
  constructor(
    readonly code: string,
    message: string,
    readonly details: Record<string, unknown> = {},
    readonly retryable = false,
  ) {
    super(message)
    this.name = 'WorkbookError'
  }
}

export type CellRecord = z.infer<typeof schemas.cellRecordSchema>
export type SheetRecord = z.infer<typeof schemas.sheetRecordSchema>
export type WorkbookState = z.infer<typeof schemas.workbookStateSchema>
export type TransactionInput = z.infer<typeof schemas.transactionInputSchema>
export type CommitReceipt = z.infer<typeof schemas.receiptSchema>
export type TaskBinding = z.infer<typeof schemas.bindingSchema>
export type ProtocolMessage = z.infer<typeof schemas.protocolMessageSchema>

export interface OriginalBase {
  readonly path: string
  readonly sha256: string
  readonly sourceUri: string
  readonly sidecarSessionId: string
}

export interface HistoryEntry {
  readonly entryId: string
  readonly beforeStateId: string
  readonly afterStateId: string
  readonly origin: TransactionInput['origin']
  readonly label: string
  readonly forwardRef: string
  readonly inverseRef: string
  readonly serializedBytes: number
}

export interface PreparedChangeset {
  readonly beforeStateId: string
  readonly nextState: WorkbookState
  readonly forward: unknown
  readonly inverse: unknown
  readonly affectedTargets: readonly string[]
  readonly calculationState: 'ready' | 'pending' | 'error'
}

export interface DocumentDescriptor {
  readonly documentUri: string
  readonly sessionId: string
  readonly revision: number
  readonly stateId: string
  readonly savedStateId: string
  readonly dirty: boolean
  readonly lifecycle:
    | 'opening'
    | 'ready'
    | 'preparing'
    | 'committing'
    | 'saving'
    | 'conflict'
    | 'recovering'
    | 'disposed'
}

export function assertNotCancelled(signal?: AbortSignal) {
  if (signal?.aborted)
    throw new WorkbookError('operation_cancelled', 'The operation was cancelled before commit.')
}
