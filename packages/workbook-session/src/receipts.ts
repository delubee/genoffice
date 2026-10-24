import { createHash } from 'node:crypto'
import { z } from 'zod'
import { WorkbookError, type CommitReceipt, type TransactionInput } from './domain'
import { receiptSchema } from './schemas'
import type { WorkbookRegistry } from './registry'

export const operationRecordSchema = z
  .strictObject({
    status: z.enum(['received', 'preparing', 'prepared', 'committed', 'rejected', 'cancelled']),
    payloadHash: z.string().regex(/^[a-f0-9]{64}$/),
    transactionId: z.string().min(1).max(256),
    receipt: receiptSchema.optional(),
    error: z.strictObject({ code: z.string(), message: z.string() }).optional(),
  })
  .superRefine((record, context) => {
    if (
      (record.status === 'committed') !== !!record.receipt ||
      (record.receipt && record.receipt.transactionId !== record.transactionId)
    )
      context.addIssue({
        code: 'custom',
        message: 'Committed operation records require their matching receipt.',
      })
    if ((record.status === 'rejected' || record.status === 'cancelled') && !record.error)
      context.addIssue({ code: 'custom', message: 'Rejected operations require a terminal error.' })
  })
export type OperationRecord = z.infer<typeof operationRecordSchema>
export const intentFingerprint = (intent: TransactionInput) =>
  createHash('sha256').update(JSON.stringify(intent)).digest('hex')

export function resolveIdempotent(previous: OperationRecord, payloadHash: string): CommitReceipt {
  if (previous.payloadHash !== payloadHash)
    throw new WorkbookError(
      'idempotency_conflict',
      'The transaction ID was already used with a different payload.',
    )
  if (previous.receipt) return structuredClone(previous.receipt)
  if (previous.error) throw new WorkbookError(previous.error.code, previous.error.message)
  throw new WorkbookError(
    'operation_pending',
    'This transaction is still being prepared.',
    {},
    true,
  )
}

/** Status queries never open a file or migrate receipts into a new session. */
export async function queryOperationStatus(
  registry: WorkbookRegistry,
  lookup: (
    documentUri: string,
  ) => { sessionId: string; operationStatus(transactionId: string): Promise<unknown> } | undefined,
  sessionId: string,
  transactionId: string,
) {
  const owner = registry.session(sessionId),
    session = owner && lookup(owner.documentUri)
  if (!session || session.sessionId !== sessionId)
    throw new WorkbookError(
      'session_closed',
      'The requested operation belongs to a closed session.',
    )
  return session.operationStatus(transactionId)
}
