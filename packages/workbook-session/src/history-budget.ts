import { LIMITS, WorkbookError } from './domain'

function lowerLimit(requested: number | undefined, maximum: number) {
  if (requested === undefined) return maximum
  if (!Number.isSafeInteger(requested) || requested < 1)
    throw new WorkbookError('undo_budget_exceeded', 'The configured history limit is invalid.')
  return Math.min(requested, maximum)
}

/** All origins preflight the actual serialized forward and inverse bytes. */
export function assertHistoryBudget(input: {
  forwardBytes: number
  inverseBytes: number
  retainedBytes: number
  retainedAdditionBytes?: number
  transactionByteLimit?: number
  historyByteLimit?: number
}) {
  const transactionLimit = lowerLimit(input.transactionByteLimit, LIMITS.transactionBytes),
    historyLimit = lowerLimit(input.historyByteLimit, LIMITS.historyBytes),
    serializedBytes = input.forwardBytes + input.inverseBytes
  if (
    ![input.forwardBytes, input.inverseBytes, input.retainedBytes, serializedBytes].every(
      (value) => Number.isSafeInteger(value) && value >= 0,
    )
  )
    throw new WorkbookError('undo_budget_exceeded', 'History byte counts are invalid.')
  if (serializedBytes > transactionLimit)
    throw new WorkbookError(
      'undo_budget_exceeded',
      'The bidirectional changeset exceeds the undo budget; shrink this transaction.',
      { scope: 'transaction', serializedBytes, limitBytes: transactionLimit },
    )
  const addition = input.retainedAdditionBytes ?? serializedBytes
  if (!Number.isSafeInteger(addition) || addition < 0)
    throw new WorkbookError('undo_budget_exceeded', 'Invalid retained byte increment.')
  if (input.retainedBytes + addition > historyLimit)
    throw new WorkbookError(
      'undo_budget_exceeded',
      'Retained undo/redo history exceeds its budget; saving will not discard it.',
      {
        scope: 'retained',
        retainedBytes: input.retainedBytes,
        serializedBytes,
        retainedAdditionBytes: addition,
        limitBytes: historyLimit,
      },
    )
  return serializedBytes
}
