import { WorkbookError, type DocumentDescriptor, type WorkbookState } from './domain'

export function assertFreshRevision(
  current: DocumentDescriptor,
  sessionId: string,
  expectedRevision?: number,
) {
  if (current.lifecycle === 'disposed' || current.sessionId !== sessionId)
    throw new WorkbookError(
      'session_closed',
      'This session has closed or been replaced; bind the current document again.',
    )
  if (expectedRevision !== undefined && expectedRevision !== current.revision)
    throw new WorkbookError(
      'revision_conflict',
      'The workbook changed; read the current revision before editing.',
      { expectedRevision, currentRevision: current.revision },
      true,
    )
}

/** Captures one immutable authority root and its revision before any awaits. */
export function captureReadView(
  descriptor: DocumentDescriptor,
  state: Readonly<WorkbookState>,
  expectedRevision?: number,
) {
  assertFreshRevision(descriptor, descriptor.sessionId, expectedRevision)
  if (descriptor.stateId !== state.stateId)
    throw new WorkbookError(
      'invalid_document',
      'The read descriptor does not match its state root.',
    )
  return Object.freeze({ descriptor: Object.freeze({ ...descriptor }), state })
}
