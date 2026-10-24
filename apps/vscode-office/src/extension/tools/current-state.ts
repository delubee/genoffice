import { WorkbookError, assertNotCancelled } from '@genoffice/workbook-session/domain'
import type { WorkbookDocument } from '../workbook-document'

/** Flush the existing editor's pending user input, then read its host authority. */
export async function flushCurrentEditor(document: WorkbookDocument, signal?: AbortSignal) {
  assertNotCancelled(signal)
  const sessionId = document.sessionId,
    bridge = document.bridge
  if (document.isClosed) throw new WorkbookError('session_closed', 'The workbook is closed.')
  if (bridge) {
    let result: { ready?: boolean }
    try {
      result = (await bridge.request('view.flush', {})) as { ready?: boolean }
    } catch (error) {
      throw new WorkbookError(
        'editor_busy',
        'The current editor could not flush pending input; finish the edit or restore its view.',
        { reason: error instanceof Error ? error.message : String(error) },
        true,
      )
    }
    if (!result?.ready)
      throw new WorkbookError(
        'editor_busy',
        'Finish or cancel the current cell edit before continuing.',
        {},
        true,
      )
  }
  await document.engine.whenIdle()
  assertNotCancelled(signal)
  if (document.isClosed || document.sessionId !== sessionId)
    throw new WorkbookError('session_closed', 'The document closed or reverted while flushing.')
}
