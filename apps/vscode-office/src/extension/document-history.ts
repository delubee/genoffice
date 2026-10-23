import * as vscode from 'vscode'
import { WorkbookError, type CommitReceipt } from '@genoffice/workbook-session/domain'
import type { WorkbookTransactions } from '@genoffice/workbook-session/transactions'
import type { WorkbookDocument } from './workbook-document'

let requestedDocument: string | undefined
export async function requestNativeHistory(document: WorkbookDocument, direction: 'undo' | 'redo') {
  const input = vscode.window.tabGroups.activeTabGroup.activeTab?.input
  if (document.isClosed) throw new WorkbookError('session_closed', 'The workbook has closed.')
  if (!(input instanceof vscode.TabInputCustom) || input.uri.toString() !== document.uri.toString())
    throw new WorkbookError(
      'view_unavailable',
      'Activate this workbook before using its native undo/redo.',
    )
  if (requestedDocument)
    throw new WorkbookError('editor_busy', 'Wait for the current native history operation.')
  requestedDocument = document.uri.toString()
  try {
    await vscode.commands.executeCommand(direction)
  } finally {
    requestedDocument = undefined
  }
}

/** Exactly one native undo element per committed canonical transaction. */
export function nativeHistoryEdit(
  document: WorkbookDocument,
  engine: WorkbookTransactions,
  receipt: CommitReceipt,
): vscode.CustomDocumentEditEvent<WorkbookDocument> {
  const entry = engine.history[engine.historyPosition.cursor - 1]
  if (!entry || entry.afterStateId !== receipt.stateId)
    throw new WorkbookError(
      'invalid_changeset',
      'The native edit has no matching durable history entry.',
    )
  return {
    document,
    label: receipt.summary,
    async undo() {
      if (requestedDocument && requestedDocument !== document.uri.toString())
        throw new WorkbookError(
          'revision_conflict',
          'Editor focus changed before the requested undo; no workbook change was applied.',
        )
      await engine.undo(entry.entryId)
    },
    async redo() {
      if (requestedDocument && requestedDocument !== document.uri.toString())
        throw new WorkbookError(
          'revision_conflict',
          'Editor focus changed before the requested redo; no workbook change was applied.',
        )
      await engine.redo(entry.entryId)
    },
  }
}
