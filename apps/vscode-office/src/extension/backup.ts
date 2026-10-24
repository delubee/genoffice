import * as vscode from 'vscode'
import { dirname } from 'node:path'
import { createRecoveryPoint, readRecoveryPoint } from '@genoffice/workbook-session/recovery'
import { publishRecoverySlot, readRecoverySlot } from '@genoffice/workbook-session/recovery-slot'
import { WorkbookError } from '@genoffice/workbook-session/domain'
import type { WorkbookDocument } from './workbook-document'

export async function backupWorkbookDocument(
  document: WorkbookDocument,
  destination: vscode.Uri,
  token: vscode.CancellationToken,
): Promise<vscode.CustomDocumentBackup> {
  if (destination.scheme !== 'file')
    throw new WorkbookError('unsupported_resource', 'Backup storage must be on the extension host.')
  if (document.isClosed)
    throw new WorkbookError('session_closed', 'The workbook session has closed.')
  const controller = new AbortController(),
    subscription = token.onCancellationRequested(() => controller.abort())
  if (token.isCancellationRequested) controller.abort()
  let snapshot: ReturnType<WorkbookDocument['snapshot']> | undefined
  try {
    if (controller.signal.aborted)
      throw new WorkbookError('operation_cancelled', 'The backup was cancelled.')
    if (document.bridge) {
      const flush = (await document.bridge.request('view.flush', {})) as { ready?: boolean }
      if (!flush.ready)
        throw new WorkbookError(
          'editor_busy',
          'Finish the current edit before acknowledging a backup.',
        )
    }
    snapshot = document.snapshot()
    const point = await createRecoveryPoint({
      storageDirectory: dirname(destination.fsPath),
      original: snapshot.original,
      documentUri: document.uri.toString(),
      state: snapshot.state,
      savedStateId: snapshot.descriptor.savedStateId,
      baselineHash: snapshot.baselineHash,
      signal: controller.signal,
    })
    let slot
    try {
      slot = await publishRecoverySlot({
        storageDirectory: dirname(destination.fsPath),
        documentUri: document.uri.toString(),
        point,
        slotPath: document.backupId ? vscode.Uri.parse(document.backupId).fsPath : undefined,
      })
    } catch (error) {
      await point.dispose()
      throw error
    }
    const id = vscode.Uri.file(slot.path).toString()
    document.backupId = id
    return {
      id,
      delete() {
        void slot.dispose().catch(() => {})
      },
    }
  } finally {
    subscription.dispose()
    await snapshot?.release()
  }
}

export async function readWorkbookBackup(
  backupId: string,
  originalUri: vscode.Uri,
  allowedRoots: readonly string[],
) {
  const uri = vscode.Uri.parse(backupId),
    hash = new URLSearchParams(uri.query).get('sha256')
  if (uri.scheme !== 'file')
    throw new WorkbookError('recovery_invalid', 'Invalid acknowledged backup URI.')
  if (!hash)
    return readRecoverySlot({
      slotPath: uri.fsPath,
      documentUri: originalUri.toString(),
      allowedRoots,
    })
  return readRecoveryPoint({
    manifestPath: uri.fsPath,
    expectedManifestHash: hash,
    documentUri: originalUri.toString(),
    allowedRoots,
  })
}
