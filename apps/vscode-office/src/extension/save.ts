import type * as vscode from 'vscode'
import { SessionPackageWriter } from '@genoffice/workbook-session/package-writer'
import { WorkbookError } from '@genoffice/workbook-session/domain'
import { captureFileBaseline } from '@genoffice/workbook-session/publish'
import type { XlsxSidecarClient } from '@genoffice/workbook-session/sidecar/client'
import type { WorkbookTransactions } from '@genoffice/workbook-session/transactions'
import type { OriginalBase } from '@genoffice/workbook-session/domain'
import type { FileBaseline } from '@genoffice/workbook-session/publish'
import type { BridgeMethod } from '@genoffice/xlsx-gateway/shared/session-bridge'
import { rm } from 'node:fs/promises'
import { dirname } from 'node:path'
export interface WorkbookSaveDocument {
  readonly engine: WorkbookTransactions
  readonly sessionId: string
  readonly isClosed: boolean
  readonly original: OriginalBase
  readonly metadata: unknown
  readonly targetPath: string
  baselineHash: string | null
  fileBaseline: FileBaseline
  readonly bridge: { request(method: BridgeMethod, payload: unknown): Promise<unknown> } | undefined
  pauseFileChecks(): () => void
}
export async function saveWorkbookDocument(
  document: WorkbookSaveDocument,
  destination: Pick<vscode.Uri, 'scheme' | 'path' | 'fsPath'>,
  client: XlsxSidecarClient,
  token: vscode.CancellationToken,
  saveAs = false,
) {
  if (token.isCancellationRequested)
    throw new WorkbookError('operation_cancelled', 'Save was cancelled before starting.')
  if (destination.scheme !== 'file')
    throw new WorkbookError('unsupported_resource', 'Save to a file on the extension host.')
  if (!destination.path.toLowerCase().endsWith('.xlsx'))
    throw new WorkbookError('unsupported_format', 'Save this workbook with the .xlsx extension.')
  if (
    !saveAs &&
    (document.engine.state.sheets.some((sheet) => sheet.readOnly) ||
      (document.fileBaseline.mode & 0o222) === 0)
  )
    throw new WorkbookError('save_failed', 'The workbook source is read-only; choose Save As.')
  const flush = (await document.bridge?.request('view.flush', {})) as
    { ready?: boolean } | undefined
  if (document.bridge && !flush?.ready)
    throw new WorkbookError('editor_busy', 'Finish the current cell edit before saving.')
  const controller = new AbortController(),
    subscription = token.onCancellationRequested(() => controller.abort())
  if (token.isCancellationRequested) controller.abort()
  const resumeChecks = document.pauseFileChecks()
  const sessionId = document.sessionId
  try {
    await document.engine.saveWith(async (state) => {
      const targetPath = saveAs ? destination.fsPath : document.targetPath
      const baseline = saveAs ? await captureFileBaseline(targetPath) : document.fileBaseline,
        expected = saveAs ? baseline.sha256 : document.baselineHash
      const writer = new SessionPackageWriter(client, document.metadata)
      const staged = await writer.stage(
        document.original.path,
        state,
        targetPath,
        controller.signal,
      )
      if (controller.signal.aborted || document.isClosed || document.sessionId !== sessionId) {
        await rm(dirname(staged.stagedPath), { recursive: true, force: true })
        throw new WorkbookError('operation_cancelled', 'Save was cancelled before publication.')
      }
      await writer.publish(staged.stagedPath, targetPath, expected, baseline)
      if (!saveAs) {
        document.fileBaseline = writer.publishedBaseline!
        document.baselineHash = staged.hash
      }
    }, !saveAs)
  } finally {
    resumeChecks()
    subscription.dispose()
  }
}
