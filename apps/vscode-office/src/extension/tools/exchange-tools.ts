import * as vscode from 'vscode'
import type { z } from 'zod'
import { WorkbookError } from '@genoffice/workbook-session/domain'
import { assertFreshRevision } from '@genoffice/workbook-session/freshness'
import type { WorkbookEditorProvider } from '../provider'
import {
  chooseWorkspaceRoot,
  assertCurrentWorkspaceRoot,
  workspaceResultPath,
} from '../workspace-resources'
import { flushCurrentEditor } from './current-state'
import { toolSchemas } from './schemas'

export class WorkbookExchangeTools {
  constructor(private readonly provider: WorkbookEditorProvider) {}
  describeExport(bindingId: string, workspaceUri?: string) {
    return chooseWorkspaceRoot(this.target(bindingId).uri, workspaceUri).toString()
  }
  private target(bindingId: string) {
    const binding = this.provider.bindings.get(bindingId),
      document = this.provider.getDocument(vscode.Uri.parse(binding.documentUri))
    if (
      !document ||
      document.isClosed ||
      document.sessionId !== binding.sessionId ||
      binding.status === 'invalidated'
    )
      throw new WorkbookError(
        'session_closed',
        'The exchange binding belongs to a closed or replaced workbook.',
      )
    return document
  }
  async export(input: z.infer<typeof toolSchemas.export>, signal?: AbortSignal) {
    const document = this.target(input.bindingId)
    assertFreshRevision(document.engine.descriptor, document.sessionId, input.expectedRevision)
    const workspace = chooseWorkspaceRoot(document.uri, input.workspaceUri)
    await flushCurrentEditor(document, signal)
    assertFreshRevision(document.engine.descriptor, document.sessionId, input.expectedRevision)
    const view = document.snapshot()
    try {
      const result = await this.provider.snapshots.export({
        view,
        sheetId: input.sheetId,
        range: input.range,
        workspaceRoot: workspace.fsPath,
        signal,
      })
      if (document.isClosed || document.sessionId !== view.descriptor.sessionId) {
        this.provider.snapshots.invalidateSession(view.descriptor.sessionId)
        throw new WorkbookError(
          'session_closed',
          'The export completed after its source session closed; its workspace files are retained without import authority.',
        )
      }
      if (
        document.sessionId === view.descriptor.sessionId &&
        document.engine.descriptor.revision === view.descriptor.revision
      )
        this.provider.bindings.acknowledgeRead(input.bindingId, view.descriptor.revision)
      return {
        sessionId: view.descriptor.sessionId,
        revision: view.descriptor.revision,
        data: {
          snapshotId: result.manifest.snapshotId,
          sourceRevision: result.manifest.sourceRevision,
          sourceStateId: result.manifest.sourceStateId,
          manifestUri: vscode.Uri.file(result.manifestPath).toString(),
          manifestPath: result.manifestPath,
          dataUri: vscode.Uri.file(result.dataPath).toString(),
          dataPath: result.dataPath,
          taskDirectoryUri: vscode.Uri.file(result.directory).toString(),
          taskDirectory: result.directory,
          workspaceRoot: result.workspaceRoot,
          manifestSha256: result.manifestSha256,
          dataSha256: result.manifest.data.sha256,
          bytes: result.manifest.data.bytes,
          shape: result.manifest.shape,
          dateSystem: result.manifest.dateSystem,
          calculationState: result.manifest.calculationState,
        },
      }
    } finally {
      await view.release()
    }
  }
  async prepareImport(input: z.infer<typeof toolSchemas.prepareImport>, signal?: AbortSignal) {
    const document = this.target(input.bindingId)
    assertFreshRevision(document.engine.descriptor, document.sessionId, input.expectedRevision)
    await flushCurrentEditor(document, signal)
    this.provider.bindings.authorizeWrite(input.bindingId, input.expectedRevision, false)
    const source = this.provider.snapshots.get(input.snapshotId)
    await assertCurrentWorkspaceRoot(source.workspaceRoot)
    const view = document.snapshot()
    try {
      const result = await this.provider.imports.prepare({
        snapshotId: input.snapshotId,
        bindingId: input.bindingId,
        resultManifestPath: workspaceResultPath(input.resultManifestUri, source.workspaceRoot),
        privateStorage: this.provider.exchangeStorage,
        view,
        expectedRevision: input.expectedRevision,
        currentDescriptor: () => document.engine.descriptor,
        signal,
      })
      return {
        sessionId: view.descriptor.sessionId,
        revision: view.descriptor.revision,
        data: result,
      }
    } finally {
      await view.release()
    }
  }
  describePlan(bindingId: string, planId: string) {
    const document = this.target(bindingId),
      plan = this.provider.imports.status(planId)
    if (plan.sessionId !== document.sessionId || plan.bindingId !== bindingId)
      throw new WorkbookError(
        'source_mismatch',
        'This plan belongs to a different binding or session.',
      )
    return plan
  }
  async applyImport(input: z.infer<typeof toolSchemas.applyImport>, signal?: AbortSignal) {
    const document = this.target(input.bindingId),
      plan = this.describePlan(input.bindingId, input.planId)
    // Retry receipts remain facts even if the caller is cancelled or still has the old revision.
    if (input.cursor >= plan.nextCursor && !signal?.aborted)
      await flushCurrentEditor(document, signal)
    const result = await this.provider.imports.apply({ ...input, engine: document.engine, signal })
    return { sessionId: result.sessionId, revision: result.receipt.revision, data: result }
  }
  planStatus(sessionId: string, planId: string) {
    if (!this.provider.openDocuments().some((document) => document.sessionId === sessionId))
      throw new WorkbookError('session_closed', 'Status queries do not reopen a closed workbook.')
    const plan = this.provider.imports.status(planId)
    if (plan.sessionId !== sessionId)
      throw new WorkbookError('source_mismatch', 'The plan belongs to a different session.')
    return plan
  }
}
