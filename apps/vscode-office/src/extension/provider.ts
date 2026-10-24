import * as vscode from 'vscode'
import { randomBytes } from 'node:crypto'
import { readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { WorkbookError } from '@genoffice/workbook-session/domain'
import { XlsxSidecarClient } from '@genoffice/workbook-session/sidecar/client'
import { SidecarWorkbookSource } from '@genoffice/workbook-session/sidecar/source'
import { WorkbookDocument } from './workbook-document'
import { saveWorkbookDocument } from './save'
import { backupWorkbookDocument, readWorkbookBackup } from './backup'
import { WorkbookBindings } from '@genoffice/workbook-session/bindings'
import { DocumentFileConflicts } from './file-conflicts'
import { WorkbookRegistry } from '@genoffice/workbook-session/registry'
import { captureFileBaseline } from '@genoffice/workbook-session/publish'
import { queryOperationStatus } from '@genoffice/workbook-session/receipts'
import { listRecoverySlots } from '@genoffice/workbook-session/recovery-slot'
import { WorkbookSnapshots } from '@genoffice/workbook-session/exchange/snapshots'
import { WorkbookImports } from '@genoffice/workbook-session/exchange/import-plan'
import { checkNative } from '../../scripts/native-resource.mjs'
import { workbookHtml } from './webview-html'
import { workbookAppearance } from '@genoffice/sheets-editor/appearance'
export const EDITOR_ID = 'genoffice.workbook'

export class WorkbookEditorProvider
  implements vscode.CustomEditorProvider<WorkbookDocument>, vscode.Disposable
{
  readonly bindings = new WorkbookBindings()
  readonly snapshots = new WorkbookSnapshots()
  readonly imports = new WorkbookImports(this.snapshots)
  private sweepTimer: ReturnType<typeof setInterval>
  get exchangeStorage() {
    return (this.context.storageUri ?? this.context.globalStorageUri).fsPath
  }
  private registry = new WorkbookRegistry()
  private documents = new Map<string, WorkbookDocument>()
  private opening = new Map<string, Promise<WorkbookDocument>>()
  private requestedRecovery = new Map<string, string>()
  private recoveredWithoutNativeContext = new WeakSet<WorkbookDocument>()
  private panels = new Map<WorkbookDocument, vscode.WebviewPanel>()
  private panelDetach = new Map<vscode.WebviewPanel, () => void>()
  private panelSubscriptions = new Map<vscode.WebviewPanel, vscode.Disposable>()
  private changes = new vscode.EventEmitter<
    | vscode.CustomDocumentEditEvent<WorkbookDocument>
    | vscode.CustomDocumentContentChangeEvent<WorkbookDocument>
  >()
  readonly onDidChangeCustomDocument = this.changes.event
  private client: XlsxSidecarClient
  private source: SidecarWorkbookSource
  constructor(
    private readonly context: vscode.ExtensionContext,
    private readonly output: vscode.OutputChannel,
  ) {
    const target = `${process.platform}-${process.arch}`
    if (!['linux-x64', 'win32-x64', 'darwin-arm64'].includes(target))
      throw new WorkbookError('unsupported_platform', `Unsupported extension host: ${target}`)
    this.client = new XlsxSidecarClient(
      join(
        context.extensionPath,
        'native',
        target,
        process.platform === 'win32' ? 'xlsx-sidecar.exe' : 'xlsx-sidecar',
      ),
    )
    this.source = new SidecarWorkbookSource(this.client)
    this.sweepTimer = setInterval(() => {
      void this.imports.sweep().catch((error) => this.output.appendLine(String(error)))
    }, 60_000)
    this.sweepTimer.unref()
  }
  getDocument(uri: vscode.Uri) {
    return this.documents.get(uri.toString())
  }
  openDocuments() {
    return [...this.documents.values()]
  }
  async recoverWorkbook(uri?: vscode.Uri) {
    if (!vscode.workspace.isTrusted)
      throw new WorkbookError('untrusted_workspace', 'Trust this workspace before recovery.')
    const storage = this.context.storageUri ?? this.context.globalStorageUri
    const slots = await listRecoverySlots(storage.fsPath)
    const unique = [
      ...new Map([...slots].reverse().map((slot) => [slot.documentUri, slot])).values(),
    ]
    const selected = uri
      ? unique.find((slot) => slot.documentUri === uri.toString())
      : (
          await vscode.window.showQuickPick(
            unique.map((slot) => ({
              label: vscode.Uri.parse(slot.documentUri).fsPath,
              description: new Date(slot.modifiedAt).toISOString(),
              slot,
            })),
            { title: 'GenOffice: Recover Unsaved Workbook' },
          )
        )?.slot
    if (!selected) {
      if (uri)
        throw new WorkbookError(
          'recovery_invalid',
          'No acknowledged recovery copy exists for this workbook.',
        )
      return
    }
    const target = vscode.Uri.parse(selected.documentUri)
    if (this.getDocument(target) || this.opening.has(target.toString()))
      throw new WorkbookError(
        'editor_busy',
        'Close the open workbook before choosing its recovery copy.',
      )
    this.requestedRecovery.set(target.toString(), vscode.Uri.file(selected.path).toString())
    try {
      await vscode.commands.executeCommand('vscode.openWith', target, EDITOR_ID)
    } finally {
      this.requestedRecovery.delete(target.toString())
    }
  }
  operationStatus(sessionId: string, transactionId: string) {
    return queryOperationStatus(
      this.registry,
      (uri) => {
        const document = this.documents.get(uri)
        return document
          ? {
              sessionId: document.sessionId,
              operationStatus: document.engine.operationStatus.bind(document.engine),
            }
          : undefined
      },
      sessionId,
      transactionId,
    )
  }
  async openCustomDocument(
    uri: vscode.Uri,
    openContext: vscode.CustomDocumentOpenContext,
  ): Promise<WorkbookDocument> {
    if (!vscode.workspace.isTrusted)
      throw new WorkbookError(
        'untrusted_workspace',
        'Trust the workspace before opening an editable workbook.',
      )
    if (uri.scheme !== 'file')
      throw new WorkbookError('unsupported_resource', 'Open a workbook file on the extension host.')
    const existing = this.documents.get(uri.toString())
    if (existing) return existing
    const pending = this.opening.get(uri.toString())
    if (pending) return pending
    const operation = (async () => {
      await checkNative(this.context.extensionPath, `${process.platform}-${process.arch}`)
      const lease = this.registry.reserve(uri.toString(), await captureFileBaseline(uri.fsPath))
      let incomplete: WorkbookDocument | undefined
      try {
        const storage = this.context.storageUri ?? this.context.globalStorageUri
        const backupId = openContext.backupId ?? this.requestedRecovery.get(uri.toString())
        const recovery = backupId
          ? await readWorkbookBackup(backupId, uri, [storage.fsPath])
          : undefined
        const document = await WorkbookDocument.open(
          uri,
          storage.fsPath,
          this.source,
          (error) => this.output.appendLine(String(error)),
          recovery,
        )
        incomplete = document
        document.backupId = backupId
        if (recovery && !openContext.backupId && document.engine.descriptor.dirty)
          this.recoveredWithoutNativeContext.add(document)
        document.attachOwnership(lease)
        this.documents.set(uri.toString(), document)
        document.watchFiles(
          new DocumentFileConflicts(document, (error) => this.output.appendLine(String(error))),
        )
        this.bindings.observe(document.engine.descriptor, document.engine.state)
        const stateChange = document.onDidStateChange(() => {
          this.bindings.observe(document.engine.descriptor, document.engine.state)
        })
        const change = document.onDidChange((event) => this.changes.fire(event))
        document.onDidDispose(() => {
          void this.imports
            .invalidateSession(document.sessionId)
            .catch((error) => this.output.appendLine(String(error)))
          this.bindings.invalidateSession(document.sessionId)
          change.dispose()
          stateChange.dispose()
          this.documents.delete(uri.toString())
          this.panels.delete(document)
        })
        return document
      } catch (error) {
        lease.release()
        incomplete?.dispose()
        throw error
      }
    })()
    this.opening.set(uri.toString(), operation)
    try {
      return await operation
    } finally {
      this.opening.delete(uri.toString())
    }
  }
  async resolveCustomEditor(document: WorkbookDocument, panel: vscode.WebviewPanel) {
    this.panels.set(document, panel)
    const root = vscode.Uri.joinPath(this.context.extensionUri, 'dist', 'webview')
    panel.webview.options = { enableScripts: true, localResourceRoots: [root] }
    this.panelDetach.get(panel)?.()
    this.panelDetach.set(panel, document.attach(panel.webview))
    const configuration = () =>
      vscode.workspace.getConfiguration('genoffice.workbook', document.uri)
    document.bridge!.handle('appearance.change', async (raw) => {
      const value = (raw as { appearance?: unknown })?.appearance
      if (value !== 'excel' && value !== 'vscode') throw new Error('Invalid workbook appearance.')
      await configuration().update(
        'appearance',
        value,
        vscode.workspace.getWorkspaceFolder(document.uri)
          ? vscode.ConfigurationTarget.WorkspaceFolder
          : vscode.ConfigurationTarget.Global,
      )
      return { appearance: workbookAppearance(configuration().get('appearance')) }
    })
    const appearanceSubscription = vscode.workspace.onDidChangeConfiguration((event) => {
      if (
        event.affectsConfiguration('genoffice.workbook.appearance', document.uri) &&
        !document.isClosed
      )
        void document.bridge
          ?.request('appearance.changed', {
            appearance: workbookAppearance(configuration().get('appearance')),
          })
          .catch((error) => this.output.appendLine(String(error)))
    })
    const previousSubscription = this.panelSubscriptions.get(panel)
    if (previousSubscription) appearanceSubscription.dispose()
    if (!this.panelSubscriptions.has(panel))
      this.panelSubscriptions.set(
        panel,
        vscode.Disposable.from(
          appearanceSubscription,
          panel.onDidDispose(() => {
            this.panelDetach.get(panel)?.()
            this.panelDetach.delete(panel)
            this.panelSubscriptions.delete(panel)
            this.panels.delete(document)
          }),
        ),
      )
    panel.webview.html = await this.html(document, panel.webview, root)
    if (this.recoveredWithoutNativeContext.delete(document)) this.changes.fire({ document })
  }
  private async html(document: WorkbookDocument, webview: vscode.Webview, root: vscode.Uri) {
    const nonce = randomBytes(24).toString('base64')
    const html = await readFile(join(root.fsPath, 'index.html'), 'utf8')
    return workbookHtml({
      html,
      root: root.fsPath,
      nonce,
      cspSource: webview.cspSource,
      language: vscode.env.language,
      sessionId: document.sessionId,
      appearance: workbookAppearance(
        vscode.workspace.getConfiguration('genoffice.workbook', document.uri).get('appearance'),
      ),
      resource: (path) => webview.asWebviewUri(vscode.Uri.file(path)).toString(),
    })
  }
  async saveCustomDocument(document: WorkbookDocument, token: vscode.CancellationToken) {
    await saveWorkbookDocument(document, document.uri, this.client, token)
    this.bindings.observe(document.engine.descriptor, document.engine.state)
  }
  saveCustomDocumentAs(
    document: WorkbookDocument,
    destination: vscode.Uri,
    token: vscode.CancellationToken,
  ) {
    return (async () => {
      const lease = this.registry.reserve(
        destination.toString(),
        await captureFileBaseline(destination.fsPath),
      )
      try {
        await saveWorkbookDocument(document, destination, this.client, token, true)
      } finally {
        lease.release()
      }
    })()
  }
  async revertCustomDocument(document: WorkbookDocument) {
    const previousSessionId = document.sessionId
    await this.imports.invalidateSession(previousSessionId)
    document.reserveReloadBaseline(await captureFileBaseline(document.uri.fsPath))
    await document.revert()
    this.bindings.invalidateSession(previousSessionId)
    this.bindings.observe(document.engine.descriptor, document.engine.state)
    document.watchFiles(
      new DocumentFileConflicts(document, (error) => this.output.appendLine(String(error))),
    )
    const panel = this.panels.get(document)
    if (panel) await this.resolveCustomEditor(document, panel)
  }
  async backupCustomDocument(
    document: WorkbookDocument,
    context: vscode.CustomDocumentBackupContext,
    token: vscode.CancellationToken,
  ) {
    return backupWorkbookDocument(document, context.destination, token)
  }
  async saveCopy(uri?: vscode.Uri) {
    if (!uri) {
      const selected = await vscode.window.showQuickPick(
        [...this.documents.values()].map((document) => ({
          label: document.uri.path.split('/').pop() ?? document.uri.path,
          description: document.uri.fsPath,
          uri: document.uri,
        })),
        { placeHolder: 'Choose the workbook to save as a new file' },
      )
      if (!selected) return
      uri = selected.uri
    }
    const document = this.getDocument(uri)
    if (!document) throw new WorkbookError('session_closed', 'The workbook is no longer open.')
    const sessionId = document.sessionId
    const destination = await vscode.window.showSaveDialog({
      defaultUri: vscode.Uri.file(document.uri.fsPath.replace(/\.xlsx$/i, '-copy.xlsx')),
      filters: { Workbook: ['xlsx'] },
    })
    if (!destination) return
    if (document !== this.getDocument(uri) || document.sessionId !== sessionId)
      throw new WorkbookError(
        'session_closed',
        'The workbook session changed while choosing a destination.',
      )
    const cancellation = new vscode.CancellationTokenSource()
    try {
      await this.saveCustomDocumentAs(document, destination, cancellation.token)
      await vscode.commands.executeCommand('vscode.openWith', destination, EDITOR_ID)
    } finally {
      cancellation.dispose()
    }
  }
  dispose() {
    clearInterval(this.sweepTimer)
    void this.imports.dispose().catch((error) => this.output.appendLine(String(error)))
    for (const detach of this.panelDetach.values()) detach()
    this.panelDetach.clear()
    for (const subscription of this.panelSubscriptions.values()) subscription.dispose()
    this.panelSubscriptions.clear()
    for (const document of this.documents.values()) document.dispose()
    this.documents.clear()
    this.changes.dispose()
    this.client.stop()
  }
}
