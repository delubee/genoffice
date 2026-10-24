import * as vscode from 'vscode'
import { randomUUID } from 'node:crypto'
import { join } from 'node:path'
import { rm } from 'node:fs/promises'
import { WorkbookError, type OriginalBase } from '@genoffice/workbook-session/domain'
import { WorkbookTransactions } from '@genoffice/workbook-session/transactions'
import { FileResourceStore } from '@genoffice/workbook-session/store'
import { openWorkbookSnapshot } from '@genoffice/workbook-session/open'
import { initialState } from '@genoffice/workbook-session/state'
import type { WorkbookSource, CandidateInterpreter } from '@genoffice/workbook-session/ports'
import type { readWorkbookBackup } from './backup'
import { WebviewCandidateBridge } from './editor-bridge'
import { captureFileBaseline, type FileBaseline } from '@genoffice/workbook-session/publish'
import type { WorkbookLease } from '@genoffice/workbook-session/registry'
import { nativeHistoryEdit, requestNativeHistory } from './document-history'
import { readCurrentRange } from '@genoffice/workbook-session/read'
import { prepareImportCandidate } from '@genoffice/workbook-session/exchange/import-candidate'
type Recovery = Awaited<ReturnType<typeof readWorkbookBackup>>
interface Actor {
  sessionId: string
  original: OriginalBase
  metadata: unknown
  baselineHash: string | null
  fileBaseline: FileBaseline
  targetPath: string
  opened: Awaited<ReturnType<typeof openWorkbookSnapshot>>
  store: FileResourceStore
  engine: WorkbookTransactions
  view?: WebviewCandidateBridge
}

/** VS Code document identity survives revert; its canonical session does not. */
export class WorkbookDocument implements vscode.CustomDocument {
  backupId: string | undefined
  private actor!: Actor
  private changes = new vscode.EventEmitter<vscode.CustomDocumentEditEvent<WorkbookDocument>>()
  private disposedEvent = new vscode.EventEmitter<void>()
  readonly onDidChange = this.changes.event
  private stateChanges = new vscode.EventEmitter<void>()
  readonly onDidStateChange = this.stateChanges.event
  readonly onDidDispose = this.disposedEvent.event
  private closed = false
  private shutdown: Promise<void> | undefined
  private fileWatcher: { pause(): () => void; dispose(): void } | undefined
  private ownership: WorkbookLease | undefined
  private constructor(
    readonly uri: vscode.Uri,
    private readonly storageDirectory: string,
    private readonly source: WorkbookSource,
    private readonly log: (error: unknown) => void,
  ) {}
  get sessionId() {
    return this.actor.sessionId
  }
  get engine() {
    return this.actor.engine
  }
  get original() {
    return this.actor.original
  }
  get metadata() {
    return this.actor.metadata
  }
  get baselineHash() {
    return this.actor.baselineHash
  }
  set baselineHash(value: string | null) {
    this.actor.baselineHash = value
  }
  get targetPath() {
    return this.actor.targetPath
  }
  get fileBaseline() {
    return this.actor.fileBaseline
  }
  set fileBaseline(value: FileBaseline) {
    this.ownership?.update(value)
    this.actor.fileBaseline = value
  }
  attachOwnership(lease: WorkbookLease) {
    lease.update(this.fileBaseline)
    lease.bind(this.sessionId)
    this.ownership = lease
  }
  reserveReloadBaseline(baseline: FileBaseline) {
    this.ownership?.update(baseline)
  }
  get isClosed() {
    return this.closed
  }
  watchFiles(watcher: { pause(): () => void; dispose(): void }) {
    this.fileWatcher?.dispose()
    this.fileWatcher = watcher
  }
  pauseFileChecks() {
    return this.fileWatcher?.pause() ?? (() => {})
  }
  requestHistory(direction: 'undo' | 'redo') {
    return requestNativeHistory(this, direction)
  }
  get bridge() {
    return this.actor.view?.bridge
  }
  async whenViewReady() {
    const view = this.actor.view
    if (!view) throw new WorkbookError('view_unavailable', 'No workbook view is attached.')
    await view.ready
  }
  snapshot() {
    const actor = this.actor
    const state = actor.engine.state
    return {
      original: actor.original,
      state,
      descriptor: actor.engine.descriptor,
      baselineHash: actor.baselineHash,
      read: (sheetId: string, range: string, signal?: AbortSignal) =>
        readCurrentRange({
          state,
          sessionId: actor.original.sidecarSessionId,
          source: this.source,
          logicalSheetId: sheetId,
          range,
          signal,
        }),
      release: actor.opened.retain(),
    }
  }
  static async open(
    uri: vscode.Uri,
    storageDirectory: string,
    source: WorkbookSource,
    log: (error: unknown) => void,
    recovery?: Recovery,
  ) {
    if (uri.scheme !== 'file')
      throw new WorkbookError('unsupported_resource', 'Open a workbook file on the extension host.')
    if (!uri.path.toLowerCase().endsWith('.xlsx'))
      throw new WorkbookError('unsupported_format', 'This version supports .xlsx workbooks.')
    const document = new WorkbookDocument(uri, storageDirectory, source, log)
    document.actor = await document.createActor(recovery)
    return document
  }
  private async createActor(recovery?: Recovery): Promise<Actor> {
    const opened = await openWorkbookSnapshot({
      uri: recovery ? vscode.Uri.file(recovery.basePath).toString() : this.uri.toString(),
      storageDirectory: this.storageDirectory,
      source: this.source,
    })
    try {
      const fileBaseline = await captureFileBaseline(this.uri.fsPath)
      if (!recovery && fileBaseline.sha256 !== opened.original.sha256)
        throw new WorkbookError(
          'external_change_conflict',
          'The source changed while opening; reopen the workbook.',
        )
      const store = await FileResourceStore.create(
        join(this.storageDirectory, `session-${randomUUID()}`),
      )
      const candidate: CandidateInterpreter = {
        prepareImport: (state, intent, batch, signal) =>
          prepareImportCandidate({
            state,
            intent,
            batch,
            signal,
            source: this.source,
            sourceSessionId: opened.original.sidecarSessionId,
          }),
        prepare: (state, intent, signal) => {
          if (!actor.view)
            throw new WorkbookError('view_unavailable', 'Open the workbook view before editing.')
          return actor.view.prepare(state, intent, signal)
        },
        dispose: () => actor.view?.dispose(),
      }
      const sessionId = randomUUID()
      const engine = new WorkbookTransactions({
        documentUri: this.uri.toString(),
        sessionId,
        state: recovery?.state ?? initialState(opened.metadata),
        ...(recovery ? { savedStateId: recovery.savedStateId } : {}),
        store,
        interpreter: candidate,
        host: {
          committed: (receipt) => {
            this.stateChanges.fire()
            this.changes.fire(nativeHistoryEdit(this, actor.engine, receipt))
          },
          historyChanged: () => this.stateChanges.fire(),
          projectionChanged: async () => {
            if (actor.view)
              await actor.view.bridge.request('state.committed', {
                descriptor: actor.engine.descriptor,
                state: actor.engine.state,
              })
          },
          conflict: (_uri, reason) =>
            this.log(new WorkbookError('external_change_conflict', reason)),
        },
      })
      const actor: Actor = {
        sessionId,
        engine,
        opened,
        store,
        original: Object.freeze({ ...opened.original, sourceUri: this.uri.toString() }),
        metadata: opened.metadata,
        baselineHash: recovery ? recovery.baselineHash : opened.original.sha256,
        fileBaseline,
        targetPath: recovery ? this.uri.fsPath : opened.canonicalPath,
      }
      if (recovery && fileBaseline.sha256 !== recovery.baselineHash)
        engine.noteExternalConflict(
          'The source changed after this recovery copy was created; save to a new file or reload.',
        )
      return actor
    } catch (error) {
      await opened.dispose()
      throw error
    }
  }
  attach(webview: vscode.Webview) {
    if (this.closed) throw new WorkbookError('session_closed', 'The document is closed.')
    const actor = this.actor
    actor.view?.dispose()
    const view = new WebviewCandidateBridge({
      webview,
      sessionId: actor.sessionId,
      original: actor.original,
      source: this.source,
      getState: () => actor.engine.state,
      getDescriptor: () => actor.engine.descriptor,
      submit: (intent) => actor.engine.submit(intent),
      history: async (direction) => {
        if (this.closed || this.actor !== actor)
          throw new WorkbookError('session_closed', 'The history view belongs to a closed session.')
        await this.requestHistory(direction)
      },
    })
    actor.view = view
    return () => {
      if (actor.view === view) {
        view.dispose()
        actor.view = undefined
      }
    }
  }
  async revert() {
    if (this.closed) throw new WorkbookError('session_closed', 'The document is closed.')
    await this.actor.engine.whenIdle()
    const replacement = await this.createActor()
    if (this.closed) {
      await this.destroyActor(replacement)
      throw new WorkbookError('session_closed', 'The document closed while reloading.')
    }
    const previous = this.actor
    await this.destroyActor(previous)
    this.actor = replacement
    this.ownership?.update(replacement.fileBaseline)
    this.ownership?.bind(replacement.sessionId)
  }
  private async destroyActor(actor: Actor) {
    actor.view?.dispose()
    actor.view = undefined
    try {
      await actor.engine.dispose()
    } finally {
      await actor.opened.dispose()
      await rm(actor.store.directory, { recursive: true, force: true })
    }
  }
  dispose() {
    if (this.closed) return
    this.closed = true
    this.fileWatcher?.dispose()
    this.ownership?.release()
    this.disposedEvent.fire()
    this.disposedEvent.dispose()
    this.changes.dispose()
    this.stateChanges.dispose()
    this.shutdown = this.destroyActor(this.actor)
    void this.shutdown.catch(this.log)
  }
  async whenDisposed() {
    await this.shutdown
  }
}
