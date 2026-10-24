import * as vscode from 'vscode'
import { basename, dirname } from 'node:path'
import { captureFileBaseline, sameFileBaseline } from '@genoffice/workbook-session/publish'
import type { WorkbookDocument } from './workbook-document'

/** Watches both the requested URI and canonical path, including alias changes. */
export class DocumentFileConflicts implements vscode.Disposable {
  private disposables: vscode.Disposable[] = []
  private timer: ReturnType<typeof setTimeout> | undefined
  private paused = 0
  private closed = false
  private warned = false
  constructor(
    private readonly document: WorkbookDocument,
    private readonly log: (error: unknown) => void,
  ) {
    for (const path of new Set([document.uri.fsPath, document.targetPath])) {
      const watcher = vscode.workspace.createFileSystemWatcher(
        new vscode.RelativePattern(vscode.Uri.file(dirname(path)), basename(path)),
      )
      this.disposables.push(
        watcher,
        watcher.onDidChange(() => this.schedule()),
        watcher.onDidCreate(() => this.schedule()),
        watcher.onDidDelete(() => this.schedule()),
      )
    }
  }
  private schedule() {
    if (this.closed || this.paused) return
    if (this.timer) clearTimeout(this.timer)
    this.timer = setTimeout(() => {
      void this.check().catch(this.log)
    }, 40)
  }
  async check() {
    if (this.closed || this.paused) return
    const sessionId = this.document.sessionId,
      baseline = this.document.fileBaseline
    let changed: boolean
    try {
      changed = !sameFileBaseline(await captureFileBaseline(this.document.uri.fsPath), baseline)
    } catch (error) {
      this.log(error)
      changed = true
    }
    if (
      this.closed ||
      this.paused ||
      this.document.sessionId !== sessionId ||
      this.document.fileBaseline !== baseline
    )
      return
    if (!changed) return
    const reason =
      'The workbook changed on disk. Current edits are preserved; use Save As or reload before replacing the original.'
    this.document.engine.noteExternalConflict(reason)
    if (!this.warned) {
      this.warned = true
      void Promise.resolve(
        vscode.window
          .showWarningMessage(`${this.document.uri.fsPath}: ${reason}`, 'Save As')
          .then((choice) => {
            if (choice === 'Save As' && !this.closed)
              return vscode.commands.executeCommand('genoffice.saveWorkbookAs', this.document.uri)
          }),
      ).catch(this.log)
    }
  }
  pause() {
    this.paused++
    let released = false
    return () => {
      if (released) return
      released = true
      this.paused--
      this.schedule()
    }
  }
  dispose() {
    this.closed = true
    if (this.timer) clearTimeout(this.timer)
    for (const disposable of this.disposables) disposable.dispose()
  }
}
