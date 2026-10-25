import * as vscode from 'vscode'
import type { FileBaseline } from '@genoffice/workbook-session/publish'
import type { DocxSession } from '@genoffice/docx-session/session'
import type { JsonBridge } from '@genoffice/xlsx-gateway/shared/session-bridge'

export class DocxDocument implements vscode.CustomDocument {
  bridge?: JsonBridge
  private ready?: Promise<void>
  private resolveReady?: () => void
  private rejectReady?: (error: Error) => void
  private busy = false
  private closed = false
  get isClosed() {
    return this.closed
  }
  private disposed = new vscode.EventEmitter<void>()
  readonly onDidDispose = this.disposed.event
  constructor(
    readonly uri: vscode.Uri,
    public session: DocxSession,
    public baseline: FileBaseline,
  ) {}
  prepareView() {
    this.ready = new Promise<void>((resolve, reject) => {
      this.resolveReady = resolve
      this.rejectReady = reject
    })
    void this.ready.catch(() => {})
  }
  viewReady(error?: string) {
    if (error) this.rejectReady?.(new Error(error))
    else this.resolveReady?.()
  }
  async exclusive<T>(work: () => T | Promise<T>): Promise<T> {
    if (this.closed) throw new Error('DOCX session is closed')
    if (this.busy) throw new Error('DOCX editor is busy saving or committing another operation')
    this.busy = true
    try {
      return await work()
    } finally {
      this.busy = false
    }
  }
  projection() {
    return {
      descriptor: this.session.descriptor(),
      content: this.session.content(),
      presentation: this.session.presentation(),
    }
  }
  async flush(freeze = false) {
    if (!this.bridge) return
    let timer: ReturnType<typeof setTimeout> | undefined
    try {
      await Promise.race([
        this.ready,
        new Promise<never>((_, reject) => {
          timer = setTimeout(() => reject(new Error('DOCX editor is busy initializing')), 10000)
        }),
      ])
    } finally {
      clearTimeout(timer)
    }
    const response = (await this.bridge.request('view.flush', { freeze })) as { ready: boolean }
    if (!response.ready) throw new Error('DOCX editor is busy composing or synchronizing')
  }
  async project() {
    const bridge = this.bridge
    try {
      await bridge?.request('view.resync', this.projection())
    } catch (error) {
      if (bridge === this.bridge) throw error
    }
  }
  dispose() {
    this.closed = true
    this.rejectReady?.(new Error('DOCX editor closed'))
    const bridge = this.bridge
    this.bridge = undefined
    bridge?.dispose()
    this.disposed.fire()
    this.disposed.dispose()
  }
}
