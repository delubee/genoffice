import { randomUUID } from 'node:crypto'
import { WorkbookError } from './domain'
import type { FileBaseline } from './publish'

interface Owner {
  documentUri: string
  epoch: string
  sessionId?: string
  baseline: FileBaseline
  released: boolean
}
export interface WorkbookLease {
  readonly epoch: string
  readonly baseline: FileBaseline
  bind(sessionId: string): void
  update(baseline: FileBaseline): void
  release(): void
}

/** Owns writer aliases in this extension host; file baselines protect other processes. */
export class WorkbookRegistry {
  private uris = new Map<string, Owner>()
  private keys = new Map<string, Owner>()
  private sessions = new Map<string, Owner>()
  private aliases(baseline: FileBaseline) {
    const path =
      process.platform === 'win32' ? baseline.canonicalPath.toLowerCase() : baseline.canonicalPath
    return [`path:${path}`, ...(baseline.physicalId ? [`physical:${baseline.physicalId}`] : [])]
  }
  private check(owner: Owner, baseline: FileBaseline) {
    for (const key of this.aliases(baseline)) {
      const existing = this.keys.get(key)
      if (existing && existing !== owner)
        throw new WorkbookError(
          'external_change_conflict',
          `This workbook is already open at ${existing.documentUri}; use its existing editor.`,
          { documentUri: existing.documentUri },
        )
    }
  }
  reserve(documentUri: string, baseline: FileBaseline): WorkbookLease {
    if (this.uris.has(documentUri))
      throw new WorkbookError('external_change_conflict', 'This URI is already open or opening.')
    const owner: Owner = {
      documentUri,
      epoch: randomUUID(),
      baseline: { ...baseline },
      released: false,
    }
    this.check(owner, baseline)
    this.uris.set(documentUri, owner)
    for (const key of this.aliases(baseline)) this.keys.set(key, owner)
    const active = () => {
      if (owner.released)
        throw new WorkbookError('session_closed', 'The document lease has closed.')
    }
    return {
      epoch: owner.epoch,
      get baseline() {
        return { ...owner.baseline }
      },
      bind: (sessionId) => {
        active()
        const existing = this.sessions.get(sessionId)
        if (existing && existing !== owner)
          throw new WorkbookError(
            'invalid_document',
            'This session ID belongs to another document.',
          )
        if (owner.sessionId) this.sessions.delete(owner.sessionId)
        owner.sessionId = sessionId
        this.sessions.set(sessionId, owner)
      },
      update: (next) => {
        active()
        this.check(owner, next)
        for (const key of this.aliases(owner.baseline))
          if (this.keys.get(key) === owner) this.keys.delete(key)
        owner.baseline = { ...next }
        for (const key of this.aliases(next)) this.keys.set(key, owner)
      },
      release: () => {
        if (owner.released) return
        owner.released = true
        if (this.uris.get(documentUri) === owner) this.uris.delete(documentUri)
        if (owner.sessionId && this.sessions.get(owner.sessionId) === owner)
          this.sessions.delete(owner.sessionId)
        for (const key of this.aliases(owner.baseline))
          if (this.keys.get(key) === owner) this.keys.delete(key)
      },
    }
  }
  session(sessionId: string) {
    const owner = this.sessions.get(sessionId)
    return owner ? { documentUri: owner.documentUri, epoch: owner.epoch, sessionId } : undefined
  }
}
