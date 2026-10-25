import { useState } from 'react'
import type { OfficeActionLabels } from './strings-office-actions'

/** Shared menu entry points; each native host owns resource selection, saving and chat. */
export function OfficeFileActions({
  labels,
  onOpen,
  onSave,
  onSaveAs,
  onChat,
  onError,
}: {
  labels: OfficeActionLabels
  onOpen(): Promise<void>
  onSave(): Promise<void>
  onSaveAs(): Promise<void>
  onChat(scope: 'document' | 'selection'): Promise<void>
  onError(error: unknown): void
}) {
  const [busy, setBusy] = useState(false)
  const run = async (action: () => Promise<void>) => {
    if (busy) return
    setBusy(true)
    try {
      await action()
    } catch (error) {
      onError(error)
    } finally {
      setBusy(false)
    }
  }
  return (
    <span role="group" aria-label="GenOffice" style={{ display: 'inline-flex', gap: 4 }}>
      <button disabled={busy} title={labels.open} onClick={() => void run(onOpen)}>
        {labels.open}
      </button>
      <button disabled={busy} title={labels.save} onClick={() => void run(onSave)}>
        {labels.save}
      </button>
      <button disabled={busy} title={labels.saveAs} onClick={() => void run(onSaveAs)}>
        {labels.saveAs}
      </button>
      <button
        disabled={busy}
        title={labels.chat}
        onClick={() => void run(() => onChat('document'))}
      >
        {labels.chat}
      </button>
      <button
        disabled={busy}
        title={labels.chatSelection}
        onClick={() => void run(() => onChat('selection'))}
      >
        {labels.chatSelection}
      </button>
    </span>
  )
}
