import { useRef, useState } from 'react'
import { normalizeLang } from '@genoffice/i18n'
import { useI18n, tFor, type TFunc } from './i18n/locale'
import { useModalDialog } from './modal-dialog'

export function RemoveDuplicatesDialog({
  locale,
  onCommand,
  onClose,
}: {
  readonly locale?: string
  readonly onCommand: (command: string) => void | Promise<void>
  readonly onClose: () => void
}): React.JSX.Element {
  const { t: desktopT } = useI18n()
  const t: TFunc = locale ? (key, params) => tFor(normalizeLang(locale), key, params) : desktopT
  const [busy, setBusy] = useState(false)
  const pending = useRef(false)
  const [error, setError] = useState('')
  const close = () => {
    if (!pending.current) onClose()
  }
  const apply = async () => {
    if (pending.current) return
    pending.current = true
    setBusy(true)
    setError('')
    try {
      await onCommand(`remove-duplicates:${hasHeader ? 1 : 0}`)
      pending.current = false
      onClose()
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason))
    } finally {
      pending.current = false
      setBusy(false)
    }
  }
  const [hasHeader, setHasHeader] = useState(true)
  const modal = useModalDialog(close)
  return (
    <div className="dialog-backdrop" onClick={close}>
      <div
        className="format-cells-dialog sort-dialog"
        role="dialog"
        {...modal}
        aria-label={t('appRemoveDuplicates')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('appRemoveDuplicates')}</header>
        <fieldset
          className="dialog-body"
          disabled={busy}
          style={{ border: 0, margin: 0, minWidth: 0 }}
        >
          <p className="dialog-note">{t('appRemoveDuplicatesNote')}</p>
          <label className="sort-header-check">
            <input
              type="checkbox"
              checked={hasHeader}
              onChange={(event) => setHasHeader(event.target.checked)}
            />
            {t('appMyDataHasHeaders')}
          </label>
        </fieldset>
        {error && (
          <p role="alert" className="dialog-note">
            {error}
          </p>
        )}
        <footer className="dialog-actions">
          <button onClick={close} disabled={busy}>
            {t('appCancel')}
          </button>
          <button className="primary" disabled={busy} onClick={() => void apply()}>
            {t('appOk')}
          </button>
        </footer>
      </div>
    </div>
  )
}
