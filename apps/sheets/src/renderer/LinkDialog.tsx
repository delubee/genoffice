import { useState } from 'react'
import { useI18n } from './i18n/locale'
import { useModalDialog } from './modal-dialog'

export function LinkDialog({
  currentTarget,
  onCommand,
  onClose,
  onApply,
}: {
  readonly currentTarget: string | null
  readonly onCommand: (command: string) => void
  readonly onClose: () => void
  readonly onApply?: (command: string) => boolean | Promise<boolean>
}): React.JSX.Element {
  const { t } = useI18n()
  const [value, setValue] = useState(currentTarget ?? '')
  const [pending, setPending] = useState(false)
  const close = () => {
    if (!pending) onClose()
  }
  const commit = async (command: string) => {
    if (pending) return
    if (!onApply) {
      onCommand(command)
      onClose()
      return
    }
    setPending(true)
    try {
      if (await onApply(command)) onClose()
    } finally {
      setPending(false)
    }
  }
  const apply = () => {
    if (value.trim()) void commit(`link-set:${encodeURIComponent(value)}`)
    else if (!onApply) onClose()
  }
  const modal = useModalDialog(close)
  return (
    <div className="dialog-backdrop" onClick={close}>
      <div
        className="format-cells-dialog link-dialog"
        role="dialog"
        {...modal}
        aria-label={t(currentTarget ? 'appEditLinkTitle' : 'appInsertLinkTitle')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t(currentTarget ? 'appEditLinkTitle' : 'appInsertLinkTitle')}</header>
        <div className="dialog-body">
          <label>
            {t('appLinkAddressLabel')}
            <input
              disabled={pending}
              autoFocus
              type="text"
              value={value}
              placeholder={t('appLinkPlaceholder')}
              onChange={(event) => setValue(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === 'Enter') apply()
              }}
            />
          </label>
        </div>
        <footer className="dialog-actions">
          {currentTarget && (
            <button
              disabled={pending}
              onClick={() => {
                void commit('link-remove')
              }}
            >
              {t('appRemoveLink')}
            </button>
          )}
          <button disabled={pending} onClick={close}>
            {t('appCancel')}
          </button>
          <button className="primary" disabled={pending || !value.trim()} onClick={apply}>
            {t('appOk')}
          </button>
        </footer>
      </div>
    </div>
  )
}
