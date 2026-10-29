import { useState } from 'react'
import { useI18n } from './i18n/locale'
import { useModalDialog } from './modal-dialog'

export function AxisSizeDialog({
  axis,
  onCommand,
  onClose,
  onApply,
}: {
  readonly axis: 'row' | 'col'
  readonly onCommand: (command: string) => void
  readonly onClose: () => void
  readonly onApply?: (command: string) => Promise<boolean>
}): React.JSX.Element {
  const { t } = useI18n()
  const [value, setValue] = useState('')
  const [pending, setPending] = useState(false)
  const close = () => {
    if (!pending) onClose()
  }
  const max = axis === 'row' ? 409.5 : 255
  const parsed = Number(value.trim().replace(',', '.'))
  const valid = value.trim() !== '' && Number.isFinite(parsed) && parsed >= 0 && parsed <= max
  const apply = (): void => {
    if (!valid || pending) return
    if (onApply) {
      setPending(true)
      void onApply(`${axis === 'row' ? 'row-height' : 'col-width'}:${parsed}`)
        .then((success) => {
          if (success) onClose()
        })
        .finally(() => setPending(false))
      return
    }
    onCommand(`${axis === 'row' ? 'row-height' : 'col-width'}:${parsed}`)
    onClose()
  }
  const modal = useModalDialog(close)
  return (
    <div className="dialog-backdrop" onClick={close}>
      <div
        className="format-cells-dialog link-dialog"
        role="dialog"
        {...modal}
        aria-label={t(axis === 'row' ? 'appRowHeight' : 'appColWidth')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t(axis === 'row' ? 'appRowHeight' : 'appColWidth')}</header>
        <div className="dialog-body">
          <label>
            {t(axis === 'row' ? 'appRowHeightLabel' : 'appColWidthLabel')}
            <input
              autoFocus
              disabled={pending}
              type="text"
              inputMode="decimal"
              value={value}
              onChange={(event) => setValue(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === 'Enter') apply()
              }}
            />
          </label>
        </div>
        <footer className="dialog-actions">
          <button disabled={pending} onClick={close}>
            {t('appCancel')}
          </button>
          <button className="primary" disabled={!valid || pending} onClick={apply}>
            {t('appOk')}
          </button>
        </footer>
      </div>
    </div>
  )
}
