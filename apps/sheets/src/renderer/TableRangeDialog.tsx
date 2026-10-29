import { useState } from 'react'

import { useI18n } from './i18n/locale'
import { useModalDialog } from './modal-dialog'

/// Create Table (Ctrl+T) / Resize Table: one range box, optionally the
/// "My table has headers" check.
export function TableRangeDialog({
  title,
  note,
  initialRange,
  headersOption = false,
  onSubmit,
  onClose,
}: {
  readonly title: string
  readonly note: string
  readonly initialRange: string
  readonly headersOption?: boolean
  readonly onSubmit: (range: string, headers: boolean) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [range, setRange] = useState(initialRange)
  const [headers, setHeaders] = useState(true)
  const modal = useModalDialog(onClose)
  const submit = (): void => {
    if (range.trim().length === 0) return
    onSubmit(range.trim().toUpperCase(), headers)
    onClose()
  }
  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog sort-dialog"
        role="dialog"
        {...modal}
        aria-label={title}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{title}</header>
        <div className="dialog-body">
          <p className="dialog-note">{note}</p>
          <input
            className="table-range-input"
            autoFocus
            value={range}
            spellCheck={false}
            onChange={(event) => setRange(event.target.value)}
            onKeyDown={(event) => {
              if (event.key === 'Enter') submit()
            }}
          />
          {headersOption && (
            <label className="sort-header-check">
              <input
                type="checkbox"
                checked={headers}
                onChange={(event) => setHeaders(event.target.checked)}
              />
              {t('appMyTableHasHeaders')}
            </label>
          )}
        </div>
        <footer className="dialog-actions">
          <button onClick={onClose}>{t('appCancel')}</button>
          <button className="primary" onClick={submit}>
            {t('appOk')}
          </button>
        </footer>
      </div>
    </div>
  )
}
