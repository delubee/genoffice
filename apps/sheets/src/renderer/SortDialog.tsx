import { useRef, useState } from 'react'
import { Dropdown } from '@genoffice/ui/chrome'
import { normalizeLang } from '@genoffice/i18n'
import { useI18n, tFor, type TFunc } from './i18n/locale'
import { useModalDialog } from './modal-dialog'

const NO_SORT_LEVEL = -1

export function SortDialog({
  columns,
  locale,
  onCommand,
  onClose,
}: {
  readonly locale?: string
  readonly columns: { label: string; colIndex: number }[]
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
  const [hasHeader, setHasHeader] = useState(true)
  const first = columns[0]?.colIndex ?? NO_SORT_LEVEL
  const [levels, setLevels] = useState<{ colIndex: number; order: 'a' | 'd' }[]>([
    { colIndex: first, order: 'a' },
    { colIndex: NO_SORT_LEVEL, order: 'a' },
    { colIndex: NO_SORT_LEVEL, order: 'a' },
  ])
  const apply = async (): Promise<void> => {
    if (pending.current) return
    const rules = levels
      .filter((level) => level.colIndex !== NO_SORT_LEVEL)
      .map((level) => `${level.colIndex}${level.order}`)
    if (!rules.length) return
    pending.current = true
    setBusy(true)
    setError('')
    try {
      await onCommand(`sort-custom:${hasHeader ? 1 : 0}:${rules.join(',')}`)
      pending.current = false
      onClose()
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason))
    } finally {
      pending.current = false
      setBusy(false)
    }
  }
  const setLevel = (
    index: number,
    patch: Partial<{ colIndex: number; order: 'a' | 'd' }>,
  ): void => {
    setLevels((previous) =>
      previous.map((level, at) => (at === index ? { ...level, ...patch } : level)),
    )
  }
  const modal = useModalDialog(close)
  return (
    <div className="dialog-backdrop" onClick={close}>
      <div
        className="format-cells-dialog sort-dialog"
        role="dialog"
        {...modal}
        aria-label={t('appCustomSort')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('appSort')}</header>
        <fieldset
          className="dialog-body"
          disabled={busy}
          style={{ border: 0, margin: 0, minWidth: 0 }}
        >
          <label className="sort-header-check">
            <input
              type="checkbox"
              checked={hasHeader}
              onChange={(event) => setHasHeader(event.target.checked)}
            />
            {t('appMyDataHasHeaders')}
          </label>
          {levels.map((level, index) => (
            <div className="sort-level" key={index}>
              <span>{t(index === 0 ? 'appSortBy' : 'appThenBy')}</span>
              <Dropdown
                ariaLabel={t(index === 0 ? 'appSortBy' : 'appThenBy')}
                value={String(level.colIndex)}
                options={[
                  ...(index > 0 ? [{ value: String(NO_SORT_LEVEL), label: t('appSortNone') }] : []),
                  ...columns.map((column) => ({
                    value: String(column.colIndex),
                    label: column.label,
                  })),
                ]}
                onPick={(v) => setLevel(index, { colIndex: Number(v) })}
              />
              <Dropdown
                value={level.order}
                options={[
                  { value: 'a', label: t('appSortAsc') },
                  { value: 'd', label: t('appSortDesc') },
                ]}
                onPick={(v) => setLevel(index, { order: v })}
              />
            </div>
          ))}
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
          <button
            className="primary"
            disabled={busy || levels.every((level) => level.colIndex === NO_SORT_LEVEL)}
            onClick={() => void apply()}
          >
            {t('appOk')}
          </button>
        </footer>
      </div>
    </div>
  )
}
