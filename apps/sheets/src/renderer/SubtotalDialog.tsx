import { useRef, useState } from 'react'
import { Dropdown } from '@genoffice/ui/chrome'
import type { PivotField } from './PivotDialog'
import { normalizeLang } from '@genoffice/i18n'
import { useI18n, tFor, type TFunc } from './i18n/locale'
import type { SubtotalConfig } from '@genoffice/xlsx-gateway/domain/subtotal-plan'
export type { SubtotalConfig } from '@genoffice/xlsx-gateway/domain/subtotal-plan'
import { useModalDialog } from './modal-dialog'

/// Excel's Data → Subtotal, minimal: one "at each change in" column, one
/// aggregated column, SUBTOTAL rows inserted per group plus a grand total.

export function SubtotalDialog({
  fields,
  locale,
  onCreate,
  onClose,
}: {
  readonly locale?: string
  readonly fields: readonly PivotField[]
  /// Returns an error message, or null on success.
  readonly onCreate: (config: SubtotalConfig) => string | null | Promise<string | null>
  readonly onClose: () => void
}): React.JSX.Element {
  const { t: desktopT } = useI18n()
  const t: TFunc = locale ? (key, params) => tFor(normalizeLang(locale), key, params) : desktopT
  const pending = useRef(false)
  const [busy, setBusy] = useState(false)
  const [groupCol, setGroupCol] = useState(fields[0]?.colIndex ?? 0)
  const [valueCol, setValueCol] = useState(fields[1]?.colIndex ?? fields[0]?.colIndex ?? 0)
  const [agg, setAgg] = useState<SubtotalConfig['agg']>('sum')
  const [error, setError] = useState<string | null>(null)

  const close = () => {
    if (!pending.current) onClose()
  }
  const apply = async () => {
    if (pending.current) return
    pending.current = true
    setBusy(true)
    setError(null)
    try {
      const failure = await onCreate({ groupCol, valueCol, agg })
      setError(failure)
      if (failure === null) onClose()
    } catch (reason) {
      const message = reason instanceof Error ? reason.message : String(reason)
      const key = (
        ['appSubtotalSameColumns', 'appSubtotalTooManyGroups', 'appSelectRangeWithHeader'] as const
      ).find((key) => message.includes(key))
      setError(key ? t(key) : message)
    } finally {
      pending.current = false
      setBusy(false)
    }
  }
  const modal = useModalDialog(close)
  return (
    <div className="dialog-backdrop" onClick={close}>
      <div
        className="format-cells-dialog"
        role="dialog"
        {...modal}
        aria-label={t('dlgSubtotalTitle')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('dlgSubtotalTitle')}</header>
        {fields.length === 0 ? (
          <p className="dialog-note">{t('dlgSubtotalNoFields')}</p>
        ) : (
          <fieldset
            className="dialog-grid"
            disabled={busy}
            style={{ border: 0, margin: 0, minWidth: 0 }}
          >
            <label>
              {t('dlgSubtotalGroupBy')}
              <Dropdown
                ariaLabel={t('dlgSubtotalGroupBy')}
                value={String(groupCol)}
                options={fields.map((field) => ({
                  value: String(field.colIndex),
                  label: field.label,
                }))}
                onPick={(v) => setGroupCol(Number(v))}
              />
            </label>
            <label>
              {t('dlgSubtotalFunction')}
              <Dropdown
                ariaLabel={t('dlgSubtotalFunction')}
                value={agg}
                options={[
                  { value: 'sum', label: t('dlgPivotAggSum') },
                  { value: 'count', label: t('dlgPivotAggCount') },
                  { value: 'average', label: t('dlgPivotAggAverage') },
                ]}
                onPick={(v) => setAgg(v as SubtotalConfig['agg'])}
              />
            </label>
            <label>
              {t('dlgSubtotalAddTo')}
              <Dropdown
                ariaLabel={t('dlgSubtotalAddTo')}
                value={String(valueCol)}
                options={fields.map((field) => ({
                  value: String(field.colIndex),
                  label: field.label,
                }))}
                onPick={(v) => setValueCol(Number(v))}
              />
            </label>
            <p className="dialog-note">{t('dlgSubtotalNote')}</p>
          </fieldset>
        )}
        {error && (
          <p className="dialog-note" role="alert">
            {error}
          </p>
        )}
        <div className="dialog-actions">
          <button className="secondary" onClick={close} disabled={busy}>
            {t('dlgCancel')}
          </button>
          {fields.length > 0 && (
            <button className="primary-action" disabled={busy} onClick={() => void apply()}>
              {t('dlgOk')}
            </button>
          )}
        </div>
      </div>
    </div>
  )
}
