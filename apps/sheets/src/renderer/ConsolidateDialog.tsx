import { useRef, useState } from 'react'

import { Dropdown } from '@genoffice/ui/chrome'

import {
  isValidConsolidateReference,
  type ConsolidateConfig,
  type ConsolidateFn,
} from './consolidate'
import { normalizeLang } from '@genoffice/i18n'
import { useI18n, tFor, type TFunc } from './i18n/locale'
import { useModalDialog } from './modal-dialog'

/// Excel's Data → Consolidate: a function, a list of source references, and
/// an optional "use labels in left column" mode. Output lands at the active
/// cell as live formulas over the sources.

export function ConsolidateDialog({
  defaultReference,
  targetLabel,
  locale,
  onCreate,
  onClose,
}: {
  readonly locale?: string
  /// Prefill for the reference input, usually the current selection.
  readonly defaultReference: string
  readonly targetLabel: string
  /// Returns an error message, or null on success.
  readonly onCreate: (config: ConsolidateConfig) => string | null | Promise<string | null>
  readonly onClose: () => void
}): React.JSX.Element {
  const { t: desktopT } = useI18n()
  const t: TFunc = locale ? (key, params) => tFor(normalizeLang(locale), key, params) : desktopT
  const pendingApply = useRef(false)
  const [busy, setBusy] = useState(false)
  const [fn, setFn] = useState<ConsolidateFn>('sum')
  const [references, setReferences] = useState<string[]>([])
  const [input, setInput] = useState(defaultReference)
  const [leftLabels, setLeftLabels] = useState(false)
  const [error, setError] = useState<string | null>(null)

  const addReference = (): string | null => {
    const trimmed = input.trim()
    if (trimmed === '') return null
    if (!isValidConsolidateReference(trimmed)) {
      return t('dlgConsBadRef', { ref: trimmed })
    }
    if (references.includes(trimmed)) {
      setInput('')
      return null
    }
    setReferences([...references, trimmed])
    setInput('')
    return null
  }

  const close = () => {
    if (!pendingApply.current) onClose()
  }
  const apply = async () => {
    if (pendingApply.current) return
    // A typed-but-not-added reference still counts on OK.
    const pending = input.trim()
    let all = references
    if (pending !== '') {
      if (!isValidConsolidateReference(pending)) {
        setError(t('dlgConsBadRef', { ref: pending }))
        return
      }
      if (!references.includes(pending)) all = [...references, pending]
    }
    if (all.length === 0) {
      setError(t('dlgConsNeedOneRef'))
      return
    }
    pendingApply.current = true
    setBusy(true)
    setError(null)
    try {
      const failure = await onCreate({ fn, references: all, leftLabels })
      setError(failure)
      if (failure === null) onClose()
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason))
    } finally {
      pendingApply.current = false
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
        aria-label={t('dlgConsTitle')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('dlgConsTitle')}</header>
        <fieldset
          className="dialog-grid"
          disabled={busy}
          style={{ border: 0, margin: 0, minWidth: 0 }}
        >
          <label>
            {t('dlgConsFunction')}
            {/* Label-based consolidation supports sum, count and average. */}
            <Dropdown
              ariaLabel={t('dlgConsFunction')}
              value={fn}
              options={[
                { value: 'sum', label: t('dlgPivotAggSum') },
                { value: 'count', label: t('dlgPivotAggCount') },
                { value: 'average', label: t('dlgPivotAggAverage') },
                ...(leftLabels
                  ? []
                  : [
                      { value: 'max', label: t('dlgPivotAggMax') },
                      { value: 'min', label: t('dlgPivotAggMin') },
                    ]),
              ]}
              onPick={(v) => setFn(v as ConsolidateFn)}
            />
          </label>
          <label className="dialog-span">
            {t('dlgConsReference')}
            <span className="input-row">
              <input
                value={input}
                placeholder="Sheet2!A1:C5"
                onChange={(event) => setInput(event.target.value)}
                onKeyDown={(event) => {
                  if (event.key === 'Enter') {
                    event.preventDefault()
                    setError(addReference())
                  }
                }}
              />
              <button className="secondary" onClick={() => setError(addReference())}>
                {t('dlgConsAdd')}
              </button>
            </span>
          </label>
          {references.length > 0 && (
            <ul className="reference-list" aria-label={t('dlgConsAllRefs')}>
              {references.map((reference) => (
                <li key={reference}>
                  <code>{reference}</code>
                  <button
                    className="secondary"
                    aria-label={t('dlgConsDeleteRef', { ref: reference })}
                    onClick={() => setReferences(references.filter((entry) => entry !== reference))}
                  >
                    {t('dlgConsDelete')}
                  </button>
                </li>
              ))}
            </ul>
          )}
          <label className="dialog-check dialog-span">
            <input
              type="checkbox"
              checked={leftLabels}
              onChange={(event) => {
                setLeftLabels(event.target.checked)
                if (event.target.checked && (fn === 'max' || fn === 'min')) setFn('sum')
              }}
            />
            {t('dlgConsLeftLabels')}
          </label>
          <p className="dialog-note dialog-span">
            {leftLabels
              ? t('dlgConsNoteLabels', { target: targetLabel })
              : t('dlgConsNotePosition', { target: targetLabel })}
          </p>
        </fieldset>
        {error && (
          <p className="dialog-note" role="alert">
            {error}
          </p>
        )}
        <div className="dialog-actions">
          <button className="secondary" onClick={close} disabled={busy}>
            {t('dlgCancel')}
          </button>
          <button className="primary-action" disabled={busy} onClick={() => void apply()}>
            {t('dlgOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
