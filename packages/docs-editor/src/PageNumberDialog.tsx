import { useState } from 'react'
import { t } from '../../../apps/docs/src/renderer/i18n/locale'
import type { DocumentSettings } from './shared/protocol'

const FORMATS: Array<[string, string]> = [
  ['decimal', '1, 2, 3, …'],
  ['numberInDash', '- 1 -, - 2 -, - 3 -, …'],
  ['lowerLetter', 'a, b, c, …'],
  ['upperLetter', 'A, B, C, …'],
  ['lowerRoman', 'i, ii, iii, …'],
  ['upperRoman', 'I, II, III, …'],
  ['chineseCounting', '一, 二, 三, …'],
]

/** Header & Footer → Page Number → Format Page Numbers (the final section's w:pgNumType). */
export function PageNumberDialog({
  value,
  onApply,
  onClose,
}: {
  value: NonNullable<DocumentSettings['pageNumbering']>
  onApply(value: NonNullable<DocumentSettings['pageNumbering']>): void
  onClose(): void
}) {
  const [fmt, setFmt] = useState(value.fmt ?? 'decimal')
  const [start, setStart] = useState(value.start !== undefined ? String(value.start) : '')
  return (
    <div className="modal-backdrop" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className="modal">
        <h2>{t('appPageNumFormatTitle')}</h2>
        <label className="pgnum-row">
          {t('appNumberFormat')}
          <select
            aria-label={t('appNumberFormat')}
            value={fmt}
            onChange={(event) => setFmt(event.target.value)}
          >
            {FORMATS.map(([id, label]) => (
              <option key={id} value={id}>
                {label}
              </option>
            ))}
          </select>
        </label>
        <label className="pgnum-row">
          {t('appStartAt')}
          <input
            type="number"
            min={0}
            max={9999}
            placeholder={t('appContinueFromPrev')}
            value={start}
            onChange={(event) => setStart(event.target.value)}
          />
        </label>
        <p className="pgnum-hint">{t('appPgNumHintBlank')}</p>
        <div className="modal-actions">
          <button onClick={onClose}>{t('appCancel')}</button>
          <button
            className="btn-primary"
            onClick={() =>
              onApply({
                ...(fmt !== 'decimal' ? { fmt } : {}),
                ...(start.trim() !== ''
                  ? { start: Math.min(9999, Math.max(0, parseInt(start, 10) || 0)) }
                  : {}),
              })
            }
          >
            {t('appOk')}
          </button>
        </div>
      </div>
    </div>
  )
}
