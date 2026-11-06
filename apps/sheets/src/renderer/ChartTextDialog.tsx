import { useRef, useState } from 'react'
import { useI18n, type StringKey } from './i18n/locale'
import { useModalDialog } from './modal-dialog'
export type ChartTextTarget = 'title' | 'axis-category' | 'axis-value'

const CHART_TEXT_LABELS: Record<ChartTextTarget, { heading: StringKey; command: string }> = {
  title: { heading: 'appChartTitleEl', command: 'chart-title' },
  'axis-category': { heading: 'appCategoryAxisTitle', command: 'chart-axis-cat' },
  'axis-value': { heading: 'appValueAxisTitle', command: 'chart-axis-val' },
}

/// Chart Design → Add Chart Element text prompt (title / axis titles).
/// An empty axis title removes that axis title; an empty chart title blanks it.
export function ChartTextDialog({
  target,
  initial,
  onCommand,
  onClose,
}: {
  readonly target: ChartTextTarget
  readonly initial: string
  readonly onCommand: (command: string) => void | Promise<void>
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [value, setValue] = useState(initial)
  const { heading, command } = CHART_TEXT_LABELS[target]
  const busy = useRef(false)
  const [pending, setPending] = useState(false)
  const [error, setError] = useState('')
  const close = () => {
    if (!busy.current) onClose()
  }
  const apply = async (): Promise<void> => {
    if (busy.current) return
    busy.current = true
    setPending(true)
    setError('')
    try {
      await onCommand(`${command}:${value.trim()}`)
      onClose()
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason))
    } finally {
      busy.current = false
      setPending(false)
    }
  }
  const modal = useModalDialog(close)
  return (
    <div className="dialog-backdrop" onClick={close}>
      <div
        className="format-cells-dialog link-dialog"
        role="dialog"
        {...modal}
        aria-label={t(heading)}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t(heading)}</header>
        <div className="dialog-body" aria-busy={pending}>
          <label>
            {t(target === 'title' ? 'appTitleText' : 'appAxisTitleText')}
            <input
              disabled={pending}
              autoFocus
              type="text"
              value={value}
              maxLength={255}
              onChange={(event) => setValue(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === 'Enter') void apply()
              }}
            />
          </label>
        </div>
        {error && (
          <div className="chart-editor-error" role="alert">
            {error}
          </div>
        )}
        <footer className="dialog-actions">
          <button disabled={pending} onClick={close}>
            {t('appCancel')}
          </button>
          <button disabled={pending} className="primary" onClick={() => void apply()}>
            {t('appOk')}
          </button>
        </footer>
      </div>
    </div>
  )
}
