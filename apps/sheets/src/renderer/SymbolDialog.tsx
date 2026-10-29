import { useRef, useState } from 'react'

import { useI18n, type StringKey } from './i18n/locale'
import { useModalDialog } from './modal-dialog'

/// Excel's Insert → Symbol, minimal: a tabbed grid of common special
/// characters. Clicking a symbol emits it to App (which appends it to the
/// active cell) and keeps the dialog open, so several symbols can be
/// inserted in a row — Close dismisses it.

/// Appends a picked symbol to whatever the cell already holds. Numbers and
/// booleans coerce to text: Excel's Symbol dialog also turns the cell into
/// text when characters are appended.
export function appendSymbol(
  existing: string | number | boolean | null | undefined,
  char: string,
): string {
  if (existing === null || existing === undefined || existing === '') return char
  return `${String(existing)}${char}`
}

const SYMBOL_CATEGORIES: readonly {
  readonly labelKey: StringKey
  readonly symbols: readonly string[]
}[] = [
  {
    labelKey: 'dlgSymbolCatCurrency',
    symbols: [
      '€',
      '£',
      '¥',
      '¢',
      '₩',
      '₽',
      '₹',
      '©',
      '®',
      '™',
      '§',
      '¶',
      '†',
      '‡',
      '•',
      '…',
      '°',
      '‰',
      '№',
      '¿',
    ],
  },
  {
    labelKey: 'dlgSymbolCatMath',
    symbols: [
      '±',
      '×',
      '÷',
      '≠',
      '≈',
      '≤',
      '≥',
      '∞',
      '√',
      '∑',
      '∏',
      '∫',
      '∂',
      '∆',
      'π',
      'µ',
      '∈',
      '∉',
      '⊂',
      '⊃',
      '∩',
      '∪',
      '∀',
      '∃',
      '∅',
      '∝',
      '∴',
      '∵',
      '⊥',
      '∥',
    ],
  },
  {
    labelKey: 'dlgSymbolCatArrows',
    symbols: [
      '←',
      '→',
      '↑',
      '↓',
      '↔',
      '↕',
      '⇐',
      '⇒',
      '⇑',
      '⇓',
      '⇔',
      '↗',
      '↘',
      '↙',
      '↖',
      '↩',
      '↪',
      '⇄',
      '⇅',
      '↦',
    ],
  },
  {
    labelKey: 'dlgSymbolCatGreek',
    symbols: [
      'α',
      'β',
      'γ',
      'δ',
      'ε',
      'ζ',
      'η',
      'θ',
      'λ',
      'μ',
      'π',
      'ρ',
      'σ',
      'τ',
      'φ',
      'χ',
      'ψ',
      'ω',
      'Α',
      'Β',
      'Γ',
      'Δ',
      'Θ',
      'Λ',
      'Σ',
      'Φ',
      'Ω',
    ],
  },
  {
    labelKey: 'dlgSymbolCatMisc',
    symbols: [
      '★',
      '☆',
      '●',
      '○',
      '◆',
      '◇',
      '■',
      '□',
      '▲',
      '△',
      '▼',
      '▽',
      '✓',
      '✗',
      '☑',
      '☐',
      '♠',
      '♣',
      '♥',
      '♦',
      '♪',
      '♫',
      '☀',
      '☁',
      '☂',
      '☎',
    ],
  },
]

export function SymbolDialog({
  onInsert,
  onClose,
}: {
  readonly onInsert: (char: string) => void | Promise<void>
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [categoryIndex, setCategoryIndex] = useState(0)
  const active = SYMBOL_CATEGORIES[categoryIndex]
  const pendingRef = useRef(false)
  const [pending, setPending] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const close = () => {
    if (!pendingRef.current) onClose()
  }
  const insert = async (symbol: string) => {
    if (pendingRef.current) return
    pendingRef.current = true
    setError(null)
    try {
      const result = onInsert(symbol)
      if (result && typeof result.then === 'function') {
        setPending(true)
        await result
      }
    } catch (error) {
      setError(error instanceof Error ? error.message : String(error))
    } finally {
      pendingRef.current = false
      setPending(false)
    }
  }

  const modal = useModalDialog(close)
  return (
    <div className="dialog-backdrop" onClick={close}>
      <div
        className="format-cells-dialog"
        role="dialog"
        {...modal}
        aria-label={t('dlgSymbolTitle')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('dlgSymbolTitle')}</header>
        <nav className="dialog-tabs">
          {SYMBOL_CATEGORIES.map((candidate, index) => (
            <button
              disabled={pending}
              key={candidate.labelKey}
              className={index === categoryIndex ? 'active' : ''}
              onClick={() => setCategoryIndex(index)}
            >
              {t(candidate.labelKey)}
            </button>
          ))}
        </nav>
        <section className="dialog-body">
          <div className="symbol-grid">
            {(active?.symbols ?? []).map((symbol) => (
              <button
                disabled={pending}
                key={symbol}
                type="button"
                data-tip={t('dlgSymbolInsertHint', { symbol })}
                onClick={() => {
                  void insert(symbol)
                }}
              >
                {symbol}
              </button>
            ))}
          </div>
          {error && (
            <p className="dialog-note dialog-error" role="alert">
              {error}
            </p>
          )}
          <p className="dialog-note">{t('dlgSymbolNote')}</p>
        </section>
        <div className="dialog-actions">
          <button className="secondary" disabled={pending} onClick={close}>
            {t('dlgClose')}
          </button>
        </div>
      </div>
    </div>
  )
}
