import { useState } from 'react'
import { dvRuleSchema, type DvRule } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import type { WorkbookRuntime } from './runtime'
import { editorText } from './i18n'

export function DataValidationDialog({
  runtime,
  target,
  initial,
  locale,
  onClose,
  onError,
}: {
  runtime: WorkbookRuntime
  target: {
    target: { sheetId: string; address: string }
    stateId: string
    revision: number
  }
  initial?: DvRule
  locale?: string
  onClose(): void
  onError(message: string): void
}) {
  const text = editorText(locale)
  const [kind, setKind] = useState<DvRule['kind']>(initial?.kind ?? 'list'),
    [items, setItems] = useState(initial?.kind === 'list' ? initial.values.join(',') : 'Yes,No'),
    [source, setSource] = useState(initial?.kind === 'listRef' ? initial.range : 'A1:A10'),
    [minimum, setMinimum] = useState(
      initial?.kind === 'numberBetween'
        ? String(initial.min)
        : initial?.kind === 'dateBetween'
          ? initial.start
          : '0',
    ),
    [maximum, setMaximum] = useState(
      initial?.kind === 'numberBetween'
        ? String(initial.max)
        : initial?.kind === 'dateBetween'
          ? initial.end
          : '100',
    ),
    [formula, setFormula] = useState(initial?.kind === 'formula' ? initial.formula : '=TRUE'),
    [criterionType, setCriterionType] = useState<
      'whole' | 'decimal' | 'date' | 'time' | 'textLength'
    >(initial?.kind === 'criterion' ? initial.type : 'whole'),
    [operator, setOperator] = useState<Extract<DvRule, { kind: 'criterion' }>['operator']>(
      initial?.kind === 'criterion' ? initial.operator : 'between',
    ),
    [criterionFormula1, setCriterionFormula1] = useState(
      initial?.kind === 'criterion' ? initial.formula1 : '0',
    ),
    [criterionFormula2, setCriterionFormula2] = useState(
      initial?.kind === 'criterion' ? (initial.formula2 ?? '100') : '100',
    ),
    [allowBlank, setAllowBlank] = useState(initial?.options?.allowBlank ?? true),
    [showInputMessage, setShowInputMessage] = useState(initial?.options?.showInputMessage ?? false),
    [promptTitle, setPromptTitle] = useState(initial?.options?.promptTitle ?? ''),
    [prompt, setPrompt] = useState(initial?.options?.prompt ?? ''),
    [showErrorMessage, setShowErrorMessage] = useState(initial?.options?.showErrorMessage ?? true),
    [errorStyle, setErrorStyle] = useState(initial?.options?.errorStyle ?? 'stop'),
    [errorTitle, setErrorTitle] = useState(initial?.options?.errorTitle ?? ''),
    [error, setError] = useState(initial?.options?.error ?? ''),
    [showDropdown, setShowDropdown] = useState(initial?.options?.showDropdown ?? true)
  const apply = () => {
    try {
      const options = {
        allowBlank,
        showInputMessage,
        ...(promptTitle ? { promptTitle } : {}),
        ...(prompt ? { prompt } : {}),
        showErrorMessage,
        errorStyle,
        ...(errorTitle ? { errorTitle } : {}),
        ...(error ? { error } : {}),
        ...(kind === 'list' || kind === 'listRef' ? { showDropdown } : {}),
      }
      let value: unknown
      switch (kind) {
        case 'list':
          value = { kind, values: items.split(',').map((item) => item.trim()), options }
          break
        case 'any':
          value = { kind, options }
          break
        case 'listRef':
          value = { kind, range: source.trim(), options }
          break
        case 'numberBetween':
          value = { kind, min: Number(minimum), max: Number(maximum), options }
          break
        case 'dateBetween':
          value = { kind, start: minimum, end: maximum, options }
          break
        case 'checkbox':
          value = { kind, options }
          break
        case 'formula':
          value = { kind, formula: formula.trim(), options }
          break
        case 'criterion':
          value = {
            kind,
            type: criterionType,
            operator,
            formula1: criterionFormula1.trim(),
            ...(operator === 'between' || operator === 'notBetween'
              ? { formula2: criterionFormula2.trim() }
              : {}),
            options,
          }
          break
      }
      const rule = dvRuleSchema.parse(value)
      onError('')
      void runtime
        .setDataValidation(rule, target)
        .then(onClose)
        .catch((error) => onError(String(error)))
    } catch (error) {
      onError(error instanceof Error ? error.message : String(error))
    }
  }
  return (
    <div
      role="presentation"
      onMouseDown={(event) => {
        if (event.target === event.currentTarget) onClose()
      }}
      style={{
        position: 'fixed',
        inset: 0,
        zIndex: 10000,
        display: 'grid',
        placeItems: 'center',
        background: 'rgba(0,0,0,.35)',
      }}
    >
      <section
        role="dialog"
        aria-modal="true"
        aria-labelledby="workbook-dv-title"
        style={{
          width: 'min(440px, calc(100vw - 32px))',
          maxHeight: 'calc(100vh - 32px)',
          overflowY: 'auto',
          padding: 20,
          border: '1px solid var(--vscode-panel-border, #888)',
          borderRadius: 6,
          background: 'var(--vscode-editorWidget-background, Canvas)',
          color: 'var(--vscode-editorWidget-foreground, CanvasText)',
          boxShadow: '0 8px 28px rgba(0,0,0,.3)',
        }}
      >
        <h2 id="workbook-dv-title" style={{ marginTop: 0 }}>
          {text.dataValidation}
        </h2>
        <p>
          {text.validationSelection} <code>{target.target.address}</code>
        </p>
        <label>
          {text.validationType}
          <select value={kind} onChange={(event) => setKind(event.target.value as DvRule['kind'])}>
            <option value="any">{text.validationAny}</option>
            <option value="list">{text.validationList}</option>
            <option value="listRef">{text.validationListRef}</option>
            <option value="numberBetween">{text.validationNumber}</option>
            <option value="dateBetween">{text.validationDate}</option>
            <option value="checkbox">{text.validationCheckbox}</option>
            <option value="formula">{text.validationFormula}</option>
            <option value="criterion">{text.validationCriteria}</option>
          </select>
        </label>
        {kind === 'list' && (
          <label>
            {text.validationItems}
            <input value={items} onChange={(event) => setItems(event.target.value)} />
          </label>
        )}
        {kind === 'listRef' && (
          <label>
            {text.validationSource}
            <input value={source} onChange={(event) => setSource(event.target.value)} />
          </label>
        )}
        {kind === 'numberBetween' && (
          <>
            <label>
              {text.validationMinimum}
              <input
                type="number"
                value={minimum}
                onChange={(event) => setMinimum(event.target.value)}
              />
            </label>
            <label>
              {text.validationMaximum}
              <input
                type="number"
                value={maximum}
                onChange={(event) => setMaximum(event.target.value)}
              />
            </label>
          </>
        )}
        {kind === 'dateBetween' && (
          <>
            <label>
              {text.validationStartDate}
              <input
                type="date"
                value={minimum}
                onChange={(event) => setMinimum(event.target.value)}
              />
            </label>
            <label>
              {text.validationEndDate}
              <input
                type="date"
                value={maximum}
                onChange={(event) => setMaximum(event.target.value)}
              />
            </label>
          </>
        )}
        {kind === 'formula' && (
          <label>
            {text.validationFormula}
            <input value={formula} onChange={(event) => setFormula(event.target.value)} />
          </label>
        )}
        {kind === 'criterion' && (
          <>
            <label>
              {text.validationCriteria}
              <select
                value={criterionType}
                onChange={(event) => setCriterionType(event.target.value as typeof criterionType)}
              >
                <option value="whole">{text.validationWhole}</option>
                <option value="decimal">{text.validationDecimal}</option>
                <option value="date">{text.validationDate}</option>
                <option value="time">{text.validationTime}</option>
                <option value="textLength">{text.validationTextLength}</option>
              </select>
            </label>
            <label>
              {text.validationOperator}
              <select
                value={operator}
                onChange={(event) => setOperator(event.target.value as typeof operator)}
              >
                <option value="between">{text.validationBetween}</option>
                <option value="notBetween">{text.validationNotBetween}</option>
                <option value="equal">{text.validationEqual}</option>
                <option value="notEqual">{text.validationNotEqual}</option>
                <option value="greaterThan">{text.validationGreaterThan}</option>
                <option value="greaterThanOrEqual">{text.validationGreaterOrEqual}</option>
                <option value="lessThan">{text.validationLessThan}</option>
                <option value="lessThanOrEqual">{text.validationLessOrEqual}</option>
              </select>
            </label>
            <label>
              {text.validationFormula}
              <input
                value={criterionFormula1}
                onChange={(event) => setCriterionFormula1(event.target.value)}
              />
            </label>
            {(operator === 'between' || operator === 'notBetween') && (
              <label>
                {text.validationSecondValue}
                <input
                  value={criterionFormula2}
                  onChange={(event) => setCriterionFormula2(event.target.value)}
                />
              </label>
            )}
          </>
        )}
        <fieldset>
          <legend>{text.validationOptions}</legend>
          <label>
            <input
              type="checkbox"
              checked={allowBlank}
              onChange={(event) => setAllowBlank(event.target.checked)}
            />
            {text.validationAllowBlank}
          </label>
          {(kind === 'list' || kind === 'listRef') && (
            <label>
              <input
                type="checkbox"
                checked={showDropdown}
                onChange={(event) => setShowDropdown(event.target.checked)}
              />
              {text.validationShowDropdown}
            </label>
          )}
          <label>
            <input
              type="checkbox"
              checked={showInputMessage}
              onChange={(event) => setShowInputMessage(event.target.checked)}
            />
            {text.validationShowPrompt}
          </label>
          {showInputMessage && (
            <>
              <label>
                {text.validationPromptTitle}
                <input
                  value={promptTitle}
                  onChange={(event) => setPromptTitle(event.target.value)}
                />
              </label>
              <label>
                {text.validationPrompt}
                <input value={prompt} onChange={(event) => setPrompt(event.target.value)} />
              </label>
            </>
          )}
          <label>
            <input
              type="checkbox"
              checked={showErrorMessage}
              onChange={(event) => setShowErrorMessage(event.target.checked)}
            />
            {text.validationShowError}
          </label>
          {showErrorMessage && (
            <>
              <label>
                {text.validationErrorStyle}
                <select
                  value={errorStyle}
                  onChange={(event) => setErrorStyle(event.target.value as typeof errorStyle)}
                >
                  <option value="stop">{text.validationStop}</option>
                  <option value="warning">{text.validationWarning}</option>
                  <option value="information">{text.validationInformation}</option>
                </select>
              </label>
              <label>
                {text.validationErrorTitle}
                <input value={errorTitle} onChange={(event) => setErrorTitle(event.target.value)} />
              </label>
              <label>
                {text.validationError}
                <input value={error} onChange={(event) => setError(event.target.value)} />
              </label>
            </>
          )}
        </fieldset>
        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, marginTop: 18 }}>
          <button type="button" onClick={onClose}>
            {text.cancel}
          </button>
          <button type="button" onClick={apply}>
            {text.apply}
          </button>
        </div>
      </section>
    </div>
  )
}
