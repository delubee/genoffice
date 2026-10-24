import { useRef, useState, type ReactNode } from 'react'
import { columnLabel, formatAddress } from '@genoffice/xlsx-gateway/domain/cell-address'
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import type { WorkbookRuntime } from './runtime'
import { editorText } from './i18n'

const tabs = ['home', 'insert', 'pageLayout', 'formulas', 'data', 'review', 'view'] as const
type Tab = (typeof tabs)[number]
export function RibbonGroup({ title, children }: { title: string; children: ReactNode }) {
  return (
    <fieldset className="workbook-ribbon-group">
      <legend>{title}</legend>
      <div>{children}</div>
    </fieldset>
  )
}
export function WorkbookRibbon({
  home,
  insert,
  runtime,
  locale,
  onError,
}: {
  home: ReactNode
  insert: ReactNode
  runtime: WorkbookRuntime | null
  locale?: string
  onError(message: string): void
}) {
  const text = editorText(locale)
  const [tab, setTab] = useState<Tab>('home'),
    [formula, setFormula] = useState('=SUM(A1:A10)')
  const [fn, setFn] = useState<'SUM' | 'AVERAGE' | 'COUNT' | 'MIN' | 'MAX'>('SUM')
  const [hasHeader, setHasHeader] = useState(true),
    [find, setFind] = useState(''),
    [replace, setReplace] = useState('')
  const [noteDrafts, setNoteDrafts] = useState<Record<string, string>>({})
  const [filterColumn, setFilterColumn] = useState('A'),
    [filterValues, setFilterValues] = useState('')
  const buttons = useRef<(HTMLButtonElement | null)[]>([])
  const selectedNote = runtime?.selectedNote ?? ''
  const noteIdentity = () => {
    const current = runtime?.selection
    return current
      ? `${current.sheetId}:${current.range.startRow}:${current.range.startColumn}`
      : ''
  }
  const note = noteDrafts[noteIdentity()] ?? selectedNote
  const features: Record<string, unknown> = runtime?.selectedFeatures ?? {}
  const filePage = (features.filePageSetup ?? {}) as {
    orientation?: 'portrait' | 'landscape'
    paperSize?: number
    margins?: 'normal' | 'wide' | 'narrow'
  }
  const page = (features.pageSetupEdits ?? {}) as {
    orientation?: 'portrait' | 'landscape'
    paperSize?: number
    margins?: 'normal' | 'wide' | 'narrow'
    zoomScale?: number
    showGridlines?: boolean
    showHeadings?: boolean
  }
  const original = features.originalMetadata as
    { zoomScale?: number; showGridLines?: boolean; showRowColHeaders?: boolean } | undefined
  const run = (action: () => Promise<unknown> | void) => {
    onError('')
    try {
      Promise.resolve(action()).catch((error) => onError(String(error)))
    } catch (error) {
      onError(String(error))
    }
  }
  const selected = () => {
    const value = runtime?.selection
    if (!value) throw new Error(text.selectFirst)
    return value
  }
  const submit = (make: () => WorkbookOperation) => run(() => runtime!.enqueue([make()]))
  const pageSetting = (
    patch: Omit<Extract<WorkbookOperation, { op: 'set_page_setup' }>, 'op' | 'sheetId'>,
  ) => submit(() => ({ op: 'set_page_setup', sheetId: selected().sheetId, ...patch }))
  const commitNote = (remove: boolean) =>
    run(async () => {
      const s = selected(),
        key = noteIdentity()
      const value = remove ? '' : (noteDrafts[key] ?? runtime!.selectedNote)
      if (remove) setNoteDrafts((drafts) => ({ ...drafts, [key]: '' }))
      await runtime!.enqueue([
        {
          op: 'set_note',
          sheetId: s.sheetId,
          address: formatAddress(s.range.startRow, s.range.startColumn),
          text: value || null,
        },
      ])
      // Preserve newer unsaved typing, but let subsequent AI/undo updates show
      // the canonical note after this draft has been committed.
      setNoteDrafts((drafts) => {
        if (drafts[key] !== value) return drafts
        const next = { ...drafts }
        delete next[key]
        return next
      })
    })
  const panels: Record<Tab, ReactNode> = {
    home: (
      <>
        {home}
        <RibbonGroup title={text.mergeCells}>
          <button
            onClick={() =>
              submit(() => ({
                op: 'merge_cells',
                sheetId: selected().sheetId,
                range: selected().address,
              }))
            }
          >
            {text.mergeCells}
          </button>
          <button
            onClick={() =>
              submit(() => ({
                op: 'unmerge_cells',
                sheetId: selected().sheetId,
                range: selected().address,
              }))
            }
          >
            {text.unmergeCells}
          </button>
        </RibbonGroup>
      </>
    ),
    insert,
    pageLayout: (
      <>
        <RibbonGroup title={text.pageSetup}>
          <label>
            {text.orientation}
            <select
              aria-label={text.orientation}
              value={page.orientation ?? filePage.orientation ?? 'portrait'}
              onChange={(event) =>
                pageSetting({ orientation: event.target.value as 'portrait' | 'landscape' })
              }
            >
              <option value="portrait">{text.portrait}</option>
              <option value="landscape">{text.landscape}</option>
            </select>
          </label>
          <label>
            {text.paperSize}
            <select
              aria-label={text.paperSize}
              value={page.paperSize ?? filePage.paperSize ?? 9}
              onChange={(event) => pageSetting({ paperSize: Number(event.target.value) })}
            >
              <option value="9">A4</option>
              <option value="8">A3</option>
              <option value="1">Letter</option>
              <option value="11">A5</option>
            </select>
          </label>
          <label>
            {text.margins}
            <select
              aria-label={text.margins}
              value={page.margins ?? filePage.margins ?? 'normal'}
              onChange={(event) =>
                pageSetting({ margins: event.target.value as 'normal' | 'wide' | 'narrow' })
              }
            >
              <option value="normal">{text.normal}</option>
              <option value="narrow">{text.narrow}</option>
              <option value="wide">{text.wide}</option>
            </select>
          </label>
        </RibbonGroup>
        <RibbonGroup title={text.printArea}>
          <button onClick={() => pageSetting({ printArea: selected().address })}>
            {text.setPrintArea}
          </button>
          <button onClick={() => pageSetting({ printArea: null })}>{text.clearPrintArea}</button>
          <button onClick={() => pageSetting({ fitToWidth: 1, fitToHeight: 0 })}>
            {text.fitOnePageWide}
          </button>
          <button onClick={() => pageSetting({ scale: 100 })}>{text.actualSize}</button>
        </RibbonGroup>
      </>
    ),
    formulas: (
      <>
        <RibbonGroup title={text.functionLibrary}>
          <select
            aria-label={text.functionLibrary}
            value={fn}
            onChange={(event) => setFn(event.target.value as typeof fn)}
          >
            {['SUM', 'AVERAGE', 'COUNT', 'MIN', 'MAX'].map((name) => (
              <option key={name}>{name}</option>
            ))}
          </select>
          <button onClick={() => run(() => runtime!.insertSummary(fn))}>{text.autoSum}</button>
          <button onClick={() => run(() => runtime!.recalculate())}>{text.calculateNow}</button>
        </RibbonGroup>
        <RibbonGroup title={text.insertFormula}>
          <input
            aria-label={text.formulaText}
            value={formula}
            onChange={(event) => setFormula(event.target.value)}
          />
          <button
            onClick={() =>
              submit(() => {
                const s = selected()
                return {
                  op: 'set_formula',
                  sheetId: s.sheetId,
                  address: formatAddress(s.range.startRow, s.range.startColumn),
                  formula,
                }
              })
            }
          >
            {text.insertFormula}
          </button>
        </RibbonGroup>
      </>
    ),
    data: (
      <>
        <RibbonGroup title={text.sort}>
          <label>
            <input
              type="checkbox"
              checked={hasHeader}
              onChange={(event) => setHasHeader(event.target.checked)}
            />
            {text.hasHeader}
          </label>
          {(['asc', 'desc'] as const).map((order) => (
            <button
              key={order}
              onClick={() =>
                submit(() => {
                  const s = selected()
                  return {
                    op: 'sort_range',
                    sheetId: s.sheetId,
                    range: s.address,
                    byColumn: columnLabel(s.range.startColumn),
                    order,
                    hasHeader,
                  }
                })
              }
            >
              {order === 'asc' ? text.sortAscending : text.sortDescending}
            </button>
          ))}
        </RibbonGroup>
        <RibbonGroup title={text.dataTools}>
          <button
            onClick={() =>
              submit(() => ({
                op: 'clear_range',
                sheetId: selected().sheetId,
                range: selected().address,
              }))
            }
          >
            {text.clearContents}
          </button>
          <p>{text.rangeOperationsHint}</p>
        </RibbonGroup>
        <RibbonGroup title={text.filter}>
          <button
            onClick={() =>
              submit(() => ({
                op: 'set_filter',
                sheetId: selected().sheetId,
                range: selected().address,
              }))
            }
          >
            {text.createFilter}
          </button>
          <label>
            {text.filterColumn}
            <input
              aria-label={text.filterColumn}
              value={filterColumn}
              onChange={(event) => setFilterColumn(event.target.value.toUpperCase())}
            />
          </label>
          <textarea
            aria-label={text.filterValues}
            placeholder={text.filterValues}
            value={filterValues}
            onChange={(event) => setFilterValues(event.target.value)}
          />
          <button
            onClick={() =>
              submit(() => ({
                op: 'set_filter_criteria',
                sheetId: selected().sheetId,
                column: filterColumn,
                values: filterValues.split('\n'),
              }))
            }
          >
            {text.applyFilter}
          </button>
          <button
            onClick={() => submit(() => ({ op: 'clear_filter', sheetId: selected().sheetId }))}
          >
            {text.clearFilter}
          </button>
        </RibbonGroup>
      </>
    ),
    review: (
      <>
        <RibbonGroup title={text.notes}>
          <textarea
            aria-label={text.noteText}
            value={note}
            maxLength={32000}
            onChange={(event) => {
              const value = event.target.value
              setNoteDrafts((drafts) => ({ ...drafts, [noteIdentity()]: value }))
            }}
          />
          <button onClick={() => commitNote(false)}>{text.saveNote}</button>
          <button onClick={() => commitNote(true)}>{text.deleteNote}</button>
        </RibbonGroup>
        <RibbonGroup title={text.findReplace}>
          <input
            aria-label={text.find}
            placeholder={text.find}
            value={find}
            onChange={(event) => setFind(event.target.value)}
          />
          <input
            aria-label={text.replaceWith}
            placeholder={text.replaceWith}
            value={replace}
            onChange={(event) => setReplace(event.target.value)}
          />
          <button
            onClick={() =>
              submit(() => ({
                op: 'find_replace',
                sheetId: selected().sheetId,
                range: selected().address,
                find,
                replace,
              }))
            }
          >
            {text.replaceAll}
          </button>
        </RibbonGroup>
      </>
    ),
    view: (
      <>
        <RibbonGroup title={text.show}>
          <label>
            <input
              type="checkbox"
              checked={page.showGridlines ?? original?.showGridLines ?? true}
              onChange={(event) => pageSetting({ showGridlines: event.target.checked })}
            />
            {text.gridlines}
          </label>
          <label>
            <input
              type="checkbox"
              checked={page.showHeadings ?? original?.showRowColHeaders ?? true}
              onChange={(event) => pageSetting({ showHeadings: event.target.checked })}
            />
            {text.headings}
          </label>
          <label>
            {text.zoom}
            <select
              aria-label={text.zoom}
              value={page.zoomScale ?? original?.zoomScale ?? 100}
              onChange={(event) => pageSetting({ zoomScale: Number(event.target.value) })}
            >
              {[50, 75, 100, 125, 150, 200, 300, 400].map((zoom) => (
                <option key={zoom} value={zoom}>
                  {zoom}%
                </option>
              ))}
            </select>
          </label>
        </RibbonGroup>
        <RibbonGroup title={text.freezePanes}>
          <button
            onClick={() =>
              submit(() => ({ op: 'set_freeze', sheetId: selected().sheetId, rows: 1, columns: 0 }))
            }
          >
            {text.freezeTopRow}
          </button>
          <button
            onClick={() =>
              submit(() => ({ op: 'set_freeze', sheetId: selected().sheetId, rows: 0, columns: 1 }))
            }
          >
            {text.freezeFirstColumn}
          </button>
          <button
            onClick={() =>
              submit(() => ({
                op: 'set_freeze',
                sheetId: selected().sheetId,
                rows: selected().range.startRow,
                columns: selected().range.startColumn,
              }))
            }
          >
            {text.freezeSelection}
          </button>
          <button
            onClick={() =>
              submit(() => ({ op: 'set_freeze', sheetId: selected().sheetId, rows: 0, columns: 0 }))
            }
          >
            {text.unfreezePanes}
          </button>
        </RibbonGroup>
      </>
    ),
  }
  return (
    <div className="workbook-ribbon">
      <div className="workbook-ribbon-tabs" role="tablist" aria-label={text.ribbon}>
        {tabs.map((key, index) => (
          <button
            key={key}
            ref={(button) => {
              buttons.current[index] = button
            }}
            id={`ribbon-tab-${key}`}
            type="button"
            role="tab"
            aria-selected={tab === key}
            aria-controls={`ribbon-panel-${key}`}
            tabIndex={tab === key ? 0 : -1}
            onClick={() => setTab(key)}
            onKeyDown={(event) => {
              let next: number
              if (event.key === 'ArrowRight') next = (index + 1) % tabs.length
              else if (event.key === 'ArrowLeft') next = (index + tabs.length - 1) % tabs.length
              else if (event.key === 'Home') next = 0
              else if (event.key === 'End') next = tabs.length - 1
              else return
              event.preventDefault()
              setTab(tabs[next]!)
              buttons.current[next]?.focus()
            }}
          >
            {text[key]}
          </button>
        ))}
      </div>
      {tabs.map((key) => (
        <div
          key={key}
          id={`ribbon-panel-${key}`}
          role="tabpanel"
          aria-labelledby={`ribbon-tab-${key}`}
          hidden={tab !== key}
          className="workbook-ribbon-panel workbook-toolbar"
        >
          {panels[key]}
        </div>
      ))}
    </div>
  )
}
