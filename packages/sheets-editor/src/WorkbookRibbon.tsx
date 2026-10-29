import { SubtotalDialog } from '../../../apps/sheets/src/renderer/SubtotalDialog'
import {
  DataToolsRibbonGroup,
  ForecastRibbonGroup,
  GetDataRibbonGroup,
  OutlineRibbonGroup,
  SortFilterRibbonGroup,
} from '../../../apps/sheets/src/renderer/SheetsRibbon'
import { AdvancedFilterDialog } from '../../../apps/sheets/src/renderer/AdvancedFilterDialog'
import { RemoveDuplicatesDialog } from '../../../apps/sheets/src/renderer/RemoveDuplicatesDialog'
import { GoalSeekDialog } from '../../../apps/sheets/src/renderer/GoalSeekDialog'
import { tFor, LocaleScope } from '../../../apps/sheets/src/renderer/i18n/locale'
import { normalizeLang } from '@genoffice/i18n'
import { useEffect, useRef, useState, type ReactNode } from 'react'
import { formatAddress } from '@genoffice/xlsx-gateway/domain/cell-address'
import type { DvRule, WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import type { WorkbookRuntime } from './runtime'
import { editorText } from './i18n'
import { DesktopHomeRibbon } from './DesktopHomeRibbon'
import { DataValidationDialog } from './DataValidationDialog'
import { ConsolidateDialog } from '../../../apps/sheets/src/renderer/ConsolidateDialog'
import { SortDialog } from '../../../apps/sheets/src/renderer/SortDialog'
import type { ConsolidateConfig } from '../../../apps/sheets/src/renderer/consolidate'

const tabs = ['home', 'insert', 'pageLayout', 'formulas', 'data', 'review', 'view'] as const
type Tab = (typeof tabs)[number] | 'chartDesign' | 'shapeFormat' | 'pictureFormat'
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
  const hasChart = !!runtime?.chartRibbonTarget
  const shapeKind = runtime?.shapeRibbonTarget?.kind
  const visibleTabs: Tab[] = [
    ...tabs,
    ...(hasChart ? ['chartDesign' as const] : []),
    ...(shapeKind === 'shape'
      ? ['shapeFormat' as const]
      : shapeKind === 'image'
        ? ['pictureFormat' as const]
        : []),
  ]
  const [homeAnchor, setHomeAnchor] = useState<HTMLDivElement | null>(null)
  const [formulaAnchor, setFormulaAnchor] = useState<HTMLDivElement | null>(null)
  const [insertAnchor, setInsertAnchor] = useState<HTMLDivElement | null>(null)
  const [tab, setTab] = useState<Tab>('home'),
    [formula, setFormula] = useState('=SUM(A1:A10)')
  const [find, setFind] = useState(''),
    [replace, setReplace] = useState('')
  const [subtotalContext, setSubtotalContext] = useState<Awaited<
    ReturnType<WorkbookRuntime['subtotalContext']>
  > | null>(null)
  const [goalSeekContext, setGoalSeekContext] = useState<WorkbookRuntime['goalSeekContext']>(null)
  const [dedupeContext, setDedupeContext] = useState<Awaited<
    ReturnType<WorkbookRuntime['removeDuplicatesContext']>
  > | null>(null)
  const [dedupeNotice, setDedupeNotice] = useState<number | null>(null)
  const [csvNotice, setCsvNotice] = useState<{
    rows: number
    columns: number
    cell: string
  } | null>(null)
  const [advancedFilterContext, setAdvancedFilterContext] = useState<Awaited<
    ReturnType<WorkbookRuntime['advancedFilterContext']>
  > | null>(null)
  const [dataValidationOpen, setDataValidationOpen] = useState(false)
  const [sortContext, setSortContext] = useState<Awaited<
    ReturnType<WorkbookRuntime['customSortContext']>
  > | null>(null)
  const [consolidateTarget, setConsolidateTarget] = useState<
    | (NonNullable<WorkbookRuntime['consolidateDialogContext']> & {
        defaultReference: string
        targetLabel: string
      })
    | null
  >(null)
  const [dataValidationInitial, setDataValidationInitial] = useState<DvRule | undefined>()
  const [dataValidationTarget, setDataValidationTarget] = useState<{
    target: { sheetId: string; address: string }
    stateId: string
    revision: number
  } | null>(null)
  const [noteDrafts, setNoteDrafts] = useState<Record<string, string>>({})
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
  const openDataValidation = () => {
    try {
      const selection = selected(),
        context = runtime!.selectionContext
      setDataValidationTarget({
        target: { sheetId: selection.sheetId, address: selection.address },
        stateId: context.stateId,
        revision: context.revision,
      })
      setDataValidationInitial(runtime!.dataValidationDraft(selection.sheetId, selection.address))
      setDataValidationOpen(true)
      onError('')
    } catch (error) {
      onError(String(error))
    }
  }
  const openConsolidate = () => {
    try {
      const context = runtime?.consolidateDialogContext
      if (!context) throw new Error(text.selectFirst)
      setConsolidateTarget(context)
      onError('')
    } catch (error) {
      onError(error instanceof Error ? error.message : String(error))
    }
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
  useEffect(() => {
    if (
      (!hasChart && tab === 'chartDesign') ||
      (tab === 'shapeFormat' && shapeKind !== 'shape') ||
      (tab === 'pictureFormat' && shapeKind !== 'image')
    )
      setTab('home')
  }, [hasChart, shapeKind, tab])
  const currentTab =
    (tab === 'chartDesign' && !hasChart) ||
    (tab === 'shapeFormat' && shapeKind !== 'shape') ||
    (tab === 'pictureFormat' && shapeKind !== 'image')
      ? 'home'
      : tab
  const [visualAnchor, setVisualAnchor] = useState<HTMLDivElement | null>(null)
  const [pageAnchor, setPageAnchor] = useState<HTMLDivElement | null>(null)
  const [chartAnchor, setChartAnchor] = useState<HTMLDivElement | null>(null)
  const panels: Record<Tab, ReactNode> = {
    chartDesign: <div ref={setChartAnchor} />,
    shapeFormat: <div ref={setVisualAnchor} />,
    pictureFormat: <div ref={setVisualAnchor} />,
    home: (
      <>
        <div ref={setHomeAnchor} />
        <details className="workbook-additional-controls">
          <summary>{text.cells}</summary>
          <div className="workbook-toolbar">{home}</div>
        </details>
      </>
    ),
    insert: (
      <>
        <div ref={setInsertAnchor} />
        <details className="workbook-additional-controls">
          <summary>{text.cells}</summary>
          <div className="workbook-toolbar">{insert}</div>
        </details>
      </>
    ),
    pageLayout: <div ref={setPageAnchor} />,
    formulas: (
      <>
        <div ref={setFormulaAnchor} />
        <details className="workbook-additional-controls">
          <summary>{text.insertFormula}</summary>
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
          <button onClick={() => run(() => runtime!.recalculate())}>{text.calculateNow}</button>
        </details>
      </>
    ),
    data: (
      <>
        <LocaleScope lang={normalizeLang(locale ?? navigator.language)}>
          <div className="ribbon workbook-shared-sort-filter" data-ribbon-body="">
            <GetDataRibbonGroup
              onCommand={(command) => {
                if (command === 'import-csv')
                  run(async () => {
                    setCsvNotice(null)
                    const imported = await runtime!.importCsv(
                      normalizeLang(locale ?? navigator.language),
                    )
                    if (imported) setCsvNotice(imported)
                  })
                else if (command === 'refresh-all') run(() => runtime!.refreshAll())
                else if (command === 'merge-workbooks')
                  onError('Merge Workbooks is not available in VS Code yet.')
                else onError(`Unknown Get Data command: ${command}`)
              }}
            />
            <SortFilterRibbonGroup
              onCommand={(command) => {
                if (command === 'sort-custom-open')
                  run(async () => setSortContext(await runtime!.customSortContext()))
                else if (command === 'filter-advanced')
                  run(async () => setAdvancedFilterContext(await runtime!.advancedFilterContext()))
                else if (
                  command === 'filter-toggle' ||
                  command === 'filter-clear' ||
                  command === 'filter-reapply' ||
                  command === 'sort:asc' ||
                  command === 'sort:desc'
                )
                  run(() => runtime!.sortFilterCommand(command))
                else onError(`Unknown Sort & Filter command: ${command}`)
              }}
            />
            <DataToolsRibbonGroup
              onCommand={(command) => {
                if (command === 'flash-fill') run(() => runtime!.flashFill())
                else if (command === 'remove-duplicates-open')
                  run(async () => {
                    setDedupeNotice(null)
                    setDedupeContext(await runtime!.removeDuplicatesContext())
                  })
                else if (command === 'dv-open') openDataValidation()
                else if (command === 'consolidate-open') openConsolidate()
                else if (command === 'text-to-columns:1') run(() => runtime!.textToColumns(1))
                else if (command === 'text-to-columns:2') run(() => runtime!.textToColumns(2))
                else if (command === 'text-to-columns:4') run(() => runtime!.textToColumns(4))
                else if (command === 'text-to-columns:8') run(() => runtime!.textToColumns(8))
                else onError(`Unknown Data Tools command: ${command}`)
              }}
            />
            <ForecastRibbonGroup
              onCommand={(command) => {
                if (command === 'goal-seek-open')
                  run(async () => {
                    if (!runtime || !(await runtime.flushEditor()))
                      throw new Error(text.selectFirst)
                    const context = runtime?.goalSeekContext
                    if (!context) throw new Error(text.selectFirst)
                    setGoalSeekContext(context)
                  })
                else onError(`Unknown Forecast command: ${command}`)
              }}
            />
            <OutlineRibbonGroup
              outlineSummary={runtime?.outlineSummary ?? { below: true, right: true }}
              onCommand={(command) => {
                if (command === 'subtotal-open')
                  run(async () => setSubtotalContext(await runtime!.subtotalContext()))
                else if (command === 'outline-summary-below')
                  run(() => runtime!.outlineCommand('summary-below'))
                else if (command === 'outline-summary-right')
                  run(() => runtime!.outlineCommand('summary-right'))
                else if (command === 'outline-clear') run(() => runtime!.outlineCommand('clear'))
                else {
                  const match =
                    /^outline-(group|ungroup|hide-detail|show-detail):(rows|cols)$/.exec(command)
                  if (match)
                    run(() =>
                      runtime!.outlineCommand(
                        match[1] as 'group' | 'ungroup' | 'hide-detail' | 'show-detail',
                        match[2] as 'rows' | 'cols',
                      ),
                    )
                  else onError(`Unknown Outline command: ${command}`)
                }
              }}
            />
          </div>
        </LocaleScope>
        <RibbonGroup title={text.dataTools}>
          <button onClick={() => run(() => runtime!.setDataValidation(null))}>
            {text.clearValidation}
          </button>
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
        {visibleTabs.map((key, index) => (
          <button
            key={key}
            ref={(button) => {
              buttons.current[index] = button
            }}
            id={`ribbon-tab-${key}`}
            type="button"
            role="tab"
            aria-selected={currentTab === key}
            aria-controls={`ribbon-panel-${key}`}
            tabIndex={currentTab === key ? 0 : -1}
            onClick={() => setTab(key)}
            onKeyDown={(event) => {
              let next: number
              if (event.key === 'ArrowRight') next = (index + 1) % visibleTabs.length
              else if (event.key === 'ArrowLeft')
                next = (index + visibleTabs.length - 1) % visibleTabs.length
              else if (event.key === 'Home') next = 0
              else if (event.key === 'End') next = visibleTabs.length - 1
              else return
              event.preventDefault()
              event.stopPropagation()
              setTab(visibleTabs[next]!)
              buttons.current[next]?.focus()
            }}
          >
            {key === 'shapeFormat' || key === 'pictureFormat'
              ? tFor(
                  normalizeLang(locale ?? navigator.language),
                  key === 'shapeFormat' ? 'appTabShapeFormat' : 'appTabPictureFormat',
                )
              : key === 'chartDesign'
                ? tFor(normalizeLang(locale ?? navigator.language), 'appTabChartDesign')
                : text[key]}
          </button>
        ))}
      </div>
      <DesktopHomeRibbon
        runtime={runtime}
        locale={locale}
        onError={onError}
        activeTab={
          currentTab === 'shapeFormat'
            ? 'Shape Format'
            : currentTab === 'pictureFormat'
              ? 'Picture Format'
              : currentTab === 'chartDesign'
                ? 'Chart Design'
                : currentTab === 'formulas'
                  ? 'Formulas'
                  : currentTab === 'pageLayout'
                    ? 'Page Layout'
                    : currentTab === 'insert'
                      ? 'Insert'
                      : 'Home'
        }
        ribbonContainer={
          currentTab === 'home'
            ? homeAnchor
            : currentTab === 'formulas'
              ? formulaAnchor
              : currentTab === 'pageLayout'
                ? pageAnchor
                : currentTab === 'insert'
                  ? insertAnchor
                  : currentTab === 'chartDesign'
                    ? chartAnchor
                    : currentTab === 'shapeFormat' || currentTab === 'pictureFormat'
                      ? visualAnchor
                      : null
        }
      />
      {visibleTabs.map((key) => (
        <div
          key={key}
          id={`ribbon-panel-${key}`}
          role="tabpanel"
          aria-labelledby={`ribbon-tab-${key}`}
          hidden={currentTab !== key}
          className={`workbook-ribbon-panel ${['home', 'insert', 'pageLayout', 'formulas', 'chartDesign', 'shapeFormat', 'pictureFormat'].includes(key) ? 'workbook-desktop-panel' : 'workbook-toolbar'}`}
        >
          {panels[key]}
        </div>
      ))}
      {dedupeNotice !== null && (
        <p role="status">
          {tFor(
            normalizeLang(locale ?? navigator.language),
            dedupeNotice ? 'appDuplicatesRemoved' : 'appNoDuplicates',
            { count: dedupeNotice },
          )}
        </p>
      )}
      {csvNotice && (
        <p role="status">
          {tFor(normalizeLang(locale ?? navigator.language), 'appCsvImported', csvNotice)}
        </p>
      )}
      {advancedFilterContext && runtime && (
        <AdvancedFilterDialog
          locale={locale}
          columns={advancedFilterContext.columns}
          onApply={async (criteria) => {
            await runtime.advancedFilter(criteria, advancedFilterContext)
            return null
          }}
          onClose={() => setAdvancedFilterContext(null)}
        />
      )}
      {subtotalContext && runtime && (
        <SubtotalDialog
          locale={locale ?? navigator.language}
          fields={subtotalContext.fields}
          onCreate={async (config) => {
            await runtime.subtotal(config, subtotalContext)
            return null
          }}
          onClose={() => setSubtotalContext(null)}
        />
      )}
      {dedupeContext && runtime && (
        <RemoveDuplicatesDialog
          locale={locale}
          onCommand={async (command) => {
            if (!/^remove-duplicates:[01]$/.test(command))
              throw new Error('Invalid duplicate removal settings.')
            setDedupeNotice(
              await runtime.removeDuplicates(command === 'remove-duplicates:1', dedupeContext),
            )
          }}
          onClose={() => setDedupeContext(null)}
        />
      )}
      {dataValidationOpen && runtime && dataValidationTarget && (
        <DataValidationDialog
          runtime={runtime}
          target={dataValidationTarget}
          initial={dataValidationInitial}
          locale={locale}
          onClose={() => setDataValidationOpen(false)}
          onError={onError}
        />
      )}
      {consolidateTarget && runtime && (
        <ConsolidateDialog
          locale={locale}
          defaultReference={consolidateTarget.defaultReference}
          targetLabel={consolidateTarget.targetLabel}
          onCreate={async (config: ConsolidateConfig) => {
            await runtime.consolidate(config, consolidateTarget)
            onError('')
            return null
          }}
          onClose={() => setConsolidateTarget(null)}
        />
      )}
      {goalSeekContext && runtime && (
        <GoalSeekDialog
          locale={locale}
          initialSetCell={goalSeekContext.initialSetCell}
          onSolve={(setCell, value, byCell) =>
            runtime.goalSeek(setCell, value, byCell, goalSeekContext)
          }
          onClose={() => setGoalSeekContext(null)}
        />
      )}
      {sortContext && runtime && (
        <SortDialog
          locale={locale}
          columns={sortContext.columns}
          onCommand={(command) => runtime.customSort(command, sortContext)}
          onClose={() => setSortContext(null)}
        />
      )}
    </div>
  )
}
