import { WatchWindowPanel } from '../../../apps/sheets/src/renderer/WatchWindowPanel'
import { NameManagerDialog } from '../../../apps/sheets/src/renderer/NameManagerDialog'
import { PasteSpecialDialog } from '../../../apps/sheets/src/renderer/PasteSpecialDialog'
import {
  PASTE_SPECIAL_OPEN_SHELL_COMMAND,
  type PasteSpecialType,
} from '../../../apps/sheets/src/renderer/paste-special-model'
import { RecommendedChartsDialog } from '../../../apps/sheets/src/renderer/RecommendedChartsDialog'
import type {
  ChartRecommendations,
  RecommendedKind,
} from '@genoffice/xlsx-gateway/domain/chart-recommend'
import {
  HeaderFooterDialog,
  type HeaderFooterResult,
} from '../../../apps/sheets/src/renderer/HeaderFooterDialog'
import { SymbolDialog } from '../../../apps/sheets/src/renderer/SymbolDialog'
import { LinkDialog } from '../../../apps/sheets/src/renderer/LinkDialog'
import { createPortal } from 'react-dom'
import { formatAddress, parseRange } from '@genoffice/xlsx-gateway/domain/cell-address'
import { FindReplacePanel } from '../../../apps/sheets/src/renderer/FindReplacePanel'
import { useEffect, useRef, useState } from 'react'
import { resolveGlobalShortcut } from '../../../apps/sheets/src/renderer/global-shortcuts'
import { normalizeLang } from '@genoffice/i18n'
import { InsertFunctionFlow } from '../../../apps/sheets/src/renderer/InsertFunctionFlow'
import { FunctionLibraryRibbon, SheetsRibbon } from '../../../apps/sheets/src/renderer/SheetsRibbon'
import { GoToDialog } from '../../../apps/sheets/src/renderer/GoToDialog'
import { TableRangeDialog } from '../../../apps/sheets/src/renderer/TableRangeDialog'
import { IconsDialog } from '../../../apps/sheets/src/renderer/IconsDialog'
import { InsertDeleteCellsDialog } from '../../../apps/sheets/src/renderer/InsertDeleteCellsDialog'
import { FillSeriesDialog } from '../../../apps/sheets/src/renderer/FillSeriesDialog'
import { fillSeriesDefaultType } from '../../../apps/sheets/src/renderer/fill-series-apply'
import { EquationDialog } from '../../../apps/sheets/src/renderer/EquationDialog'
import {
  ThreadHoverCard,
  ThreadedCommentsPane,
} from '../../../apps/sheets/src/renderer/ThreadedCommentsPane'
import { ProtectSheetDialog } from '../../../apps/sheets/src/renderer/ProtectSheetDialog'
import { AllowEditRangesDialog } from '../../../apps/sheets/src/renderer/AllowEditRangesDialog'
import { AxisSizeDialog } from '../../../apps/sheets/src/renderer/AxisSizeDialog'
import { FormatCellsDialog } from '../../../apps/sheets/src/renderer/FormatCellsDialog'
import { LocaleScope, setModuleLang, t } from '../../../apps/sheets/src/renderer/i18n/locale'
import type { WorkbookRuntime } from './runtime'
import type { RibbonTarget } from './ribbon-intents'

/** Desktop Home/Insert/Chart Design command bands backed by the plugin's authoritative session. */
export function DesktopHomeRibbon({
  runtime,
  locale,
  onError,
  activeTab = 'Home',
  ribbonContainer,
  onPivotCreate,
  onPivotEdit,
  onPivotFilter,
  onSharedCommand,
}: {
  /** Data commands the desktop also dispatches from Home; true when handled. */
  onSharedCommand?(value: string): boolean
  onPivotCreate?(): void
  onPivotFilter?(kind: 'slicer' | 'timeline'): void
  onPivotEdit?(): void
  runtime: WorkbookRuntime | null
  locale?: string
  onError(message: string): void
  activeTab?:
    | 'Home'
    | 'Insert'
    | 'Page Layout'
    | 'Formulas'
    | 'Review'
    | 'View'
    | 'Table Design'
    | 'Chart Design'
    | 'Shape Format'
    | 'Picture Format'
  ribbonContainer?: HTMLElement | null
}) {
  const lang = normalizeLang(locale ?? navigator.language)
  setModuleLang(lang)
  const [viewNotice, setViewNotice] = useState<string | null>(null)
  const [rendered, setRendered] = useState<'icons' | 'equation' | null>(null)
  const [tableDialog, setTableDialog] = useState<'create' | 'resize' | null>(null)
  const [cellsDialog, setCellsDialog] = useState<'insert' | 'delete' | null>(null)
  const [fillSeries, setFillSeries] = useState<
    NonNullable<typeof runtime>['fillSeriesContext'] | null
  >(null)
  const [protection, setProtection] = useState<{
    mode: 'protect' | 'unprotect' | 'ranges'
    context: Awaited<ReturnType<NonNullable<typeof runtime>['protectionContext']>>
  } | null>(null)
  const [nameNotice, setNameNotice] = useState<{ count: number; skipped: number } | null>(null)
  const [dialog, setDialog] = useState<{
    kind: 'format' | 'row' | 'col' | 'link' | 'symbol' | 'header-footer' | 'recommended-charts'
    recommendations?: ChartRecommendations
    headerFooter?: HeaderFooterResult
    currentLink?: string | null
    target: RibbonTarget
    stateId: string
    revision: number
  } | null>(null)
  const rawAnchor = dialog
    ? runtime?.projectedValues(dialog.target.address.split(':')[0]!)?.[0]?.[0]
    : null
  const anchorValue =
    typeof rawAnchor === 'number' || typeof rawAnchor === 'string' ? rawAnchor : null
  const run = async (work: () => Promise<unknown>) => {
    onError('')
    try {
      await work()
      return true
    } catch (error) {
      onError(String(error))
      return false
    }
  }
  const command = (value: string) => {
    setNameNotice(null)
    setViewNotice(null)
    if (!runtime) {
      onError('The document is still loading.')
      return
    }
    if (onSharedCommand?.(value)) return
    if (
      value === 'toggle-formula-bar' ||
      value === 'toggle-cross-highlight' ||
      value === 'toggle-page-break-preview' ||
      value === 'toggle-right-to-left' ||
      value.startsWith('zoom:') ||
      value === 'zoom-reset' ||
      value === 'zoom-to-selection' ||
      value.startsWith('freeze-') ||
      value === 'unfreeze'
    ) {
      void run(async () => setViewNotice(await runtime.viewCommand(value)))
      return
    }
    if (
      value === 'note-open' ||
      value === 'note-show-toggle' ||
      value === 'note-show-all' ||
      value === 'note-delete'
    ) {
      void run(async () => setViewNotice(await runtime.noteCommand(value)))
      return
    }
    if (
      value === 'comment-new' ||
      value === 'comment-reply' ||
      value === 'comment-resolve' ||
      value === 'comment-delete' ||
      value === 'comment-prev' ||
      value === 'comment-next' ||
      value === 'comments-pane-toggle' ||
      value === 'notes-convert'
    ) {
      void run(async () => setViewNotice(await runtime.commentCommand(value)))
      return
    }
    if (value === 'insert-icons' || value === 'insert-equation') {
      setRendered(value === 'insert-icons' ? 'icons' : 'equation')
      return
    }
    if (value === 'insert-screenshot') {
      onError('Screen capture is not available in VS Code; insert the image with Pictures instead.')
      return
    }
    if (value === 'format-as-table' || value.startsWith('format-as-table:')) {
      void run(async () =>
        setViewNotice(await runtime.formatAsTable(value.split(':')[1] || undefined)),
      )
      return
    }
    if (value === 'table-create-open' || value === 'table-resize-open') {
      setTableDialog(value === 'table-create-open' ? 'create' : 'resize')
      return
    }
    if (
      value.startsWith('table-style:') ||
      value.startsWith('table-option:') ||
      value.startsWith('table-rename:') ||
      value.startsWith('table-resize:') ||
      value === 'table-convert-range'
    ) {
      void run(async () => setViewNotice(await runtime.tableCommand(value)))
      return
    }
    if (value.startsWith('insert-pivot-chart:')) {
      void run(async () =>
        setViewNotice(await runtime.insertPivotChart(value.slice('insert-pivot-chart:'.length))),
      )
      return
    }
    if (value.startsWith('sparkline:')) {
      const type = value.slice('sparkline:'.length)
      if (type === 'line' || type === 'column' || type === 'stacked')
        void run(async () => setViewNotice(await runtime.insertSparklines(type)))
      return
    }
    if (value === 'cf-open') {
      void run(() => runtime.openConditionalFormatting())
      return
    }
    if (value === 'fill-series-open') {
      void run(async () => {
        if (!(await runtime.flushEditor()))
          throw new Error('Finish editing before filling a series.')
        setFillSeries(runtime.fillSeriesContext)
      })
      return
    }
    if (value === 'insert-cells' || value === 'delete-cells') {
      void run(async () => {
        const mode = value === 'insert-cells' ? 'insert' : 'delete'
        const direct = runtime.cellsDecision()
        if (direct) await runtime.cellsCommand(mode, direct)
        else setCellsDialog(mode)
      })
      return
    }
    if (value === 'insert-sheet') {
      void run(async () => setViewNotice(await runtime.insertSheet()))
      return
    }
    if (value === 'insert-now:date' || value === 'insert-now:time') {
      void run(() => runtime.insertNow(value === 'insert-now:date' ? 'date' : 'time'))
      return
    }
    if (/^page-(row|col):-?1$/.test(value)) {
      runtime.pageMove(value.startsWith('page-col:'), value.endsWith(':-1') ? -1 : 1)
      return
    }
    if (value === 'outline-group:auto' || value === 'outline-ungroup:auto') {
      // Shift+Alt+Arrow: whole columns selected mean columns, anything else rows.
      const target = runtime.ribbonTarget
      const bounds = target ? parseRange(target.address) : null
      const dimensions = target?.dimensions
      const columns =
        bounds &&
        dimensions &&
        bounds.endRow - bounds.startRow + 1 >= dimensions.rows &&
        bounds.endColumn - bounds.startColumn + 1 < dimensions.columns
      void run(() =>
        runtime.outlineCommand(
          value === 'outline-group:auto' ? 'group' : 'ungroup',
          columns ? 'cols' : 'rows',
        ),
      )
      return
    }
    if (value === 'formula-bar-toggle') {
      // Univer keeps the expanded state in component state; the fold arrow is the only way in.
      document
        .querySelector<HTMLElement>('[data-u-comp="formula-bar"] .univer-w-5.univer-cursor-pointer')
        ?.click()
      return
    }
    if (value === 'format-painter' || value === 'format-painter:dblclick') {
      void run(() => runtime.formatPainter(value === 'format-painter' ? 1 : 2))
      return
    }
    if (value === 'insert-checkbox') {
      void run(() => runtime.insertCheckbox())
      return
    }
    if (value === 'workbook-statistics') {
      void run(async () => setViewNotice(await runtime.workbookStatistics()))
      return
    }
    if (value === 'workbook-protect') {
      void run(async () => {
        const locked = !runtime.workbookProtected
        const failure = await runtime.protectionChange({ op: 'protect_workbook', locked })
        if (failure) throw new Error(failure)
        setViewNotice(
          t(locked ? 'appWorkbookProtectionWillWrite' : 'appWorkbookProtectionWillRemove'),
        )
      })
      return
    }
    if (value === 'sheet-protect' || value === 'allow-edit-ranges-open') {
      void run(async () => {
        const context = await runtime.protectionContext()
        if (value === 'allow-edit-ranges-open') {
          setProtection({ mode: 'ranges', context })
          return
        }
        if (!context.protection) throw new Error(t('appProtectionNeedsIndexed'))
        if (!context.protection.protected) setProtection({ mode: 'protect', context })
        else if (context.protection.hasPassword) setProtection({ mode: 'unprotect', context })
        else {
          const failure = await runtime.protectionChange(
            { op: 'protect_sheet', sheetId: context.sheetId, protected: false },
            context.frozen,
          )
          if (failure) throw new Error(failure)
          setViewNotice(t('appProtectionWillRemove'))
        }
      })
      return
    }
    if (value === 'slicer-open' || value === 'timeline-open') {
      onPivotFilter?.(value === 'slicer-open' ? 'slicer' : 'timeline')
      return
    }
    if (value === 'pivot-open' || value === 'pivot-edit') {
      if (value === 'pivot-open') onPivotCreate?.()
      else onPivotEdit?.()
      return
    }
    if (value === 'calc-mode:auto' || value === 'calc-mode:manual') {
      runtime.setCalculationManual(value === 'calc-mode:manual')
      return
    }
    if (value === 'calculate-now') {
      runtime.recalculate()
      return
    }
    if (value === 'calculate-sheet') {
      runtime.recalculateSheet()
      return
    }
    if (value === 'watch-window') {
      runtime.toggleWatchWindow()
      return
    }
    if (
      [
        'trace-precedents',
        'trace-dependents',
        'remove-arrows',
        'toggle-show-formulas',
        'error-checking',
      ].includes(value)
    ) {
      void run(() => runtime.auditCommand(value))
      return
    }
    if (value === 'name-manager-open') {
      void run(() => runtime.openNameManager())
      return
    }
    if (value.startsWith('use-in-formula:')) {
      void run(() => runtime.useNameInFormula(value.slice('use-in-formula:'.length)))
      return
    }
    if (value === 'create-names:top' || value === 'create-names:left') {
      void run(async () => {
        setNameNotice(
          await runtime.createNamesFromSelection(value === 'create-names:top' ? 'top' : 'left'),
        )
      })
      return
    }
    if (value === 'insert-function-open' || value.startsWith('insert-function-open:')) {
      void run(() => runtime.openInsertFunction(value.split(':')[1] ?? 'All'))
      return
    }
    if (value.startsWith('autofn:')) {
      void run(() => runtime.insertAutoFunction(value.slice('autofn:'.length)))
      return
    }
    if (value === 'copy' || value === 'cut') {
      void run(() => runtime.copyClipboard(value === 'cut' ? 'CUT' : 'COPY'))
      return
    }
    if (value === 'paste') {
      void run(() => runtime.pasteClipboard())
      return
    }
    if (value === PASTE_SPECIAL_OPEN_SHELL_COMMAND) {
      void run(() => runtime.openPasteSpecial())
      return
    }
    if (value.startsWith('paste-special:')) {
      const types: Record<string, PasteSpecialType> = {
        value: 'values',
        formula: 'formulas',
        format: 'formats',
        'col-width': 'col-widths',
        'besides-border': 'all-except-borders',
      }
      const type = types[value.slice('paste-special:'.length)]
      if (type) void run(() => runtime.pasteSpecial(type))
      return
    }
    if (
      value.startsWith('page-layout:') ||
      value === 'toggle-gridlines' ||
      value === 'toggle-headings'
    ) {
      void run(() => runtime.pageLayoutCommand(value))
      return
    }
    if (value === 'insert-picture') {
      void run(() => runtime.insertPicture())
      return
    }
    if (value.startsWith('insert-shape:') || value === 'insert-textbox') {
      void run(() =>
        runtime.startShapeDrawing(
          value === 'insert-textbox' ? 'textbox' : value.slice('insert-shape:'.length),
        ),
      )
      return
    }
    if (value.startsWith('visual:') || value.startsWith('shape-')) {
      void run(() => runtime.shapeRibbonCommand(value))
      return
    }
    if (value.startsWith('chart-')) {
      void run(() => runtime.chartRibbonCommand(value))
      return
    }
    if (value === 'find' || value === 'replace' || value === 'goto-open') {
      void run(() => runtime.openViewCommand(value === 'goto-open' ? 'goto' : value))
      return
    }
    if (
      value === 'recommended-charts-open' ||
      /^insert-chart:(column|bar|line|area|pie|scatter|radar|doughnut|combo)$/.test(value)
    ) {
      void run(async () => {
        if (!(await runtime.flushEditor()))
          throw new Error('Finish editing before inserting a chart.')
        const target = runtime.ribbonTarget
        if (!target) throw new Error('Select data before inserting a chart.')
        const { stateId, revision } = runtime.selectionContext
        const frozen = { target, stateId, revision }
        if (value === 'recommended-charts-open') {
          const recommendations = await runtime.recommendCharts(frozen)
          setDialog({ ...frozen, kind: 'recommended-charts', recommendations })
        } else
          await runtime.insertChart(value.slice('insert-chart:'.length) as RecommendedKind, frozen)
      })
      return
    }
    if (value === 'header-footer-open') {
      void run(async () => {
        if (!(await runtime.flushEditor()))
          throw new Error('Finish editing before changing page settings.')
        const target = runtime.ribbonTarget
        if (!target) throw new Error('Select a sheet before changing page settings.')
        const { stateId, revision } = runtime.selectionContext
        const headerFooter = await runtime.readHeaderFooter(target, stateId)
        if (
          runtime.selectionContext.stateId !== stateId ||
          runtime.selectionContext.revision !== revision
        )
          throw new Error('The document changed. Open the dialog again.')
        setDialog({ kind: 'header-footer', target, stateId, revision, headerFooter })
      })
      return
    }
    if (value === 'insert-symbol') {
      void run(async () => {
        if (!(await runtime.flushEditor()))
          throw new Error('Finish editing before inserting symbols.')
        const target = runtime.ribbonTarget
        if (!target) throw new Error('Select a cell before inserting symbols.')
        const { stateId, revision } = runtime.selectionContext
        setDialog({ kind: 'symbol', target, stateId, revision })
      })
      return
    }
    if (value === 'link-open') {
      void run(async () => {
        if (!(await runtime.flushEditor()))
          throw new Error('Finish editing before inserting a link.')
        const target = runtime.ribbonTarget
        if (!target) throw new Error('Select a cell before inserting a link.')
        const { stateId, revision } = runtime.selectionContext
        const currentLink = await runtime.readHyperlink(target, stateId)
        if (runtime.selectionContext.stateId !== stateId)
          throw new Error('The document changed. Open the dialog again.')
        setDialog({ kind: 'link', target, stateId, revision, currentLink })
      })
      return
    }
    if (value === 'format-cells' || value === 'row-height-open' || value === 'col-width-open') {
      void run(async () => {
        if (!(await runtime.flushEditor()))
          throw new Error('Finish editing before opening Format Cells.')
        const target = runtime.ribbonTarget
        if (!target) throw new Error('Select cells before opening Format Cells.')
        const { stateId, revision } = runtime.selectionContext
        setDialog({
          kind: value === 'format-cells' ? 'format' : value === 'row-height-open' ? 'row' : 'col',
          target,
          stateId,
          revision,
        })
      })
      return
    }
    void run(() => runtime.ribbonCommands([value]))
  }
  // The desktop window shortcut table (global-shortcuts.ts). Capture keeps
  // VS Code's own bindings for these chords from firing inside the editor.
  const commandRef = useRef(command)
  commandRef.current = command
  useEffect(() => {
    if (!runtime) return
    runtime.commandSink = (value) => commandRef.current(value)
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.defaultPrevented) return
      const guards = runtime.shortcutGuards(event)
      const action = guards && resolveGlobalShortcut(event, guards)
      if (!action) return
      event.preventDefault()
      event.stopImmediatePropagation()
      commandRef.current(
        action.kind === 'command'
          ? action.command
          : action.dialog === 'formatCells'
            ? 'format-cells'
            : 'goto-open',
      )
    }
    window.addEventListener('keydown', onKeyDown, true)
    return () => {
      window.removeEventListener('keydown', onKeyDown, true)
      runtime.commandSink = undefined
    }
  }, [runtime])
  const unused = (): never => {
    throw new Error('A non-Home ribbon callback was invoked in the Home adapter.')
  }
  const header = (
    <header className="excel-header workbook-desktop-header" inert={!runtime}>
      {runtime?.auditNotice && <div role="status">{runtime.auditNotice}</div>}

      {activeTab === 'Formulas' ? (
        <FunctionLibraryRibbon
          auditing
          onCommand={command}
          definedNames={runtime?.goToNames.map((name) => name.name) ?? []}
          calcManual={runtime?.calculationManual ?? false}
        />
      ) : (
        <SheetsRibbon
          activeTab={activeTab}
          selectionFormat={runtime?.ribbonTarget?.format ?? null}
          formatPainterActive={runtime?.formatPainterActive ?? false}
          sheetHasContent={runtime?.ribbonEcho.hasContent ?? false}
          sheetProtected={runtime?.ribbonEcho.protected ?? null}
          workbookProtected={runtime?.workbookProtected ?? null}
          formulaBarVisible={runtime?.formulaBarVisible ?? true}
          crossHighlightVisible={runtime?.crossHighlightVisible ?? false}
          themeEditingAvailable={false}
          pageLayout={
            runtime?.pageLayoutEcho ?? {
              showGridlines: true,
              showHeadings: true,
              rightToLeft: false,
            }
          }
          selectedChart={runtime?.chartRibbonTarget?.echo ?? null}
          selectedShape={
            runtime?.shapeRibbonTarget?.kind === 'shape'
              ? {
                  fillColor: runtime.shapeRibbonTarget.fillColor,
                  lineColor: runtime.shapeRibbonTarget.lineColor,
                }
              : null
          }
          selectedTable={runtime?.selectedTable ?? null}
          outlineSummary={{ below: true, right: true }}
          onCommand={command}
          onListNames={unused}
          calcManual={false}
          onAiRun={(instruction) => {
            if (runtime) void run(() => runtime.ribbonChat(instruction))
          }}
          aiOpen={false}
          onAiToggle={() => {
            if (runtime) void run(() => runtime.ribbonChat())
          }}
          onRefreshPivot={() => {
            if (runtime) void run(() => runtime.refreshPivots('sheet'))
            return null
          }}
          onIsSelectionInPivot={() => runtime?.selectionInPivot ?? false}
        />
      )}
      {viewNotice && (
        <p className="dialog-note" role="status">
          {viewNotice}
        </p>
      )}
      {nameNotice && (
        <p className="dialog-note" role="status">
          {t(nameNotice.skipped ? 'appNamesCreatedSkipped' : 'appNamesCreated', nameNotice)}
        </p>
      )}
    </header>
  )
  return (
    <LocaleScope lang={lang}>
      {ribbonContainer === undefined
        ? header
        : ribbonContainer
          ? createPortal(header, ribbonContainer)
          : null}
      {createPortal(
        <>
          {runtime?.watchOpen && (
            <WatchWindowPanel
              watches={runtime.watches}
              onResolve={(cell) => runtime.resolveWatch(cell)}
              onRefresh={runtime.refreshWatches}
              onAddSelection={() => {
                void run(() => runtime.addWatchSelection())
              }}
              onRemove={(key) => runtime.removeWatch(key)}
              onClose={() => runtime.toggleWatchWindow()}
            />
          )}
          {runtime?.nameManagerContext && (
            <NameManagerDialog
              names={runtime.nameManagerRows}
              sheets={runtime.nameManagerSheets}
              onAction={(action) => runtime.applyNameAction(action)}
              onClose={() => runtime.closeNameManager()}
            />
          )}
          {runtime?.functionDialog && (
            <InsertFunctionFlow
              functions={runtime.functionDialog.functions}
              target={runtime.functionDialog.target}
              initialCategory={runtime.functionDialog.category}
              host={runtime.functionDialog.host}
              onClose={() => runtime.closeInsertFunction()}
            />
          )}
          {dialog?.kind === 'recommended-charts' && dialog.recommendations && runtime && (
            <RecommendedChartsDialog
              recommendations={dialog.recommendations}
              onClose={() => setDialog(null)}
              onPick={async (kind) => {
                await runtime.insertChart(kind, dialog)
                return true
              }}
            />
          )}
          {dialog?.kind === 'header-footer' && runtime && (
            <HeaderFooterDialog
              initialHeader={dialog.headerFooter?.header ?? null}
              initialFooter={dialog.headerFooter?.footer ?? null}
              onClose={() => setDialog(null)}
              onApply={async (result) => {
                try {
                  await runtime.applyHeaderFooter(result, dialog)
                  return null
                } catch (error) {
                  return String(error)
                }
              }}
            />
          )}
          {runtime?.pasteSpecialDialog && (
            <PasteSpecialDialog
              source={runtime.pasteSpecialDialog.snapshot.source}
              onClose={() => runtime.closePasteSpecial()}
              onApply={async (options) => {
                try {
                  await runtime.applySpecialPaste(runtime.pasteSpecialDialog!, options)
                  return null
                } catch (error) {
                  return String(error)
                }
              }}
            />
          )}
          {dialog?.kind === 'symbol' && runtime && (
            <SymbolDialog
              onClose={() => setDialog(null)}
              onInsert={async (symbol) => {
                await run(async () => {
                  const committed = await runtime.insertSymbol(symbol, dialog)
                  setDialog((current) =>
                    current === dialog ? { ...current, ...committed } : current,
                  )
                })
              }}
            />
          )}
          {dialog?.kind === 'link' && runtime && (
            <LinkDialog
              currentTarget={dialog.currentLink ?? null}
              onCommand={command}
              onClose={() => setDialog(null)}
              onApply={(value) => run(() => runtime.ribbonCommands([value], dialog))}
            />
          )}
          {dialog?.kind === 'format' && runtime && (
            <FormatCellsDialog
              selectionFormat={dialog.target.format}
              anchorValue={anchorValue}
              onCommand={command}
              onClose={() => setDialog(null)}
              onApply={(commands) => run(() => runtime.ribbonCommands(commands, dialog))}
            />
          )}
          {dialog && (dialog.kind === 'row' || dialog.kind === 'col') && runtime && (
            <AxisSizeDialog
              axis={dialog.kind}
              onCommand={command}
              onClose={() => setDialog(null)}
              onApply={(command) => run(() => runtime.ribbonCommands([command], dialog))}
            />
          )}
          {runtime && (
            <FindReplacePanel
              service={runtime.findReplace}
              getWorkbook={() => runtime.findWorkbook}
              onJumpTo={(sheetId, bounds) => {
                void run(async () => {
                  const sheet = runtime.findWorkbook?.getSheetBySheetId(sheetId)
                  if (!sheet) throw new Error('The search result sheet no longer exists.')
                  const error = await runtime.goToReference(
                    "'" +
                      sheet.getName().replace(/'/g, "''") +
                      "'!" +
                      formatAddress(bounds.startRow, bounds.startColumn),
                  )
                  if (error) throw new Error(error)
                })
              }}
              onClose={() => runtime.findReplace.terminate()}
              registerContainer={() => () => {}}
            />
          )}
          {runtime && protection && protection.mode !== 'ranges' && (
            <ProtectSheetDialog
              mode={protection.mode}
              onProtect={async ({ password, allow }) => {
                const failure = await runtime.protectionChange(
                  {
                    op: 'protect_sheet',
                    sheetId: protection.context.sheetId,
                    protected: true,
                    allow,
                    ...(password ? { password } : {}),
                  },
                  protection.context.frozen,
                )
                if (failure === null) setViewNotice(t('appProtectionWillWrite'))
                return failure
              }}
              onUnprotect={async (password) => {
                const failure = await runtime.protectionChange(
                  {
                    op: 'protect_sheet',
                    sheetId: protection.context.sheetId,
                    protected: false,
                    password,
                  },
                  protection.context.frozen,
                )
                if (failure === null) setViewNotice(t('appProtectionWillRemove'))
                return failure
              }}
              onClose={() => setProtection(null)}
            />
          )}
          {runtime && protection?.mode === 'ranges' && (
            <AllowEditRangesDialog
              ranges={protection.context.ranges}
              defaultRef={protection.context.activeCell}
              onApply={(ranges) => {
                if (protection.context.ranges.some((range) => range.hasPassword))
                  return t('appRangesPasswordBlocked')
                void run(async () => {
                  const failure = await runtime.protectionChange(
                    {
                      op: 'set_protected_ranges',
                      sheetId: protection.context.sheetId,
                      ranges: ranges.map(({ name, sqref }) => ({ name, sqref })),
                    },
                    protection.context.frozen,
                  )
                  if (failure) throw new Error(failure)
                  setViewNotice(t('appRangesRecorded', { count: ranges.length }))
                })
                return null
              }}
              onClose={() => setProtection(null)}
            />
          )}
          {runtime && <ThreadedCommentsPane getRuntime={() => runtime.univerRuntime as never} />}
          <ThreadHoverCard hover={runtime?.threadHover ?? null} />
          {fillSeries && runtime && (
            <FillSeriesDialog
              context={fillSeries}
              defaultType={fillSeriesDefaultType(fillSeries.numberFormat)}
              onApply={(options) => {
                void run(async () => {
                  const failure = await runtime.fillSeries(options)
                  if (failure) throw new Error(failure)
                })
                return null
              }}
              onClose={() => setFillSeries(null)}
            />
          )}
          {cellsDialog && runtime && (
            <InsertDeleteCellsDialog
              mode={cellsDialog}
              onApply={(choice) => void run(() => runtime.cellsCommand(cellsDialog, choice))}
              onClose={() => setCellsDialog(null)}
            />
          )}
          {runtime && tableDialog === 'create' && (
            <TableRangeDialog
              title={t('appCreateTable')}
              note={t('appCreateTableWhere')}
              initialRange={runtime.ribbonTarget?.address ?? ''}
              headersOption
              onSubmit={(range, headers) =>
                void run(async () => setViewNotice(await runtime.createTable(range, headers)))
              }
              onClose={() => setTableDialog(null)}
            />
          )}
          {runtime && tableDialog === 'resize' && runtime.selectedTable && (
            <TableRangeDialog
              title={t('appResizeTable')}
              note={t('appResizeTableNote')}
              initialRange={runtime.selectedTable.rangeA1}
              onSubmit={(range) =>
                void run(async () =>
                  setViewNotice(await runtime.tableCommand(`table-resize:${range}`)),
                )
              }
              onClose={() => setTableDialog(null)}
            />
          )}
          {runtime && rendered === 'icons' && (
            <IconsDialog
              onInsert={(dataUrl, _size, name) =>
                void run(() =>
                  runtime.insertRenderedPicture(dataUrl, `${name.replace(/\s+/g, '-')}.png`),
                )
              }
              onClose={() => setRendered(null)}
            />
          )}
          {runtime && rendered === 'equation' && (
            <EquationDialog
              onInsert={(dataUrl) =>
                void run(() => runtime.insertRenderedPicture(dataUrl, 'Equation.png'))
              }
              onClose={() => setRendered(null)}
            />
          )}
          {runtime?.goToOpen && (
            <GoToDialog
              names={runtime.goToNames}
              onGo={(reference) => runtime.goToReference(reference)}
              onClose={() => runtime.closeGoTo()}
            />
          )}
        </>,
        document.body,
      )}
    </LocaleScope>
  )
}
