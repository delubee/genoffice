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
import { formatAddress } from '@genoffice/xlsx-gateway/domain/cell-address'
import { FindReplacePanel } from '../../../apps/sheets/src/renderer/FindReplacePanel'
import { useState } from 'react'
import { normalizeLang } from '@genoffice/i18n'
import { InsertFunctionFlow } from '../../../apps/sheets/src/renderer/InsertFunctionFlow'
import { FunctionLibraryRibbon, SheetsRibbon } from '../../../apps/sheets/src/renderer/SheetsRibbon'
import { GoToDialog } from '../../../apps/sheets/src/renderer/GoToDialog'
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
}: {
  onPivotCreate?(): void
  onPivotEdit?(): void
  runtime: WorkbookRuntime | null
  locale?: string
  onError(message: string): void
  activeTab?:
    | 'Home'
    | 'Insert'
    | 'Page Layout'
    | 'Formulas'
    | 'Chart Design'
    | 'Shape Format'
    | 'Picture Format'
  ribbonContainer?: HTMLElement | null
}) {
  const lang = normalizeLang(locale ?? navigator.language)
  setModuleLang(lang)
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
    if (!runtime) {
      onError('The document is still loading.')
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
          formatPainterActive={false}
          sheetHasContent={runtime?.ribbonEcho.hasContent ?? false}
          sheetProtected={runtime?.ribbonEcho.protected ?? null}
          workbookProtected={null}
          formulaBarVisible={true}
          crossHighlightVisible={false}
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
          selectedTable={null}
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
