import { z } from 'zod'
import type { SubtotalConfig } from '@genoffice/xlsx-gateway/domain/subtotal-plan'
import {
  UniverSheetsSortPreset,
  SortRangeCommand,
  type ISortRangeCommandParams,
} from '@univerjs/preset-sheets-sort'
import sortEnUS from '@univerjs/preset-sheets-sort/locales/en-US'
import sortZhCN from '@univerjs/preset-sheets-sort/locales/zh-CN'
import sortZhTW from '@univerjs/preset-sheets-sort/locales/zh-TW'
import '@univerjs/preset-sheets-sort/lib/index.css'
import { UniverSheetsNotePreset } from '@univerjs/preset-sheets-note'
import noteEnUS from '@univerjs/preset-sheets-note/locales/en-US'
import noteZhCN from '@univerjs/preset-sheets-note/locales/zh-CN'
import noteZhTW from '@univerjs/preset-sheets-note/locales/zh-TW'
import '@univerjs/preset-sheets-note/lib/index.css'
import { UniverSheetsTablePreset } from '@univerjs/preset-sheets-table'
import tableEnUS from '@univerjs/preset-sheets-table/locales/en-US'
import tableZhCN from '@univerjs/preset-sheets-table/locales/zh-CN'
import tableZhTW from '@univerjs/preset-sheets-table/locales/zh-TW'
import { UniverSheetsConditionalFormattingPreset } from '@univerjs/preset-sheets-conditional-formatting'
import cfEnUS from '@univerjs/preset-sheets-conditional-formatting/locales/en-US'
import cfZhCN from '@univerjs/preset-sheets-conditional-formatting/locales/zh-CN'
import cfZhTW from '@univerjs/preset-sheets-conditional-formatting/locales/zh-TW'
import '@univerjs/preset-sheets-conditional-formatting/lib/index.css'
import {
  AddConditionalRuleMutation,
  ConditionalFormattingRuleModel,
  DeleteConditionalRuleMutation,
  type IConditionFormattingRule,
} from '@univerjs/preset-sheets-conditional-formatting'
import {
  buildConditionalRule,
  dropShadowedPaintOnceRules,
} from '../../../apps/sheets/src/renderer/cf-rule-build'
import '@univerjs/preset-sheets-table/lib/index.css'
import {
  DEFAULT_TABLE_STYLE,
  TABLE_PALETTE_KEYS,
  builtinTablePalette,
  tableThemeJson,
} from '../../../apps/sheets/src/renderer/table-styles'
import type { SelectedTableRibbon } from '../../../apps/sheets/src/renderer/ribbon-model'
import {
  totalsFunctionFor,
  uniqueColumnNames,
} from '../../../apps/sheets/src/renderer/table-design'
import { inferContinuousRegion } from '../../../apps/sheets/src/renderer/table-actions'
import { SheetsNoteModel } from '@univerjs/sheets-note'
import { SheetsNotePopupService } from '@univerjs/sheets-note-ui'
import {
  threadPane,
  threadStore,
  type CellThread,
} from '../../../apps/sheets/src/renderer/threaded-comments'
import {
  installThreadedCommentAffordances,
  type ThreadHover,
} from '../../../apps/sheets/src/renderer/threaded-comment-marker'
import {
  UniverSheetsFilterPreset,
  SheetsFilterService,
  OpenFilterPanelOperation,
  CloseFilterPanelOperation,
  ChangeFilterByOperation,
  FilterBy,
  SetSheetsFilterCriteriaCommand,
  SetSheetsFilterRangeMutation,
  type IOpenFilterPanelOperationParams,
  type ISetSheetsFilterCriteriaCommandParams,
} from '@univerjs/preset-sheets-filter'
import filterEnUS from '@univerjs/preset-sheets-filter/locales/en-US'
import filterZhCN from '@univerjs/preset-sheets-filter/locales/zh-CN'
import filterZhTW from '@univerjs/preset-sheets-filter/locales/zh-TW'
import '@univerjs/preset-sheets-filter/lib/index.css'
import { filterColumnToUniver, filterPopupOperation } from './filter-popup'
import { smartFilterRange } from './filter-region'
import type { AdvancedFilterCriteria } from '../../../apps/sheets/src/renderer/AdvancedFilterDialog'
import { workbookRangeResultSchema } from '@genoffice/xlsx-gateway/shared/workbook-protocol'
import {
  installOutlineGutter,
  type OutlineGutterHandle,
  type OutlineGutterState,
} from '../../../apps/sheets/src/renderer/outline-gutter'
import {
  visibilityEditsSchema,
  axisHiddenOverride,
} from '@genoffice/xlsx-gateway/domain/session-visibility'
import { outlineReadResultSchema } from '@genoffice/xlsx-gateway/domain/session-outline'
import {
  groupLevelOps,
  outlineGroups,
  groupTogglePlan,
  levelButtonPlan,
  type OutlineGroup,
} from '../../../apps/sheets/src/renderer/outline-model'
import { goalSeekResultSchema } from '@genoffice/xlsx-gateway/shared/session-goal-seek'
import type { GoalSeekResult } from '@genoffice/xlsx-gateway/domain/goal-seek'
import { workbookAuditResultSchema } from '@genoffice/xlsx-gateway/shared/session-audit'
import type { WorkbookRangeResult } from '@genoffice/xlsx-gateway/shared/workbook-protocol'
import { workbookDvRuleToUniver, type DvListSourceContext } from './data-validation-model'
import { dvRuleToWireRule } from '@genoffice/xlsx-gateway/gateway/xlsx-dv'
import {
  buildLabelMatrix,
  buildPositionMatrix,
  parseConsolidateReference,
  targetOverlapsSource,
  type ConsolidateConfig,
  type ConsolidateArea,
} from '../../../apps/sheets/src/renderer/consolidate'
import {
  collectPrecedents,
  installTraceArrows,
} from '../../../apps/sheets/src/renderer/trace-arrows'
import {
  installFormulaViewInterceptor,
  setShowFormulasView,
} from '../../../apps/sheets/src/renderer/formula-view'
import {
  calculateNow as calculateWorkbookNow,
  calculateSheet as calculateActiveSheet,
  isManualCalculation,
  setManualCalculation,
} from '../../../apps/sheets/src/renderer/calc-options'
import { dedupeRows, type DedupeValue } from '../../../apps/sheets/src/renderer/dedupe'
import {
  applyFlashFillTemplate,
  inferFlashFillTemplate,
  type FlashFillExample,
} from '@genoffice/xlsx-gateway/domain/flash-fill'
import {
  watchKey,
  type WatchCell,
  type WatchRowValue,
} from '../../../apps/sheets/src/renderer/WatchWindowPanel'
import type {
  DefinedNameAction,
  DefinedNameRow,
} from '../../../apps/sheets/src/renderer/NameManagerDialog'
import {
  a1RangeRef,
  a1RowRangeRef,
  definedNameFromLabel,
} from '../../../apps/sheets/src/renderer/defined-name-label'
import { isValidDefinedName } from '@genoffice/xlsx-gateway/gateway/xlsx-defined-names'
import {
  sessionDefinedNames,
  sourceTableNames,
} from '@genoffice/xlsx-gateway/domain/session-defined-names'
import { installNameDefinitions } from '../../../apps/sheets/src/renderer/defined-name-projection'
import { autoFunctionCells } from '../../../apps/sheets/src/renderer/auto-function-command'
import { functionPreviewRanges } from './function-preview'
import { readLiveFunctionInfos } from '../../../apps/sheets/src/renderer/function-catalog'
import {
  createFunctionArgumentsHost,
  installFxButtonHook,
  type FunctionArgumentsHost,
  type FunctionTarget,
} from '../../../apps/sheets/src/renderer/function-arguments-runtime'
import {
  pageSettingsCommand,
  type PageSettingsPatch,
} from '../../../apps/sheets/src/renderer/page-settings-command'
import type { PageLayoutEcho } from '../../../apps/sheets/src/renderer/SheetsRibbon'
import {
  defaultPasteSpecial,
  specialPasteBounds,
  specialPasteOperations,
  type SpecialPasteSnapshot,
} from './special-paste'
import {
  needsOwnLayer,
  parseClipboardText,
} from '../../../apps/sheets/src/renderer/paste-special-transform'
import type {
  PasteSpecialOptions,
  PasteSpecialType,
} from '../../../apps/sheets/src/renderer/paste-special-model'
import {
  createCopiedBlock,
  copiedBlockOperation,
  clipboardCopyId,
  type CopiedBlock,
} from './clipboard-copy'
import { SheetClipboardController, ISheetClipboardService } from '@univerjs/preset-sheets-core'
import { planNativeCellPaste } from './native-cell-paste'
import { cellAtClientPoint } from '../../../apps/sheets/src/renderer/picture-grid-target'
import {
  isFormulaBarKeyTarget,
  isGridKeyTarget,
} from '../../../apps/sheets/src/renderer/clear-selection-keyboard'
import {
  isModalOpen,
  type GlobalShortcutGuards,
} from '../../../apps/sheets/src/renderer/global-shortcuts'
import { getScrollAnchor } from '../../../apps/sheets/src/renderer/scroll-anchor'
import { nextSheetName } from '../../../apps/sheets/src/renderer/sheet-default-name'
import { registerExcelShortcuts } from '../../../apps/sheets/src/renderer/excel-shortcuts'
import { PASTE_SPECIAL_OPEN_SHELL_COMMAND } from '../../../apps/sheets/src/renderer/paste-special-model'
import {
  FILTER_BY_COMMAND_ID,
  filterBySelectedCell,
  installCellContextMenu,
  registerPasteSpecialOpener,
  setContextMenuCommandSink,
} from '../../../apps/sheets/src/renderer/cell-context-menu'
import { registerExcelJumpNav } from '../../../apps/sheets/src/renderer/excel-jump-nav'
import { registerEditorWordJump } from '../../../apps/sheets/src/renderer/editor-word-jump'
import {
  FILL_SERIES_OPEN_ENDED_LIMIT,
  planFillSeries,
  type FillSeriesOptions,
} from '../../../apps/sheets/src/renderer/fill-series'
import {
  pickPictureFile,
  pickTransferImage,
  readPictureFile,
  fitPictureFrame,
  installPictureTransfer,
  type PictureAnchor,
} from '../../../apps/sheets/src/renderer/picture-paste'
import {
  startSheetShapeDraw,
  isSheetShapeDragging,
} from '../../../apps/sheets/src/renderer/shape-draw'
import { ADDABLE_SHAPE_TYPES } from '@genoffice/xlsx-gateway/shared/shape-types'
import {
  alignBoxes,
  distributeBoxes,
  type AlignMode,
  type DistributeAxis,
} from '../../../apps/sheets/src/renderer/visual-arrange'
import {
  frameBox,
  frameChangeFromBox,
  anchorForBox,
  sheetMetrics,
} from '../../../apps/sheets/src/renderer/visual-geometry'
import { projectVisualOrder } from '@genoffice/xlsx-gateway/domain/visual-order'
import { reorderZ, type ZOrderOp } from '../../../apps/sheets/src/renderer/visual-arrange'
import type { WorkbookVisualObject } from '@genoffice/xlsx-gateway/shared/workbook-protocol'
import type { SizeDialogResult } from '../../../apps/sheets/src/renderer/VisualDialogs'
import { chartRibbonEdit } from '../../../apps/sheets/src/renderer/chart-ribbon-edit'
import type { SelectedChartRibbon } from '../../../apps/sheets/src/renderer/SheetsRibbon'
import type { ChartTextTarget } from '../../../apps/sheets/src/renderer/ChartTextDialog'
import type { ChartSeriesSelection } from '@genoffice/xlsx-gateway/domain/chart-series-selection'
import { resolveChartRange } from '@genoffice/xlsx-gateway/domain/chart-series-range'
import { UniverSheetsDrawingPreset } from '@univerjs/preset-sheets-drawing'
import { UniverSheetsDataValidationPreset } from '@univerjs/preset-sheets-data-validation'
import '@univerjs/preset-sheets-data-validation/lib/index.css'
import '@univerjs/preset-sheets-drawing/lib/index.css'
import {
  installWorkbookVisuals,
  installedVisualFrames,
  isEditableFileVisual,
  EMU_PER_PIXEL,
  isVisualDragActive,
  isChartEditorOpen,
  isShapeTextEditorOpen,
  isChartApplyActive,
  clearVisualSelection,
  selectedVisualIds,
  subscribeVisualSelection,
  convertibleType,
  setVisualCommandListener,
  setChartDialogListener,
  getChartElementSelection,
  subscribeChartElementSelection,
  type ShapeEditChanges,
  type ChartEditData,
} from '../../../apps/sheets/src/renderer/WorkbookVisuals'
import { installSparklines } from '../../../apps/sheets/src/renderer/WorkbookVisuals'
import {
  sparklineCells,
  type SessionSparklineGroup,
} from '@genoffice/xlsx-gateway/domain/session-sparklines'
import { installRibbonFocusBoundary } from './ribbon-focus-boundary'
import {
  recommendCharts,
  type RecommendedKind,
} from '@genoffice/xlsx-gateway/domain/chart-recommend'
import {
  applyChartStateEdit,
  chartSupportsDataLabels,
  buildChartVisual,
  type ChartGridValue,
} from '@genoffice/xlsx-gateway/domain/chart-visual'
import { decodeHeaderFooter } from '../../../apps/sheets/src/renderer/header-footer-model'
import type { HeaderFooterResult } from '../../../apps/sheets/src/renderer/HeaderFooterDialog'
import { symbolIntents } from './symbol-intent'
import type { WorkbookViewCommand } from '@genoffice/xlsx-gateway/shared/workbook-view-command'
import { HostFindReplaceService } from './find-replace-host'
import { deserializeRangeWithSheet } from '@univerjs/engine-formula'
import {
  numfmt,
  CommandType,
  DataValidationErrorStyle,
  ICommandService,
  IUndoRedoService,
  UndoCommandId,
  RedoCommandId,
  IUniverInstanceService,
  LocaleType,
  LifecycleService,
  LifecycleStages,
  ThemeService,
  RANGE_TYPE,
  BooleanNumber,
  CellValueType,
  type ICellData,
  type IWorkbookData,
} from '@univerjs/core'
import {
  ILayoutService,
  IRenderManagerService,
  SheetSkeletonManagerService,
  UniverSheetsCorePreset,
} from '@univerjs/preset-sheets-core'
import { clampZoomRatio } from '../../../apps/sheets/src/renderer/zoom-range'
import { AddWorksheetMergeMutation } from '@univerjs/sheets'
import {
  DeltaColumnWidthCommand,
  DeltaRowHeightCommand,
  SetColWidthCommand,
  SetRowHeightCommand,
  SheetsSelectionsService,
  SheetSkeletonService,
} from '@univerjs/sheets'
import { installAutofitWrapBudget } from '../../../apps/sheets/src/renderer/autofit-wrap-budget'
import { installAutofitLinePitch } from '../../../apps/sheets/src/renderer/autofit-line-pitch'
import {
  orderDefinedNames,
  type DefinedNameDefinition,
} from '../../../apps/sheets/src/renderer/defined-name-order'
import { revealCellBelowFreeze } from '../../../apps/sheets/src/renderer/reveal-cell'
import { resolveGoToRef, type GoToNameEntry } from '../../../apps/sheets/src/renderer/goto'
import { t } from '../../../apps/sheets/src/renderer/i18n/locale'
import {
  CSV_CHARSET_BY_LANG,
  CSV_IMPORT_MAX_BYTES,
  csvImportMatrix,
  pickCsvFile,
} from '../../../apps/sheets/src/renderer/csv-import-model'
import { decodeCsvBuffer } from '@genoffice/xlsx-gateway/gateway/csv-import'
import {
  installCrossHighlight,
  loadCrossHighlightPreference,
  storeCrossHighlightPreference,
  type CrossHighlightHandle,
} from '../../../apps/sheets/src/renderer/cross-highlight'
import { installPageBreakPreview } from '../../../apps/sheets/src/renderer/page-break-preview'
import type { PageSetupJournalState } from '../../../apps/sheets/src/renderer/edit-journal'
import {
  filePivotsOfSheet,
  pivotSourceSheet,
  sessionTablesOf,
  workbookStructureLocked,
} from '@genoffice/xlsx-gateway/domain/session-candidate'
import {
  PIVOT_SOURCE_COL_LIMIT,
  PIVOT_SOURCE_ROW_LIMIT,
} from '@genoffice/xlsx-gateway/domain/pivot-layout'
import { buildPivotChartData } from '@genoffice/xlsx-gateway/domain/pivot-chart'
import {
  isNumericPivotColumn,
  pivotPageRows,
  pivotSourceGrid,
} from '@genoffice/xlsx-gateway/domain/pivot-placement'
import type { WorkbookPivotAdd } from '@genoffice/xlsx-gateway/shared/workbook-protocol'
import {
  defaultPivotModel,
  modelToConfig,
  validatePivotModel,
  type PivotField,
  type PivotLayoutModel,
} from '../../../apps/sheets/src/renderer/pivot-field-model'
import {
  modelOfDefinition,
  modelOfPivotAdd,
  pivotConfigToOpParts,
  pivotErrorMessage,
} from '../../../apps/sheets/src/renderer/pivot-op-parts'
import type { PivotCreateRequest } from '../../../apps/sheets/src/renderer/PivotDialog'
import { fullRowSpans, fullColumnSpans } from '../../../apps/sheets/src/renderer/autofit-selection'
import { dimensionIntents } from './dimension-intents'
import {
  currentAxisIndex,
  originalAxisIndex,
} from '@genoffice/xlsx-gateway/domain/session-coordinates'
import enUS from '@univerjs/preset-sheets-core/locales/en-US'
import zhCN from '@univerjs/preset-sheets-core/locales/zh-CN'
import zhTW from '@univerjs/preset-sheets-core/locales/zh-TW'
import '@univerjs/preset-sheets-core/lib/index.css'
import {
  dvRuleSchema,
  workbookOperationSchema,
  type WorkbookOperation,
  type DvRule,
} from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import {
  formatAddress,
  parseAddress,
  parseRange,
  rangeCellCount,
  columnLabel,
} from '@genoffice/xlsx-gateway/domain/cell-address'
import { workbookStateSchema, receiptSchema } from '@genoffice/xlsx-gateway/shared/session-state'
import { MAX_GRID_COLUMNS, MAX_GRID_ROWS } from '@genoffice/xlsx-gateway/shared/grid-bounds'
import type { CandidateState } from './changeset'
import type { EditorHost, EditorProjection, GridRange } from './ports'
import { createUniver } from './create-univer'
import { manualStyleIntent, manualValueIntents } from './manual-intents'
import { gridCell, type ViewCell } from './grid-cell'
import { requestHostHistory } from './history-bridge'
import type { SheetVisual, SheetChartVisual } from '@genoffice/xlsx-gateway/domain/chart-visual'
import type { WorkbookAppearance } from './appearance'
import {
  toSelectionFormat,
  indentStepsOf,
} from '../../../apps/sheets/src/renderer/selection-format'
import {
  axisColumnWidthToPixels,
  measureNormalFontMdw,
  type NormalFontDescriptor,
} from '../../../apps/sheets/src/renderer/normal-font-metrics'
import { ribbonIntents, type RibbonTarget } from './ribbon-intents'
import type { DecimalCell } from '../../../apps/sheets/src/renderer/decimal-step-model'

export interface SpecialPasteContext {
  target: Pick<RibbonTarget, 'sheetId' | 'address'>
  stateId: string
  revision: number
  snapshot: SpecialPasteSnapshot
}

type OutlineAction =
  | 'group'
  | 'ungroup'
  | 'hide-detail'
  | 'show-detail'
  | 'clear'
  | 'summary-below'
  | 'summary-right'
  | 'level'
type OutlineGutterTarget = {
  stateId: string
  sheetId: string
  group?: OutlineGroup
  level?: number
}
export class WorkbookRuntime implements EditorProjection {
  private auditWork: Promise<unknown> | undefined
  private traceDisposables: { dispose(): void }[] = []
  private traceSequence = 0
  auditNotice: string | null = null
  watchOpen = false
  watches: WatchCell[] = []
  private watchValues = new Map<string, { stateId: string; value: WatchRowValue }>()
  private watchWork: Promise<void> | undefined
  private clearTraceArrows() {
    for (const item of this.traceDisposables.splice(0).reverse()) item.dispose()
  }
  private showFormulasFor(sheetId: string): boolean {
    const sheet = this.state?.sheets.find((item) => item.id === sheetId)
    return (
      (sheet?.features?.pageSetupEdits as { showFormulas?: boolean } | undefined)?.showFormulas ??
      (sheet?.features?.originalMetadata as { showFormulas?: boolean } | undefined)?.showFormulas ??
      (this.filePageSetup.get(sheetId) as { showFormulas?: boolean } | undefined)?.showFormulas ??
      false
    )
  }
  private syncFormulaView() {
    const id = this.runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId()
    setShowFormulasView(this.runtime, id ? this.showFormulasFor(id) : false)
  }
  toggleWatchWindow() {
    this.watchOpen = !this.watchOpen
    this.changed()
  }
  async addWatchSelection() {
    if (!(await this.flushEditor())) throw new Error(t('appClipboardBusy'))
    const selected = this.selection
    if (!selected) return
    const seen = new Set(this.watches.map(watchKey))
    const added: WatchCell[] = []
    outer: for (let row = selected.range.startRow; row <= selected.range.endRow; row++)
      for (let column = selected.range.startColumn; column <= selected.range.endColumn; column++) {
        if (added.length >= 20) break outer
        const cell = { sheetId: selected.sheetId, row, column }
        if (!seen.has(watchKey(cell))) added.push(cell)
      }
    this.watches = [...this.watches, ...added]
    this.changed()
    await this.refreshWatches()
  }
  removeWatch(key: string) {
    this.watches = this.watches.filter((cell) => watchKey(cell) !== key)
    this.watchValues.delete(key)
    this.changed()
  }
  resolveWatch(cell: WatchCell): WatchRowValue | null {
    const sheet = this.state?.sheets.find((item) => item.id === cell.sheetId)
    if (!sheet) return null
    const cached = this.watchValues.get(watchKey(cell))
    return cached && cached.stateId === this.state?.stateId
      ? { ...cached.value, sheetName: sheet.name }
      : {
          sheetName: sheet.name,
          value: '…',
          formula: sheet.cells[formatAddress(cell.row, cell.column)]?.formula ?? '',
        }
  }
  readonly refreshWatches = async (): Promise<void> => {
    if (this.watchWork) return this.watchWork
    if (!this.state || this.closed) return
    const stateId = this.state.stateId
    const work = (async () => {
      const values = await Promise.all(
        this.watches.map(async (cell) => {
          const sheet = this.state?.sheets.find((item) => item.id === cell.sheetId)
          if (!sheet) return null
          const read = (await this.host.read(
            cell.sheetId,
            {
              startRow: cell.row,
              endRow: cell.row,
              startColumn: cell.column,
              endColumn: cell.column,
            },
            stateId,
          )) as { cells: ViewCell[] }
          const value = read.cells.find(
            (item) => item.address === formatAddress(cell.row, cell.column),
          )
          const pending = value?.formula && value.calculationState !== 'ready'
          const raw = value?.cachedValue ?? value?.rawValue
          return {
            key: watchKey(cell),
            value: {
              sheetName: sheet.name,
              value: pending
                ? '…'
                : (value?.displayValue ??
                  (typeof raw === 'number' && value?.numberFormat
                    ? numfmt.format(
                        value.numberFormat,
                        raw + (value.kind === 'date' && value.dateSystem === '1904' ? 1462 : 0),
                        { throws: false },
                      )
                    : raw == null
                      ? ''
                      : String(raw))),
              formula: value?.formula ?? '',
            },
          }
        }),
      )
      if (this.closed || this.state?.stateId !== stateId) return
      const live = new Set(this.watches.map(watchKey))
      for (const item of values)
        if (item && live.has(item.key))
          this.watchValues.set(item.key, { stateId, value: item.value })
    })()
    this.watchWork = work
    try {
      await work
    } finally {
      if (this.watchWork === work) this.watchWork = undefined
    }
  }
  async auditCommand(command: string) {
    if (this.auditWork) return
    if (!(await this.flushEditor())) throw new Error(t('appClipboardBusy'))
    if (this.auditWork) return
    const selected = this.selection,
      state = this.state
    const sheet = this.runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
    if (!selected || !state || !sheet) throw new Error(t('appSelectCellFirst'))
    const frozen = { stateId: state.stateId, revision: this.revision }
    const address = formatAddress(selected.range.startRow, selected.range.startColumn)
    const current = () =>
      !this.closed && this.state?.stateId === frozen.stateId && this.revision === frozen.revision
    const assertCurrent = () => {
      if (!current()) throw new Error(t('appGoToNotReady'))
    }
    let jump: { sheetId: string; address: string } | undefined
    const work = (async () => {
      if (command === 'remove-arrows') {
        const any = this.traceDisposables.length > 0
        this.clearTraceArrows()
        this.auditNotice = t(any ? 'appTraceArrowsRemoved' : 'appNoTraceArrows')
        return
      }
      if (command === 'toggle-show-formulas') {
        const result = (await this.host.read(
          selected.sheetId,
          {
            startRow: selected.range.startRow,
            endRow: selected.range.startRow,
            startColumn: selected.range.startColumn,
            endColumn: selected.range.startColumn,
          },
          state.stateId,
        )) as { pageSetup?: unknown }
        assertCurrent()
        if (Object.hasOwn(result, 'pageSetup'))
          this.filePageSetup.set(selected.sheetId, result.pageSetup)
        const showFormulas = !this.showFormulasFor(selected.sheetId)
        await this.enqueue(
          [{ op: 'set_page_setup', sheetId: selected.sheetId, showFormulas }],
          frozen,
          { failureInDialog: true },
        )
        this.auditNotice = t(showFormulas ? 'appShowingFormulas' : 'appShowingValues')
        return
      }
      const idPrefix = `trace-plugin-${++this.traceSequence}`
      if (command === 'trace-precedents') {
        const result = (await this.host.read(
          selected.sheetId,
          {
            startRow: selected.range.startRow,
            endRow: selected.range.startRow,
            startColumn: selected.range.startColumn,
            endColumn: selected.range.startColumn,
          },
          state.stateId,
        )) as { cells: ViewCell[] }
        assertCurrent()
        const formula = result.cells.find((cell) => cell.address === address)?.formula
        if (!formula) {
          this.auditNotice = t('appTraceNoFormula')
          return
        }
        const canonical = state.sheets.find((item) => item.id === selected.sheetId)!
        const resultRefs = collectPrecedents(formula, sheet.getSheetName(), {
          lastRow: canonical.rows - 1,
          lastColumn: canonical.columns - 1,
          maxRows: sheet.getMaxRows(),
          maxColumns: sheet.getMaxColumns(),
        })
        if (!resultRefs.areas.length && !resultRefs.offSheet) {
          this.auditNotice = t('appTraceNoRefs')
          return
        }
        this.traceDisposables.push(
          ...installTraceArrows(
            this.runtime,
            sheet,
            resultRefs.areas.map((area) => ({
              from: { row: area.startRow, column: area.startColumn },
              to: { row: selected.range.startRow, column: selected.range.startColumn },
            })),
            resultRefs.areas,
            idPrefix,
          ),
        )
        this.auditNotice =
          t('appTracedPrecedents', { count: resultRefs.areas.length }) +
          (resultRefs.offSheet
            ? ` ${t('appTraceOffSheetRefs', { count: resultRefs.offSheet })}`
            : '')
        return
      }
      if (!this.host.audit) throw new Error('Workbook audit is unavailable in this host.')
      if (command === 'error-checking') {
        this.auditNotice = t('appCheckingErrors')
        this.changed()
      }
      const result = workbookAuditResultSchema.parse(
        await this.host.audit({
          stateId: state.stateId,
          kind: command === 'trace-dependents' ? 'dependents' : 'errors',
          sheetId: selected.sheetId,
          address,
        }),
      )
      assertCurrent()
      if (result.stateId !== state.stateId) throw new Error('invalid_audit_state')
      if (command === 'trace-dependents') {
        const local = result.hits.filter((hit) => hit.sheetId === selected.sheetId)
        const offSheet = result.hits.length - local.length
        this.internal++
        try {
          const maxRow = local.reduce(
            (max, hit) => Math.max(max, parseAddress(hit.address).row),
            sheet.getMaxRows() - 1,
          )
          const maxColumn = local.reduce(
            (max, hit) => Math.max(max, parseAddress(hit.address).column),
            sheet.getMaxColumns() - 1,
          )
          if (maxRow >= sheet.getMaxRows()) sheet.setRowCount(maxRow + 1)
          if (maxColumn >= sheet.getMaxColumns()) sheet.setColumnCount(maxColumn + 1)
          this.traceDisposables.push(
            ...installTraceArrows(
              this.runtime,
              sheet,
              local.map((hit) => ({
                from: { row: selected.range.startRow, column: selected.range.startColumn },
                to: parseAddress(hit.address),
              })),
              [],
              idPrefix,
            ),
          )
        } finally {
          this.internal--
        }
        this.auditNotice = result.hits.length
          ? t('appTracedDependents', { count: local.length, note: '' }) +
            (offSheet ? ` ${t('appTraceOffSheetDeps', { count: offSheet })}` : '')
          : t('appTraceNoDependents', { note: '' })
      } else {
        const order = new Map(state.sheets.map((item, index) => [item.id, index]))
        const rank = (hit: { sheetId: string; address: string }) => {
          const point = parseAddress(hit.address)
          return (
            (order.get(hit.sheetId) ?? 0) * 1_048_576 * 16384 + point.row * 16384 + point.column
          )
        }
        jump =
          result.hits.find((hit) => rank(hit) > rank({ sheetId: selected.sheetId, address })) ??
          result.hits[0]
        this.auditNotice = jump
          ? t('appErrorsFound', {
              count: result.hits.length,
              cell: jump.address,
              value: result.hits.find((hit) => hit === jump)?.value ?? '',
            })
          : t('appNoErrorsFound')
        if (result.valuesPending) this.auditNotice += ` ${t('appAuditPendingValues')}`
      }
    })()
    this.auditWork = work
    try {
      await work
    } finally {
      if (this.auditWork === work) this.auditWork = undefined
      this.changed()
    }
    if (jump && current()) {
      const name = state.sheets
        .find((item) => item.id === jump!.sheetId)!
        .name.replaceAll("'", "''")
      const error = await this.goToReference(`'${name}'!${jump.address}`, { isCurrent: current })
      if (error) throw new Error(error)
    }
  }

  private nameWork: Promise<unknown> | undefined
  nameManagerContext: { stateId: string; revision: number } | null = null
  get nameManagerSheets() {
    return this.state?.sheets.map(({ id, name }) => ({ id, name })) ?? []
  }
  get nameManagerRows(): DefinedNameRow[] {
    return this.state
      ? sessionDefinedNames(this.state).names.map((name) => ({
          name: name.name,
          ref: name.formula,
          scopeSheetId: name.scopeSheetId ?? null,
          scopeLabel: name.scopeSheetId
            ? (this.state!.sheets.find((sheet) => sheet.id === name.scopeSheetId)?.name ??
              t('appNameGone'))
            : t('dlgNmWorkbook'),
        }))
      : []
  }
  async openNameManager() {
    if (this.functionDialog || !(await this.flushEditor())) throw new Error(t('appClipboardBusy'))
    if (!this.state || this.closed) throw new Error(t('appWorkbookNotReady'))
    this.nameManagerContext = { stateId: this.state.stateId, revision: this.revision }
    this.changed()
  }
  closeNameManager() {
    if (this.nameWork) return
    this.nameManagerContext = null
    this.changed()
  }
  private assertNameContext(frozen: { stateId: string; revision: number }) {
    if (this.closed || this.state?.stateId !== frozen.stateId || this.revision !== frozen.revision)
      throw new Error(t('appDocumentChangedDuringAction'))
  }
  async applyNameAction(action: DefinedNameAction): Promise<string | null> {
    const frozen = this.nameManagerContext
    if (!frozen || this.nameWork) return t('appClipboardBusy')
    const work = (async () => {
      if (!(await this.flushCellEditor())) throw new Error(t('appClipboardBusy'))
      this.assertNameContext(frozen)
      const scope = action.kind === 'add' ? action.sheetId : action.scopeSheetId
      const operations: WorkbookOperation[] =
        action.kind === 'add'
          ? [
              {
                op: 'add_defined_names',
                names: [
                  { name: action.name, ref: action.ref, ...(scope ? { scopeSheetId: scope } : {}) },
                ],
              },
            ]
          : [
              {
                op: 'manage_defined_name',
                name: action.kind === 'update' ? action.originalName : action.name,
                ...(scope ? { scopeSheetId: scope } : {}),
                replacement:
                  action.kind === 'remove' ? null : { name: action.name, ref: action.ref },
              },
            ]
      const receipt = receiptSchema.parse(
        await this.enqueue(operations, frozen, { failureInDialog: true }),
      )
      if (this.nameManagerContext === frozen)
        this.nameManagerContext = { stateId: receipt.stateId, revision: receipt.revision }
      this.changed()
    })()
    this.nameWork = work
    try {
      await work
      return null
    } catch (error) {
      const detail = error instanceof Error ? error.message : String(error)
      if (detail.includes('duplicate_defined_name')) return t('appDefinedNameDuplicate')
      if (
        detail.includes('invalid_defined_name') ||
        detail.includes('empty_defined_name_reference')
      )
        return t('appDefinedNameInvalid')
      if (
        detail.includes('defined_name_destination_shadowed') ||
        detail.includes('unmodeled_defined_name_binding')
      )
        return t('appDefinedNameBindingConflict')
      if (detail.includes('defined_name_missing')) return t('appNameGone')
      if (
        detail.includes('defined_name_bindings_incomplete') ||
        detail.includes('protected_workbook')
      )
        return t('appNameApplyFailed')
      return detail
    } finally {
      if (this.nameWork === work) this.nameWork = undefined
    }
  }
  async useNameInFormula(name: string) {
    if (!(await this.flushEditor())) throw new Error(t('appClipboardBusy'))
    const target = this.ribbonTarget
    if (!target || !this.state) throw new Error(t('appSelectCellFirst'))
    if (!this.goToNames.some((entry) => entry.name === name)) throw new Error(t('appNameGone'))
    const frozen = { stateId: this.state.stateId, revision: this.revision }
    const bounds = parseRange(target.address),
      address = formatAddress(bounds.startRow, bounds.startColumn)
    const work = (async () => {
      const result = (await this.host.read(
        target.sheetId,
        {
          startRow: bounds.startRow,
          endRow: bounds.startRow,
          startColumn: bounds.startColumn,
          endColumn: bounds.startColumn,
        },
        frozen.stateId,
      )) as { cells: ViewCell[] }
      this.assertNameContext(frozen)
      const cell = result.cells.find((cell) => cell.address === address)
      if (
        cell?.formula ||
        (cell?.rawValue !== undefined && cell.rawValue !== null && cell.rawValue !== '')
      )
        throw new Error(t('appUseInFormulaNeedsEmptyCell'))
      await this.enqueue(
        [{ op: 'set_formula', sheetId: target.sheetId, address, formula: `=${name}` }],
        frozen,
      )
    })()
    this.nameWork = work
    try {
      await work
    } finally {
      if (this.nameWork === work) this.nameWork = undefined
    }
  }
  async createNamesFromSelection(mode: 'top' | 'left') {
    if (!(await this.flushEditor())) throw new Error(t('appClipboardBusy'))
    const target = this.ribbonTarget,
      state = this.state
    if (!target || !state) throw new Error(t('appSelectCellFirst'))
    const bounds = parseRange(target.address),
      sheet = state.sheets.find((sheet) => sheet.id === target.sheetId)!
    if (mode === 'top' ? bounds.endRow <= bounds.startRow : bounds.endColumn <= bounds.startColumn)
      throw new Error(t('appCreateNamesNeedsData'))
    const endRow =
      mode === 'top' ? bounds.startRow : Math.min(bounds.endRow, Math.max(1, sheet.rows) - 1)
    const endColumn =
      mode === 'left'
        ? bounds.startColumn
        : Math.min(bounds.endColumn, Math.max(1, sheet.columns) - 1)
    if (endRow < bounds.startRow || endColumn < bounds.startColumn) return { count: 0, skipped: 0 }
    const labels = { startRow: bounds.startRow, endRow, startColumn: bounds.startColumn, endColumn }
    if (rangeCellCount(labels) > 2000) throw new Error(t('appClipboardUnsupported'))
    const frozen = { stateId: state.stateId, revision: this.revision }
    const work = (async () => {
      const result = (await this.host.read(target.sheetId, labels, frozen.stateId)) as {
        cells: ViewCell[]
      }
      this.assertNameContext(frozen)
      const model = sessionDefinedNames(state)
      const reserved =
        (
          state.features?.originalMetadata as
            { reservedDefinedNames?: { name: string }[] } | undefined
        )?.reservedDefinedNames ?? []
      const taken = new Set(
        [
          ...model.names.map((name) => name.name),
          ...model.preserveNames,
          ...reserved.map((name) => name.name),
          ...sourceTableNames(state),
        ].map((name) => name.toLowerCase()),
      )
      const cells = new Map(result.cells.map((cell) => [cell.address, cell]))
      const names: { name: string; ref: string }[] = []
      let skipped = 0
      for (let row = labels.startRow; row <= labels.endRow; row++)
        for (let column = labels.startColumn; column <= labels.endColumn; column++) {
          const cell = cells.get(formatAddress(row, column))
          if (
            cell?.formula &&
            cell.calculationState !== 'ready' &&
            cell.calculationState !== 'error'
          )
            throw new Error(t('appDefinedNameLabelPending'))
          let label = cell?.displayValue ?? (cell?.rawValue == null ? '' : String(cell.rawValue))
          const format = cell?.numberFormat ?? cell?.style?.numberFormat
          if (
            cell &&
            cell.displayValue === undefined &&
            typeof cell.rawValue === 'number' &&
            format
          )
            label = numfmt.format(
              format,
              cell.rawValue + (cell.kind === 'date' && cell.dateSystem === '1904' ? 1462 : 0),
              { throws: false },
            )
          const name = definedNameFromLabel(label)
          if (
            !name ||
            !isValidDefinedName(name) ||
            name.toLowerCase().startsWith('_xlnm') ||
            taken.has(name.toLowerCase())
          ) {
            skipped++
            continue
          }
          taken.add(name.toLowerCase())
          names.push({
            name,
            ref:
              mode === 'top'
                ? a1RangeRef(sheet.name, column, bounds.startRow + 1, bounds.endRow)
                : a1RowRangeRef(sheet.name, row, bounds.startColumn + 1, bounds.endColumn),
          })
        }
      if (names.length) await this.enqueue([{ op: 'add_defined_names', names }], frozen)
      return { count: names.length, skipped }
    })()
    this.nameWork = work
    try {
      return await work
    } finally {
      if (this.nameWork === work) this.nameWork = undefined
    }
  }
  private functionWork: Promise<void> | undefined
  functionDialog: {
    category: string
    target: FunctionTarget
    functions: ReturnType<typeof readLiveFunctionInfos>
    host: FunctionArgumentsHost
  } | null = null
  closeInsertFunction() {
    this.functionDialog = null
    this.changed()
  }
  private functionOpening = false
  async openInsertFunction(category: string) {
    if (this.functionDialog || this.functionOpening) return
    this.functionOpening = true
    try {
      await this.prepareInsertFunction(category)
    } finally {
      this.functionOpening = false
    }
  }
  private async prepareInsertFunction(category: string) {
    if (this.composing) throw new Error(t('appClipboardBusy'))
    await this.pending
    const native = createFunctionArgumentsHost(
      () => (this.closed ? null : this.runtime),
      () => t('appActiveSheetUnavailable'),
      this.container,
    )
    if (!this.state || this.closed) throw new Error(t('appActiveSheetUnavailable'))
    const frozen = { stateId: this.state.stateId, revision: this.revision }
    const target = await native.readTarget()
    if (!target || !(await this.flushEditor()) || !this.state)
      throw new Error(t('appActiveSheetUnavailable'))
    const assertCurrent = () => {
      if (
        this.closed ||
        this.state?.stateId !== frozen.stateId ||
        this.revision !== frozen.revision
      )
        throw new Error(t('appDocumentChangedDuringAction'))
    }
    try {
      assertCurrent()
    } catch (error) {
      native.cancel(target)
      throw error
    }
    const work = (async () => {
      const read = (await this.host.read(
        target.sheetId,
        {
          startRow: target.row,
          endRow: target.row,
          startColumn: target.col,
          endColumn: target.col,
        },
        frozen.stateId,
      )) as { cells: ViewCell[] }
      assertCurrent()
      const cell = read.cells.find((cell) => cell.address === formatAddress(target.row, target.col))
      const currentTarget =
        target.editorText === null
          ? { ...target, formula: cell?.formula ?? '', caret: (cell?.formula ?? '').length }
          : target
      const host: FunctionArgumentsHost = {
        ...native,
        async readTarget() {
          return currentTarget
        },
        write: async (_target, formula) => {
          const work = (async () => {
            if (!(await this.flushCellEditor())) throw new Error(t('appClipboardBusy'))
            assertCurrent()
            await this.enqueue(
              [
                {
                  op: 'set_formula',
                  sheetId: target.sheetId,
                  address: formatAddress(target.row, target.col),
                  formula,
                },
              ],
              frozen,
              { failureInDialog: true },
            )
          })()
          this.functionWork = work
          try {
            await work
            return null
          } catch (error) {
            return error instanceof Error ? error.message : String(error)
          } finally {
            if (this.functionWork === work) this.functionWork = undefined
          }
        },
        evaluate: async (_target, expressions) => {
          assertCurrent()
          const values = []
          for (const expression of expressions) {
            const ranges = functionPreviewRanges(expression, target.sheetId, this.state!.sheets)
            if (!ranges) {
              values.push(undefined)
              continue
            }
            const reads = []
            for (const range of ranges) {
              reads.push({
                ...range,
                result: (await this.host.read(range.sheetId, range.range, frozen.stateId)) as {
                  cells: ViewCell[]
                },
              })
              assertCurrent()
            }
            // A pending formula or unresolved calculation cannot be treated as a blank/zero precedent.
            if (reads.some((read) => read.result.cells.some((cell) => cell.formula))) {
              values.push(undefined)
              continue
            }
            for (const { sheetId, range, result } of reads) {
              const sheet = this.runtime.univerAPI.getActiveWorkbook()?.getSheetBySheetId(sheetId)
              if (!sheet) throw new Error(t('appActiveSheetUnavailable'))
              const cells: Record<number, Record<number, ICellData>> = {}
              for (let row = range.startRow; row <= range.endRow; row++)
                for (let column = range.startColumn; column <= range.endColumn; column++)
                  (cells[row] ??= {})[column] = { v: null }
              for (const cell of result.cells) {
                const at = parseAddress(cell.address)
                if (
                  at.row >= range.startRow &&
                  at.row <= range.endRow &&
                  at.column >= range.startColumn &&
                  at.column <= range.endColumn
                )
                  (cells[at.row] ??= {})[at.column] = gridCell(cell)
              }
              this.internal++
              try {
                sheet
                  .getRange(
                    range.startRow,
                    range.startColumn,
                    range.endRow - range.startRow + 1,
                    range.endColumn - range.startColumn + 1,
                  )
                  .setValues(cells)
                this.runtime.univer
                  .__getInjector()
                  .get(IUndoRedoService)
                  .clearUndoRedo(target.unitId)
              } finally {
                this.internal--
              }
            }
            values.push((await native.evaluate(currentTarget, [expression]))[0])
            assertCurrent()
          }
          return values
        },
      }
      this.functionDialog = {
        category,
        target: currentTarget,
        host,
        functions: readLiveFunctionInfos(this.runtime),
      }
      this.changed()
    })()
    this.functionWork = work
    try {
      await work
    } catch (error) {
      native.cancel(target)
      throw error
    } finally {
      if (this.functionWork === work) this.functionWork = undefined
    }
  }
  async insertAutoFunction(functionName: string) {
    if (!(await this.flushEditor())) throw new Error(t('appClipboardBusy'))
    const selected = this.selection
    if (!selected || !/^(SUM|AVERAGE|COUNT|MIN|MAX)$/.test(functionName))
      throw new Error(t('appSelectCellsFirst'))
    const bounds = selected.range
    if (bounds.endRow <= bounds.startRow)
      throw new Error(t('appAutofnSelectCells', { fn: functionName }))
    const frozen = this.selectionContext
    const operations: WorkbookOperation[] = autoFunctionCells(functionName, selected.address).map(
      (cell) => ({
        op: 'set_formula',
        sheetId: selected.sheetId,
        ...cell,
      }),
    )
    await this.enqueue(operations, frozen)
  }
  readonly findReplace: HostFindReplaceService
  private goToVisible = false
  get goToOpen() {
    return this.goToVisible
  }
  closeGoTo() {
    this.goToVisible = false
    this.changed()
  }
  async openViewCommand(command: WorkbookViewCommand) {
    if (!(await this.flushEditor())) throw new Error('Finish editing before opening this dialog.')
    if (this.closed || !this.state) throw new Error('The document view has closed.')
    if (command === 'goto') this.goToVisible = true
    else this.findReplace.open(command === 'replace')
    this.changed()
  }
  get findWorkbook() {
    return this.runtime.univerAPI.getActiveWorkbook()?.getWorkbook() ?? null
  }
  get ribbonTarget(): RibbonTarget | null {
    const selected = this.selection
    if (!selected) return null
    const sheet = this.runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
    return {
      ...(sheet
        ? { dimensions: { rows: sheet.getMaxRows(), columns: sheet.getMaxColumns() } }
        : {}),
      maximumDigitWidth: measureNormalFontMdw(
        (this.state?.features?.originalMetadata ?? {}) as NormalFontDescriptor,
      ),
      sheetId: selected.sheetId,
      address: selected.address,
      format: toSelectionFormat(selected.style, selected.style.n?.pattern ?? 'General'),
      doubleUnderline: selected.style.ul?.s === 1 && selected.style.ul.t === 10,
    }
  }
  private ribbonFocusBoundary: ReturnType<typeof installRibbonFocusBoundary> | undefined
  private visualWork: Promise<unknown> | undefined
  private chartVisualDisposables: { dispose(): void }[] = []
  private chartVisualSheet: string | undefined
  private clearChartVisuals() {
    this.internal++
    try {
      for (const item of this.chartVisualDisposables) item.dispose()
    } finally {
      this.internal--
    }
    this.chartVisualDisposables = []
    this.chartVisualSheet = undefined
  }
  private installChartVisuals() {
    const sheetId = this.runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId()
    if (!sheetId || this.closed || !this.state) return
    this.clearChartVisuals()
    const frozen = { stateId: this.state.stateId, revision: this.revision }
    this.internal++
    try {
      this.chartVisualDisposables = installWorkbookVisuals(
        this.runtime,
        {
          sessionId: 'embedded',
          visuals: [...this.visuals],
          readMedia: async ({ visualId }) => {
            if (!this.host.media) throw new Error('unsupported_media')
            return this.host.media(visualId, frozen.stateId)
          },
        },
        sheetId,
        {
          edits: new Map(),
          canEdit: (id) =>
            this.charts.some((visual) => visual.id === id || visual.chartPath === id),
          dialogs: new Set(['format', 'select-data']),
          visualCommands: new Set([
            'visual:alt-text',
            'visual:size',
            'visual:z:front',
            'visual:z:back',
            'visual:z:forward',
            'visual:z:backward',
          ]),
          onError: this.report,
          readVector: (id, range) => this.readChartVector(id, range, frozen),
          onEdit: (id, edit) => this.editInlineChart(id, edit, frozen),
        },
        {
          onEdit: (id, changes) => {
            void this.editVisual(sheetId, id, changes, frozen).catch(this.report)
          },
          onTextEdit: (id, changes) => this.editVisual(sheetId, id, changes, frozen).then(() => {}),
        },
      )
      this.chartVisualSheet = sheetId
      this.runtime.univer.__getInjector().get(IUndoRedoService).clearUndoRedo('projection')
    } finally {
      this.internal--
    }
  }
  async insertPicture() {
    if (!(await this.flushEditor())) throw new Error('Finish editing before inserting a picture.')
    const target = this.ribbonTarget
    if (!target || !this.state || this.closed) throw new Error('missing_selection')
    const frozen = { stateId: this.state.stateId, revision: this.revision }
    const file = await pickPictureFile()
    if (!file) return
    await this.insertPictureFile(file, target, frozen)
  }
  private async insertPictureFile(
    file: File,
    target: Pick<RibbonTarget, 'sheetId' | 'address'>,
    frozen: { stateId: string; revision: number },
  ) {
    if (this.visualWork || this.chartWork || this.symbolWork)
      throw new Error('Wait for the current visual edit to finish.')
    const work = (async () => {
      if (!(await this.flushCellEditor()))
        throw new Error('Finish editing before inserting a picture.')
      let picture
      try {
        picture = await readPictureFile(file)
      } catch (error) {
        throw new Error(
          t(
            error instanceof Error && error.message === 'too-large'
              ? 'appPictureTooLarge'
              : 'appPictureBadType',
          ),
          { cause: error },
        )
      }
      const { rows, columns } = fitPictureFrame(picture.width, picture.height)
      const cell = parseAddress(target.address.split(':')[0]!)
      await this.enqueue(
        [
          workbookOperationSchema.parse({
            op: 'add_picture',
            sheetId: target.sheetId,
            name: file.name,
            mediaType: picture.mediaType,
            base64: picture.dataUrl.split(',')[1],
            anchor: {
              fromRow: cell.row,
              fromColumn: cell.column,
              fromRowOffset: 0,
              fromColumnOffset: 0,
              toRow: Math.min(1048575, cell.row + rows),
              toColumn: Math.min(16383, cell.column + columns),
              toRowOffset: 0,
              toColumnOffset: 0,
            },
          }),
        ],
        frozen,
        { failureInDialog: true },
      )
    })()
    this.visualWork = work
    try {
      await work
    } finally {
      if (this.visualWork === work) this.visualWork = undefined
    }
  }
  private copiedSource: { sheetId: string; widths: number[] } | undefined
  private copiedBlock: CopiedBlock | undefined
  private copiedCut:
    { id: string; sheetId: string; address: string; stateId: string; revision: number } | undefined
  private consumedCutId: string | undefined
  async copyClipboard(mode: 'COPY' | 'CUT' = 'COPY') {
    if (this.clipboardWork) throw new Error(t('appClipboardBusy'))
    if (this.runtime.univerAPI.getActiveWorkbook()?.isCellEditing()) {
      await this.runtime.univerAPI.executeCommand(
        mode === 'CUT' ? 'univer.command.cut' : 'univer.command.copy',
      )
      return
    }
    if (selectedVisualIds().length) throw new Error(t('appClipboardUnsupported'))
    if (!(await this.flushEditor())) throw new Error(t('appClipboardBusy'))
    const target = this.ribbonTarget
    if (!target || !this.state || this.closed) throw new Error('missing_selection')
    const frozen = { stateId: this.state.stateId, revision: this.revision },
      bounds = parseRange(target.address)
    if (rangeCellCount(bounds) > 2000) throw new Error(t('appClipboardUnsupported'))
    const work = (async () => {
      const read = (await this.host.read(target.sheetId, bounds, frozen.stateId)) as {
        cells: ViewCell[]
        merges?: GridRange[]
      }
      if (
        this.closed ||
        this.state?.stateId !== frozen.stateId ||
        this.revision !== frozen.revision
      )
        throw new Error(t('appClipboardSelectionChanged'))
      const selection = this.ribbonTarget,
        sheet = this.runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
      if (!sheet || selection?.sheetId !== target.sheetId || selection.address !== target.address)
        throw new Error(t('appClipboardSelectionChanged'))
      const rows = Array.from(
          { length: bounds.endRow - bounds.startRow + 1 },
          (_, r) => bounds.startRow + r,
        ),
        columns = Array.from(
          { length: bounds.endColumn - bounds.startColumn + 1 },
          (_, c) => bounds.startColumn + c,
        )
      createCopiedBlock('preflight', bounds, read, rows, columns)
      const values: Record<number, Record<number, ICellData>> = {}
      for (const cell of read.cells) {
        const point = parseAddress(cell.address)
        if (
          point.row >= bounds.startRow &&
          point.row <= bounds.endRow &&
          point.column >= bounds.startColumn &&
          point.column <= bounds.endColumn
        )
          (values[point.row] ??= {})[point.column] = gridCell(cell)
      }
      this.internal++
      try {
        sheet
          .getRange(bounds.startRow, bounds.startColumn, rows.length, columns.length)
          .setValues(values)
        this.runtime.univer.__getInjector().get(IUndoRedoService).clearUndoRedo('projection')
      } finally {
        this.internal--
      }
      const clipboard = this.runtime.univer.__getInjector().get(ISheetClipboardService)
      const copying = mode === 'CUT' ? clipboard.cut() : clipboard.copy()
      const id = clipboard.copyContentCache().getLastCopyId(),
        copy = id ? clipboard.copyContentCache().get(id) : undefined
      const captured =
        id && copy
          ? createCopiedBlock(id, bounds, read, copy.range.rows, copy.range.cols)
          : undefined
      if (!(await copying) || !captured) throw new Error(t('appClipboardUnsupported'))
      this.copiedBlock = captured
      this.copiedSource = {
        sheetId: target.sheetId,
        widths: captured.columns.map((column) => sheet.getSheet().getColumnWidth(column)),
      }
      this.copiedCut =
        mode === 'CUT'
          ? { id: captured.id, sheetId: target.sheetId, address: target.address, ...frozen }
          : undefined
      this.consumedCutId = undefined
      if (this.state?.stateId !== frozen.stateId) clipboard.removeMarkSelection()
    })()
    this.clipboardWork = work
    try {
      await work
    } finally {
      if (this.clipboardWork === work) this.clipboardWork = undefined
    }
  }
  /** Guards for the desktop window shortcut table; null outside this editor. */
  shortcutGuards(event: KeyboardEvent): GlobalShortcutGuards | null {
    const target = event.target
    if (
      this.closed ||
      event.isComposing ||
      this.composing ||
      !(target instanceof Element) ||
      !this.container.closest('.workbook-editor')?.contains(target)
    )
      return null
    return {
      isMac: /Mac/i.test(navigator.platform),
      modalOpen: isModalOpen(this.container.ownerDocument),
      cellEditing: this.runtime.univerAPI.getActiveWorkbook()?.isCellEditing() ?? false,
      gridTarget: isGridKeyTarget(target, this.container),
      formulaBarTarget: isFormulaBarKeyTarget(target),
    }
  }
  /**
   * Home → Insert / Delete Cells: whole rows or columns act at once; any other
   * selection asks for the shift (null), as the desktop's requestCellsAction.
   */
  cellsDecision(): 'entire-row' | 'entire-column' | null {
    const sheet = this.runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
    const range = sheet?.getActiveRange()?.getRange()
    if (!sheet || !range) throw new Error(t('appSelectCellFirst'))
    if (
      range.rangeType === RANGE_TYPE.ROW ||
      range.rangeType === RANGE_TYPE.ALL ||
      (range.startColumn <= 0 && range.endColumn >= sheet.getMaxColumns() - 1)
    )
      return 'entire-row'
    if (
      range.rangeType === RANGE_TYPE.COLUMN ||
      (range.startRow <= 0 && range.endRow >= sheet.getMaxRows() - 1)
    )
      return 'entire-column'
    return null
  }
  /** Inserts or deletes the selected cells; shifts move the rest of the band like a drag-move. */
  async cellsCommand(
    mode: 'insert' | 'delete',
    choice: 'shift-horizontal' | 'shift-vertical' | 'entire-row' | 'entire-column',
  ) {
    if (!(await this.flushEditor())) throw new Error('Finish editing before changing cells.')
    const target = this.ribbonTarget,
      state = this.state
    if (!target || !state) throw new Error(t('appSelectCellFirst'))
    const sheet = state.sheets.find((item) => item.id === target.sheetId)
    if (!sheet) throw new Error(t('appSelectCellFirst'))
    const bounds = parseRange(target.address)
    const height = bounds.endRow - bounds.startRow + 1,
      width = bounds.endColumn - bounds.startColumn + 1
    const sheetId = target.sheetId
    let operations: WorkbookOperation[]
    if (choice === 'entire-row')
      operations = [
        mode === 'insert'
          ? { op: 'insert_rows', sheetId, row: bounds.startRow + 1, count: height }
          : { op: 'delete_rows', sheetId, row: bounds.startRow + 1, count: height },
      ]
    else if (choice === 'entire-column')
      operations = [
        mode === 'insert'
          ? { op: 'insert_cols', sheetId, column: columnLabel(bounds.startColumn), count: width }
          : { op: 'delete_cols', sheetId, column: columnLabel(bounds.startColumn), count: width },
      ]
    else {
      // The band runs from the selection to the used edge of the sheet.
      const vertical = choice === 'shift-vertical'
      const lastRow = Math.max(0, sheet.rows - 1),
        lastColumn = Math.max(0, sheet.columns - 1)
      const area = (startRow: number, startColumn: number, endRow: number, endColumn: number) =>
        `${formatAddress(startRow, startColumn)}:${formatAddress(endRow, endColumn)}`
      if (mode === 'insert') {
        const end = vertical ? lastRow : lastColumn,
          start = vertical ? bounds.startRow : bounds.startColumn
        operations =
          start > end
            ? []
            : [
                {
                  op: 'relocate_range',
                  sheetId,
                  sourceSheetId: sheetId,
                  source: vertical
                    ? area(bounds.startRow, bounds.startColumn, lastRow, bounds.endColumn)
                    : area(bounds.startRow, bounds.startColumn, bounds.endRow, lastColumn),
                  start: vertical
                    ? formatAddress(bounds.startRow + height, bounds.startColumn)
                    : formatAddress(bounds.startRow, bounds.startColumn + width),
                },
              ]
      } else {
        const after = vertical ? bounds.endRow + 1 : bounds.endColumn + 1
        const end = vertical ? lastRow : lastColumn
        operations =
          after > end
            ? [{ op: 'clear_range', sheetId, range: target.address }]
            : [
                {
                  op: 'relocate_range',
                  sheetId,
                  sourceSheetId: sheetId,
                  source: vertical
                    ? area(after, bounds.startColumn, lastRow, bounds.endColumn)
                    : area(bounds.startRow, after, bounds.endRow, lastColumn),
                  start: formatAddress(bounds.startRow, bounds.startColumn),
                },
              ]
      }
    }
    if (operations.length) await this.enqueue(operations)
  }
  /** Home → Fill → Series dialog defaults (direction from the selection shape). */
  get fillSeriesContext() {
    const active = this.runtime.univerAPI.getActiveWorkbook()?.getActiveRange()
    return {
      direction:
        active && active.getHeight() > active.getWidth() ? ('columns' as const) : ('rows' as const),
      date1904: this.state?.dateSystem === '1904',
      numberFormat: this.ribbonTarget?.format?.numberFormat,
    }
  }
  /** Fill → Series: the desktop's planFillSeries written as canonical cells; returns a message or null. */
  async fillSeries(options: FillSeriesOptions): Promise<string | null> {
    if (!(await this.flushEditor())) return 'Finish editing before filling a series.'
    const workbook = this.runtime.univerAPI.getActiveWorkbook()
    const active = workbook?.getActiveRange()
    const state = this.state
    if (!active || !state) return t('appSelectRangeFirst')
    const frozen = { stateId: state.stateId, revision: this.revision }
    const sheetId = active.getSheetId()
    const { startRow, startColumn } = active.getRange()
    const cells = active.getCellDatas()
    const computed = active.getValues()
    const grid = cells.map((row, r) =>
      row.map((cell, c) => {
        const v = cell?.v
        if (typeof v === 'number') return v
        if (typeof v === 'string' && cell?.t === CellValueType.NUMBER && v.trim() !== '') {
          const numeric = Number(v)
          if (Number.isFinite(numeric)) return numeric
        }
        if (cell?.f !== undefined || cell?.si !== undefined) {
          const result = computed[r]?.[c]
          if (typeof result === 'number') return result
        }
        return v ?? null
      }),
    )
    const writes = planFillSeries(grid, options, FILL_SERIES_OPEN_ENDED_LIMIT)
    if (!writes.length) return t('appFillSeriesNothing')
    const operations: WorkbookOperation[] = []
    for (const write of writes) {
      const start = formatAddress(startRow + write.row, startColumn + write.column)
      const height = write.values.length,
        width = write.values[0]?.length ?? 1
      const range = `${start}:${formatAddress(startRow + write.row + height - 1, startColumn + write.column + width - 1)}`
      for (let offset = 0; offset < height; offset += 500)
        operations.push({
          op: 'set_range',
          sheetId,
          start: formatAddress(startRow + write.row + offset, startColumn + write.column),
          values: write.values.slice(offset, offset + 500).map((line) => [...line]),
        })
      // Filled cells take their line's seed number format, so a date series keeps showing dates.
      const seed =
        options.direction === 'rows'
          ? sheetCellFormat(workbook!, sheetId, startRow + write.row, startColumn)
          : sheetCellFormat(workbook!, sheetId, startRow, startColumn + write.column)
      if (seed)
        operations.push({ op: 'format_range', sheetId, range, format: { numberFormat: seed } })
    }
    const cellsWritten = writes.reduce(
      (sum, write) => sum + write.values.length * (write.values[0]?.length ?? 0),
      0,
    )
    if (cellsWritten <= 2000) await this.enqueue(operations, frozen)
    else await this.enqueueStaged([operations], frozen)
    return null
  }
  /**
   * Univer's built-in context-menu commands act on the local model; in the
   * view they become the same canonical edits and dialogs as the ribbon.
   */
  private installContextMenuCommands() {
    const commands = this.runtime.univer.__getInjector().get(ICommandService)
    const sink = (command: string) => () => (this.commandSink?.(command), true)
    const cells =
      (mode: 'insert' | 'delete', choice: Parameters<WorkbookRuntime['cellsCommand']>[1]) => () => {
        void this.cellsCommand(mode, choice).catch(this.report)
        return true
      }
    const insert =
      (axis: 'rows' | 'cols', after: boolean) =>
      (_accessor: unknown, params?: { value?: number }) => {
        void this.insertAxis(axis, after, params?.value).catch(this.report)
        return true
      }
    const handlers: Record<string, (accessor: unknown, params?: { value?: number }) => boolean> = {
      'sheet.command.clear-selection-content': sink('clear-contents'),
      'sheet.command.clear-selection-format': sink('clear-formats'),
      'sheet.command.clear-selection-all': sink('clear-all'),
      'sheet.command.insert-row-before': insert('rows', false),
      'sheet.command.insert-multi-rows-above': insert('rows', false),
      'sheet.command.insert-multi-rows-after': insert('rows', true),
      'sheet.command.insert-col-before': insert('cols', false),
      'sheet.command.insert-multi-cols-before': insert('cols', false),
      'sheet.command.insert-multi-cols-right': insert('cols', true),
      'sheet.command.insert-range-move-right-confirm': cells('insert', 'shift-horizontal'),
      'sheet.command.insert-range-move-down-confirm': cells('insert', 'shift-vertical'),
      'sheet.command.delete-range-move-left-confirm': cells('delete', 'shift-horizontal'),
      'sheet.command.delete-range-move-up-confirm': cells('delete', 'shift-vertical'),
      'sheet.command.remove-row-confirm': cells('delete', 'entire-row'),
      'sheet.command.remove-col-confirm': cells('delete', 'entire-column'),
      'sheet.command.set-selection-frozen': sink('freeze-here'),
      'sheet.command.set-row-frozen': sink('freeze-here'),
      'sheet.command.set-col-frozen': sink('freeze-here'),
      'sheet.command.set-first-row-frozen': sink('freeze-top-row'),
      'sheet.command.set-first-column-frozen': sink('freeze-first-col'),
      'sheet.command.cancel-frozen': sink('unfreeze'),
      'sheet.command.hide-row-confirm': () => this.axisVisibility('rows', true),
      'sheet.command.hide-col-confirm': () => this.axisVisibility('cols', true),
      'sheet.command.set-selected-rows-visible': () => this.axisVisibility('rows', false),
      'sheet.command.set-selected-cols-visible': () => this.axisVisibility('cols', false),
      'sheet.command.set-row-is-auto-height': sink('autofit-row-height'),
      'sheet.command.set-col-auto-width': sink('autofit-col-width'),
    }
    // Filter by the selected cell's value or color: the filter itself is a session edit, so create it
    // through the canonical path first and wait for its projection before the criteria are applied.
    const filterBy = async (
      accessor: Parameters<typeof filterBySelectedCell>[0],
      params?: { by?: Parameters<typeof filterBySelectedCell>[1] },
    ) => {
      if (!params?.by) return false
      const filterService = this.runtime.univer.__getInjector().get(SheetsFilterService)
      // The projection of the new filter reinstalls the workbook, so the target is read again each time.
      const hasModel = () => {
        const workbook = this.runtime.univerAPI.getActiveWorkbook()
        const sheetId = workbook?.getActiveSheet()?.getSheetId()
        return !!workbook && !!sheetId && !!filterService.getFilterModel(workbook.getId(), sheetId)
      }
      try {
        if (!hasModel()) {
          await this.sortFilterCommand('filter-toggle')
          for (let waited = 0; !hasModel() && waited < 5000; waited += 50)
            await new Promise<void>((resolve) => setTimeout(resolve, 50))
          if (!hasModel()) return false
        }
        // The criteria are applied through the same frozen context a header popup opens with.
        const workbook = this.runtime.univerAPI.getActiveWorkbook()
        const sheet = workbook?.getActiveSheet()
        const range = sheet?.getActiveRange()?.getRange()
        if (!workbook || !sheet || !range) return false
        const context = await this.prepareFilterPopup({
          unitId: workbook.getId(),
          subUnitId: sheet.getSheetId(),
          col: range.startColumn,
        } as IOpenFilterPanelOperationParams)
        this.filterPopupContext = context
        try {
          return await filterBySelectedCell(accessor, params.by)
        } finally {
          this.filterPopupContext = undefined
        }
      } catch (error) {
        this.report(error)
        return false
      }
    }
    commands.unregisterCommand(FILTER_BY_COMMAND_ID)
    this.disposables.push(
      commands.registerCommand({
        id: FILTER_BY_COMMAND_ID,
        type: CommandType.COMMAND,
        handler: filterBy as (accessor: never, params?: never) => Promise<boolean>,
      }),
    )
    for (const [id, handler] of Object.entries(handlers)) {
      commands.unregisterCommand(id)
      this.disposables.push(
        commands.registerCommand({
          id,
          type: CommandType.COMMAND,
          handler: handler as (accessor: never, params?: never) => boolean,
        }),
      )
    }
  }
  /** Context menu Insert ▸ rows/columns above, below, left or right of the selection. */
  private async insertAxis(axis: 'rows' | 'cols', after: boolean, count?: number) {
    if (!(await this.flushEditor())) throw new Error('Finish editing before inserting.')
    const target = this.ribbonTarget
    if (!target) throw new Error(t('appSelectCellFirst'))
    const bounds = parseRange(target.address)
    const size =
      count ??
      (axis === 'rows'
        ? bounds.endRow - bounds.startRow + 1
        : bounds.endColumn - bounds.startColumn + 1)
    const at =
      axis === 'rows'
        ? after
          ? bounds.endRow + 1
          : bounds.startRow
        : after
          ? bounds.endColumn + 1
          : bounds.startColumn
    await this.enqueue([
      axis === 'rows'
        ? { op: 'insert_rows', sheetId: target.sheetId, row: at + 1, count: size }
        : { op: 'insert_cols', sheetId: target.sheetId, column: columnLabel(at), count: size },
    ])
  }
  /** Ctrl+9 / Ctrl+0 (+Shift): hide or show the selected rows or columns, as Excel. */
  private axisVisibility(axis: 'rows' | 'cols', hidden: boolean) {
    const target = this.ribbonTarget
    const sheet = this.runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
    if (!target || !sheet) return false
    const bounds = parseRange(target.address)
    const start = axis === 'rows' ? bounds.startRow : bounds.startColumn
    const end = axis === 'rows' ? bounds.endRow : bounds.endColumn
    // Excel refuses to hide every row or column.
    if (
      hidden &&
      start === 0 &&
      end >= (axis === 'rows' ? sheet.getMaxRows() : sheet.getMaxColumns()) - 1
    )
      return false
    const operation: WorkbookOperation =
      axis === 'rows'
        ? {
            op: 'set_rows_hidden',
            sheetId: target.sheetId,
            row: start + 1,
            count: end - start + 1,
            hidden,
          }
        : {
            op: 'set_cols_hidden',
            sheetId: target.sheetId,
            column: columnLabel(start),
            count: end - start + 1,
            hidden,
          }
    void this.enqueue([operation]).catch(this.report)
    return true
  }
  /** Shift+F11: a new SheetN at the end, as the desktop's insert-sheet. */
  async insertSheet() {
    if (!(await this.flushEditor())) throw new Error('Finish editing before inserting a sheet.')
    const state = this.state
    if (!state) throw new Error(t('appOpenXlsxFirst'))
    await this.enqueue([
      { op: 'add_sheet', name: nextSheetName(state.sheets.map((sheet) => sheet.name)) },
    ])
    return t('appSheetAdded')
  }
  /** Ctrl+; / Ctrl+Shift+;: the static current date or time into the active cell. */
  async insertNow(kind: 'date' | 'time') {
    if (!(await this.flushEditor())) throw new Error('Finish editing before inserting the date.')
    const state = this.state,
      sheet = this.runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
    const current = sheet?.getSelection()?.getCurrentCell()
    const target = this.ribbonTarget
    if (!state || !sheet || !target) throw new Error(t('appSelectRangeFirst'))
    const origin = parseRange(target.address)
    const address = formatAddress(
      current?.actualRow ?? origin.startRow,
      current?.actualColumn ?? origin.startColumn,
    )
    const now = new Date()
    const serial =
      (Date.UTC(
        now.getFullYear(),
        now.getMonth(),
        now.getDate(),
        now.getHours(),
        now.getMinutes(),
        now.getSeconds(),
      ) -
        Date.UTC(1899, 11, 30)) /
      86400000
    const value =
      kind === 'time'
        ? serial - Math.floor(serial)
        : Math.floor(serial) - (state.dateSystem === '1904' ? 1462 : 0)
    await this.enqueue([
      { op: 'set_cell', sheetId: target.sheetId, address, value },
      {
        op: 'format_range',
        sheetId: target.sheetId,
        range: address,
        format: { numberFormat: kind === 'time' ? 'h:mm' : 'yyyy/m/d' },
      },
    ])
  }
  /**
   * Excel's PageUp/PageDown (Alt+ = one screen left/right): move the active
   * cell by one viewport and scroll with it, keeping its on-screen position.
   */
  pageMove(horizontal: boolean, direction: 1 | -1) {
    const workbook = this.runtime.univerAPI.getActiveWorkbook()
    const sheet = workbook?.getActiveSheet()
    const active = workbook?.getActiveRange()
    if (!workbook || !sheet || !active || workbook.isCellEditing()) return
    let visible: { startRow: number; endRow: number; startColumn: number; endColumn: number } | null
    try {
      visible = sheet.getVisibleRange()
    } catch {
      return
    }
    if (!visible) return
    const page = horizontal
      ? Math.max(1, visible.endColumn - visible.startColumn)
      : Math.max(1, visible.endRow - visible.startRow)
    const clamp = (value: number, max: number) => Math.min(max, Math.max(0, value))
    const maxRow = sheet.getMaxRows() - 1,
      maxColumn = sheet.getMaxColumns() - 1
    const row = horizontal ? active.getRow() : clamp(active.getRow() + direction * page, maxRow)
    const column = horizontal
      ? clamp(active.getColumn() + direction * page, maxColumn)
      : active.getColumn()
    const rtl = sheet.getSheet().getConfig().rightToLeft === BooleanNumber.TRUE
    const anchor = getScrollAnchor(workbook, sheet) ?? {
      row: visible.startRow,
      column: rtl ? visible.endColumn : visible.startColumn,
    }
    const anchorColumn =
      rtl && horizontal ? Math.max(anchor.column, visible.endColumn) : anchor.column
    workbook.setActiveRange(sheet.getRange(row, column))
    sheet.scrollToCell(
      horizontal ? anchor.row : clamp(anchor.row + direction * page, maxRow),
      horizontal ? clamp(anchorColumn + direction * page, maxColumn) : anchor.column,
    )
  }
  /** The ribbon's command dispatcher, for Univer-side shortcuts that open shared dialogs. */
  commandSink?: (command: string) => void
  private painter: { mode: 'once' | 'lock'; snapshot: SpecialPasteSnapshot } | null = null
  get formatPainterActive() {
    return this.painter !== null
  }
  /**
   * Home → Format Painter: a click copies the selection's formats for one
   * paint (a second click cancels), a double-click keeps painting until Esc
   * or another click, as the desktop's painter.
   */
  async formatPainter(detail: number) {
    if (detail >= 2) {
      if (this.painter?.mode === 'once') {
        this.painter = { ...this.painter, mode: 'lock' }
        this.changed()
      }
      return
    }
    if (this.painter) {
      this.painter = null
      this.changed()
      return
    }
    if (!(await this.flushEditor())) throw new Error(t('appClipboardBusy'))
    const target = this.ribbonTarget,
      state = this.state
    if (!target || !state || this.closed) throw new Error(t('appSelectRangeFirst'))
    const bounds = parseRange(target.address)
    if (rangeCellCount(bounds) > 2000) throw new Error(t('appClipboardUnsupported'))
    const read = (await this.host.read(target.sheetId, bounds, state.stateId)) as {
      cells: ViewCell[]
      merges?: GridRange[]
    }
    const rows = Array.from(
        { length: bounds.endRow - bounds.startRow + 1 },
        (_, r) => bounds.startRow + r,
      ),
      columns = Array.from(
        { length: bounds.endColumn - bounds.startColumn + 1 },
        (_, c) => bounds.startColumn + c,
      )
    const block = createCopiedBlock(`painter-${Date.now()}`, bounds, read, rows, columns)
    this.painter = {
      mode: 'once',
      snapshot: {
        source: {
          kind: 'internal',
          unitId: 'projection',
          subUnitId: target.sheetId,
          rows,
          cols: columns,
          cut: false,
        },
        block,
        sheetId: target.sheetId,
        html: '',
        text: '',
      },
    }
    this.changed()
  }
  private async paintFormats() {
    const painter = this.painter,
      target = this.ribbonTarget,
      state = this.state
    if (!painter || !target || !state) return
    if (painter.mode === 'once') {
      this.painter = null
      this.changed()
    }
    await this.applySpecialPaste(
      {
        target: { sheetId: target.sheetId, address: target.address },
        stateId: state.stateId,
        revision: this.revision,
        snapshot: painter.snapshot,
      },
      { ...defaultPasteSpecial, type: 'formats' },
    )
  }
  pasteSpecialDialog: SpecialPasteContext | null = null
  async openPasteSpecial() {
    this.pasteSpecialDialog = await this.preparePasteSpecial()
    this.changed()
  }
  closePasteSpecial() {
    this.pasteSpecialDialog = null
    this.changed()
  }
  async preparePasteSpecial(): Promise<SpecialPasteContext> {
    if (this.clipboardWork || !(await this.flushEditor())) throw new Error(t('appClipboardBusy'))
    const target = this.ribbonTarget
    if (!target || !this.state || this.closed) throw new Error('missing_selection')
    const context = {
      target: { sheetId: target.sheetId, address: target.address },
      stateId: this.state.stateId,
      revision: this.revision,
    }
    let snapshot: SpecialPasteSnapshot = { source: { kind: 'none' }, html: '', text: '' }
    const work = (async () => {
      let html = '',
        text = ''
      try {
        if (navigator.clipboard?.read) {
          const item = (await navigator.clipboard.read())[0]
          if (item?.types.includes('text/html'))
            html = await (await item.getType('text/html')).text()
          if (item?.types.includes('text/plain'))
            text = await (await item.getType('text/plain')).text()
        } else text = await navigator.clipboard.readText()
      } catch {
        /* Desktop falls back to the last internal copy when clipboard reads are denied. */
      }
      snapshot.html = html
      snapshot.text = text
      const block = this.copiedBlock
      if (block && (clipboardCopyId(html) === block.id || (!html && !text))) {
        const owner = this.copiedSource
        const currentSource = owner
          ? this.runtime.univerAPI.getActiveWorkbook()?.getSheetBySheetId(owner.sheetId)
          : undefined
        const widths = currentSource
          ? block.columns.map((column) => currentSource.getSheet().getColumnWidth(column))
          : owner?.widths
        snapshot = {
          source: {
            kind: 'internal',
            unitId: 'projection',
            subUnitId: owner?.sheetId ?? '',
            rows: block.rows,
            cols: block.columns,
            cut: this.copiedCut?.id === block.id,
          },
          block,
          ...(owner ? { sheetId: owner.sheetId, ...(widths ? { widths } : {}) } : {}),
          html: html || `<table data-copy-id="${block.id}">`,
          text,
        }
      } else if (text.trim())
        snapshot = { source: { kind: 'text', rows: parseClipboardText(text) }, html, text }
      if (
        this.closed ||
        this.state?.stateId !== context.stateId ||
        this.revision !== context.revision
      )
        throw new Error(t('appClipboardSelectionChanged'))
    })()
    this.clipboardWork = work
    try {
      await work
    } finally {
      if (this.clipboardWork === work) this.clipboardWork = undefined
    }
    return { ...context, snapshot }
  }
  async pasteSpecial(type: PasteSpecialType) {
    await this.applySpecialPaste(await this.preparePasteSpecial(), { ...defaultPasteSpecial, type })
  }
  async applySpecialPaste(context: SpecialPasteContext, options: PasteSpecialOptions) {
    if (this.visualWork || this.chartWork || this.symbolWork) throw new Error(t('appClipboardBusy'))
    const { snapshot, target } = context
    if (snapshot.source.kind === 'internal' && snapshot.source.cut) {
      if (
        this.copiedCut?.id !== snapshot.block?.id ||
        options.type !== 'all' ||
        options.link ||
        options.transpose ||
        options.skipBlanks ||
        options.operation !== 'none'
      )
        throw new Error(t('appClipboardUnsupported'))
      await this.pasteNativeCells(snapshot.html, snapshot.text, { ...context, target })
      return
    }
    const work = (async () => {
      if (!(await this.flushCellEditor())) throw new Error(t('appClipboardBusy'))
      if (
        !snapshot.block &&
        (snapshot.html || snapshot.text) &&
        options.operation === 'none' &&
        !options.skipBlanks &&
        !options.transpose &&
        !options.link
      ) {
        const nativeTypes: Partial<Record<PasteSpecialType, string>> = {
          all: 'default-paste',
          'all-source-theme': 'default-paste',
          'all-merge-cf': 'default-paste',
          values: 'special-paste-value',
          formulas: 'special-paste-formula',
          formats: 'special-paste-format',
          'col-widths': 'special-paste-col-width',
          'all-except-borders': 'special-paste-besides-border',
        }
        const type = nativeTypes[options.type]
        if (type) {
          const point = parseRange(target.address)
          const operations = await planNativeCellPaste(
            this.runtime,
            snapshot.html,
            snapshot.text,
            { sheetId: target.sheetId, row: point.startRow, column: point.startColumn },
            type,
          )
          await this.enqueue(operations, context, { failureInDialog: true })
          return
        }
      }
      let resolved = snapshot
      const advanced = needsOwnLayer(options)
      const carriesSourceFormat = [
        'all',
        'all-source-theme',
        'all-merge-cf',
        'formats',
        'all-except-borders',
      ].includes(options.type)
      if (snapshot.block && snapshot.sheetId && (advanced || carriesSourceFormat)) {
        if (!this.state?.sheets.some((sheet) => sheet.id === snapshot.sheetId))
          throw new Error(t('appClipboardUnsupported'))
        const block = snapshot.block
        const sourceBounds = {
          startRow: block.rows[0]!,
          endRow: block.rows[block.rows.length - 1]!,
          startColumn: block.columns[0]!,
          endColumn: block.columns[block.columns.length - 1]!,
        }
        const read = (await this.host.read(snapshot.sheetId, sourceBounds, context.stateId)) as {
          cells: ViewCell[]
          merges?: GridRange[]
        }
        const fresh = createCopiedBlock(block.id, sourceBounds, read, block.rows, block.columns)
        if (advanced) resolved = { ...snapshot, block: fresh }
        else
          resolved = {
            ...snapshot,
            block: {
              ...block,
              cells: block.cells.map((row, r) =>
                row.map((cell, c) => {
                  const value = structuredClone(cell),
                    format = fresh.cells[r]![c]!
                  delete value.style
                  delete value.numberFormat
                  if (format.style) value.style = format.style
                  if (format.numberFormat) value.numberFormat = format.numberFormat
                  return value
                }),
              ),
            },
          }
      }
      const bounds = specialPasteBounds(resolved, options, target.address)
      const read = (await this.host.read(target.sheetId, bounds, context.stateId)) as {
        cells: ViewCell[]
        merges?: GridRange[]
      }
      let operations: WorkbookOperation[]
      if (
        snapshot.block &&
        ['all', 'all-source-theme', 'all-merge-cf'].includes(options.type) &&
        options.operation === 'none' &&
        !options.skipBlanks &&
        !options.transpose &&
        !options.link
      )
        operations = [
          copiedBlockOperation(resolved.block!, target.sheetId, target.address, read.merges),
        ]
      else {
        const source = snapshot.sheetId
          ? this.state?.sheets.find((sheet) => sheet.id === snapshot.sheetId)
          : undefined
        if (options.link && !source) throw new Error(t('appClipboardUnsupported'))
        operations = specialPasteOperations(resolved, options, {
          sheetId: target.sheetId,
          address: target.address,
          cells: read.cells,
          linkSheetName: source && source.id !== target.sheetId ? source.name : null,
        })
      }
      if (operations.length) await this.enqueue(operations, context, { failureInDialog: true })
    })()
    this.visualWork = work
    try {
      await work
    } finally {
      if (this.visualWork === work) this.visualWork = undefined
    }
  }
  private clipboardWork: Promise<void> | undefined
  async pasteClipboard() {
    if (this.clipboardWork) throw new Error(t('appClipboardBusy'))
    if (!(await this.flushEditor())) throw new Error(t('appClipboardBusy'))
    if (this.clipboardWork) throw new Error(t('appClipboardBusy'))
    const target = this.ribbonTarget
    if (!target || !this.state || this.closed) throw new Error('missing_selection')
    const frozen = { target, stateId: this.state.stateId, revision: this.revision }
    const work = (async () => {
      let html = '',
        text = '',
        file: File | null = null
      try {
        const clipboard = navigator.clipboard
        if (clipboard.read) {
          const item = (await clipboard.read())[0]
          if (!item) return
          if (item.types.includes('text/html'))
            html = await (await item.getType('text/html')).text()
          if (item.types.includes('text/plain'))
            text = await (await item.getType('text/plain')).text()
          const mime = item.types.find((type) => type.startsWith('image/'))
          if (mime)
            file = new File([await item.getType(mime)], `clipboard-image.${mime.split('/')[1]}`, {
              type: mime,
            })
        } else text = await clipboard.readText()
      } catch (error) {
        throw new Error(
          t('appClipboardReadUnavailable', {
            shortcut: /Mac/i.test(navigator.platform) ? '⌘V' : 'Ctrl+V',
          }),
          { cause: error },
        )
      }
      const picture = pickTransferImage({
        items: file ? [{ kind: 'file', type: file.type, getAsFile: () => file }] : [],
        getData: (format) => (format === 'text/html' ? html : format === 'text/plain' ? text : ''),
      })
      if (picture) await this.insertPictureFile(picture, target, frozen)
      else if (html || text) await this.pasteNativeCells(html, text, frozen)
      else throw new Error(t('appClipboardUnsupported'))
    })()
    this.clipboardWork = work
    try {
      await work
    } finally {
      if (this.clipboardWork === work) this.clipboardWork = undefined
    }
  }
  private async pasteNativeCells(
    html: string,
    text: string,
    context?: {
      target: Pick<RibbonTarget, 'sheetId' | 'address'>
      stateId: string
      revision: number
    },
  ) {
    const target = context?.target ?? this.ribbonTarget
    if (!target || !this.state || this.closed) return
    if (this.visualWork || this.chartWork || this.symbolWork) throw new Error(t('appClipboardBusy'))
    const frozen = context ?? { stateId: this.state.stateId, revision: this.revision },
      point = parseAddress(target.address.split(':')[0]!)
    const work = (async () => {
      if (!(await this.flushCellEditor())) throw new Error(t('appClipboardBusy'))
      const copyId = clipboardCopyId(html)
      if (copyId && copyId === this.consumedCutId) throw new Error(t('appClipboardUnsupported'))
      const cut = this.copiedCut?.id === copyId ? this.copiedCut : undefined
      if (cut) {
        if (frozen.stateId !== cut.stateId || frozen.revision !== cut.revision)
          throw new Error(t('appClipboardSelectionChanged'))
        try {
          await this.enqueue(
            [
              {
                op: 'relocate_range',
                sheetId: target.sheetId,
                sourceSheetId: cut.sheetId,
                source: cut.address,
                start: formatAddress(point.row, point.column),
              },
            ],
            frozen,
            { failureInDialog: true },
          )
        } catch (error) {
          if (String(error).includes('unsupported_move_feature:'))
            throw new Error(t('appClipboardUnsupported'), { cause: error })
          throw error
        }
        this.copiedCut = undefined
        this.copiedBlock = undefined
        this.consumedCutId = cut.id
        const clipboard = this.runtime.univer.__getInjector().get(ISheetClipboardService)
        clipboard.copyContentCache().del(cut.id)
        clipboard.removeMarkSelection()
        return
      }
      const copied = this.copiedBlock?.id === copyId ? this.copiedBlock : undefined
      if (copied) {
        const bounds = parseRange(target.address)
        const read = (await this.host.read(target.sheetId, bounds, frozen.stateId)) as {
          merges?: GridRange[]
        }
        await this.enqueue(
          [copiedBlockOperation(copied, target.sheetId, target.address, read.merges)],
          frozen,
          { failureInDialog: true },
        )
        return
      }
      const operations = await planNativeCellPaste(this.runtime, html, text, {
        sheetId: target.sheetId,
        row: point.row,
        column: point.column,
      })
      await this.enqueue(operations, frozen, { failureInDialog: true })
    })()
    this.visualWork = work
    try {
      await work
    } finally {
      if (this.visualWork === work) this.visualWork = undefined
    }
  }
  private insertTransferredPicture(file: File, anchor: PictureAnchor | null) {
    const target = this.ribbonTarget
    if (!target || !this.state || this.closed) return
    const frozen = { stateId: this.state.stateId, revision: this.revision }
    void this.insertPictureFile(
      file,
      {
        sheetId: target.sheetId,
        address: anchor ? formatAddress(anchor.row, anchor.column) : target.address,
      },
      frozen,
    ).catch(this.report)
  }
  private shapeDrawDispose: (() => void) | undefined
  async startShapeDrawing(shapeType: string) {
    if (shapeType !== 'textbox' && !ADDABLE_SHAPE_TYPES.some((type) => type === shapeType))
      throw new Error('unsupported_shape_type')
    if (!(await this.flushEditor())) throw new Error('Finish editing before drawing a shape.')
    if (!this.state || this.closed) throw new Error('view_unavailable')
    const sheet = this.runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
    if (!sheet) throw new Error('missing_selection')
    const sheetId = sheet.getSheetId(),
      frozen = { stateId: this.state.stateId, revision: this.revision }
    this.shapeDrawDispose?.()
    if (shapeType === 'textbox') {
      const anchorCell = this.ribbonTarget?.address.split(':')[0] ?? 'A1'
      const work = this.enqueue(
        [workbookOperationSchema.parse({ op: 'add_shape', sheetId, shapeType, anchorCell })],
        frozen,
        { failureInDialog: true },
      )
      this.visualWork = work
      try {
        await work
      } finally {
        if (this.visualWork === work) this.visualWork = undefined
      }
      return
    }
    this.shapeDrawDispose = startSheetShapeDraw(
      this.runtime,
      shapeType === 'textbox' ? 'rect' : shapeType,
      (anchor) => {
        const work = (async () => {
          const receipt = await this.enqueue(
            [
              workbookOperationSchema.parse({
                op: 'add_shape',
                sheetId,
                shapeType,
                anchorCell: formatAddress(anchor.fromRow, anchor.fromColumn),
                anchor,
              }),
            ],
            frozen,
            { failureInDialog: true },
          )
          return receipt
        })()
        this.visualWork = work
        const clear = () => {
          if (this.visualWork === work) this.visualWork = undefined
        }
        void work.then(clear, (error) => {
          clear()
          this.report(error)
        })
      },
      this.container,
    )
  }
  isSessionVisual(id: string) {
    return (
      this.state?.sheets.some((sheet) =>
        (
          (sheet.features?.visuals ?? sheet.features?.charts ?? []) as unknown as SheetVisual[]
        ).some((v) => v.id === id),
      ) ?? false
    )
  }
  get shapeRibbonTarget() {
    const visual = this.visuals.find(
      (v) =>
        v.id === selectedVisualIds().at(-1) &&
        (v.kind === 'shape' || v.kind === 'image') &&
        (isEditableFileVisual(v) || this.isSessionVisual(v.id)),
    )
    return visual ?? null
  }
  async shapeRibbonCommand(command: string): Promise<void> {
    if (!(await this.flushEditor())) throw new Error('Finish editing before changing a shape.')
    const visual = this.shapeRibbonTarget
    if (!visual || !this.state) return Promise.reject(new Error('unknown_visual'))
    if (command.startsWith('visual:')) return this.visualCommand(command, visual.id)
    if (command === 'shape-delete') return this.visualCommand('visual:delete', visual.id)
    const paint = /^shape-(fill|outline):(#[0-9A-Fa-f]{6}|none)$/.exec(command)
    if (paint && visual.kind === 'shape')
      return this.editVisual(
        visual.sheetId,
        visual.id,
        paint[1] === 'fill' ? { fillColor: paint[2]! } : { lineColor: paint[2]! },
        { stateId: this.state.stateId, revision: this.revision },
      ).then(() => {})
    return Promise.reject(new Error('unsupported_visual_command'))
  }
  private installedFrame(visual: WorkbookVisualObject) {
    return installedVisualFrames(visual.sheetId).find((frame) => frame.visual.id === visual.id)
  }
  private visualFrameSize(visual: WorkbookVisualObject) {
    const frame = this.installedFrame(visual)
    if (!frame) return null
    return visual.rotation && visual.frameWidth && visual.frameHeight
      ? { width: visual.frameWidth / EMU_PER_PIXEL, height: visual.frameHeight / EMU_PER_PIXEL }
      : { width: frame.width, height: frame.height }
  }
  private frameEdit(
    visual: WorkbookVisualObject,
    size: { width: number; height: number; rotation: number },
  ) {
    const frame = this.installedFrame(visual),
      sheet = this.runtime.univerAPI.getActiveWorkbook()?.getSheetBySheetId(visual.sheetId)
    if (!frame || !sheet) throw new Error('visual_frame_unavailable')
    const metrics = sheetMetrics(sheet)
    return frameChangeFromBox(frameBox(frame, metrics), size, metrics)
  }
  private visualDialog: {
    kind: 'alt-text' | 'size'
    visual: WorkbookVisualObject
    frame: { width: number; height: number } | null
    stateId: string
    revision: number
  } | null = null
  get visualDialogTarget() {
    return this.visualDialog
  }
  closeVisualDialog() {
    this.visualDialog = null
    this.changed()
  }
  editVisualDialog(changes: ShapeEditChanges | SizeDialogResult): Promise<unknown> {
    const target = this.visualDialog
    if (!target) return Promise.reject(new Error('unknown_visual'))
    if ('size' in changes && changes.size) {
      try {
        return this.editVisual(
          target.visual.sheetId,
          target.visual.id,
          {
            ...this.frameEdit(target.visual, changes.size),
            ...(changes.editAs ? { editAs: changes.editAs } : {}),
          },
          target,
        )
      } catch (reason) {
        return Promise.reject(reason)
      }
    }
    return this.editVisual(
      target.visual.sheetId,
      target.visual.id,
      changes as ShapeEditChanges,
      target,
    )
  }
  private async visualCommand(command: string, id: string): Promise<void> {
    const visual = this.visuals.find((v) => v.id === id)
    if (!visual || !this.state) throw new Error('unknown_visual')
    const frozen = { stateId: this.state.stateId, revision: this.revision }
    if (command === 'visual:alt-text' || command === 'visual:size') {
      this.visualDialog = {
        kind: command === 'visual:alt-text' ? 'alt-text' : 'size',
        visual: structuredClone({
          ...visual,
          ...(this.isSessionChart(id) ? { editAs: visual.editAs ?? 'twoCell' } : {}),
        }),
        frame: this.visualFrameSize(visual),
        ...frozen,
      }
      this.changed()
      return
    }
    if (command === 'visual:delete') {
      const ids = selectedVisualIds()
      await this.editVisual(visual.sheetId, ids.length ? ids : [id], { remove: true }, frozen)
      clearVisualSelection()
      return
    }
    const rotate = /^visual:rotate:(cw|ccw)$/.exec(command)
    if (rotate && visual.kind !== 'chart') {
      const size = this.visualFrameSize(visual)
      if (!size) throw new Error('visual_frame_unavailable')
      await this.editVisual(
        visual.sheetId,
        id,
        this.frameEdit(visual, {
          ...size,
          rotation: (visual.rotation ?? 0) + (rotate[1] === 'cw' ? 90 : -90),
        }),
        frozen,
      )
      return
    }
    const flip = /^visual:flip:(h|v)$/.exec(command)
    if (flip && visual.kind !== 'chart') {
      await this.editVisual(
        visual.sheetId,
        id,
        flip[1] === 'h' ? { flipH: !visual.flipH } : { flipV: !visual.flipV },
        frozen,
      )
      return
    }
    const arrange =
      /^visual:(align|distribute):(left|center|right|top|middle|bottom|horizontal|vertical)$/.exec(
        command,
      )
    if (arrange) {
      const ids = selectedVisualIds(),
        minimum = arrange[1] === 'distribute' ? 3 : 2
      if (ids.length < minimum)
        throw new Error(t(minimum === 3 ? 'appVisualSelectThree' : 'appVisualSelectTwo'))
      const frames = installedVisualFrames(visual.sheetId).filter((frame) =>
        ids.includes(frame.visual.id),
      )
      if (frames.length !== ids.length || frames.length > 100)
        throw new Error('invalid_visual_selection')
      const sheet = this.runtime.univerAPI.getActiveWorkbook()!.getSheetBySheetId(visual.sheetId)!,
        metrics = sheetMetrics(sheet)
      const boxes = frames.map((frame) => frameBox(frame, metrics)),
        moves =
          arrange[1] === 'align'
            ? alignBoxes(boxes, arrange[2] as AlignMode)
            : distributeBoxes(boxes, arrange[2] as DistributeAxis)
      const operations = moves.map((move) => ({
        op: 'set_visual_anchor' as const,
        sheetId: visual.sheetId,
        visualId: move.id,
        anchor: anchorForBox(
          { ...boxes.find((box) => box.id === move.id)!, x: move.x, y: move.y },
          metrics,
        ),
      }))
      if (!operations.length) return
      if (this.visualWork || this.chartWork)
        throw new Error('Wait for the current visual edit to finish.')
      const work = (async () => {
        if (!(await this.flushCellEditor()))
          throw new Error('Finish editing before arranging visuals.')
        await this.enqueue(operations, frozen, { failureInDialog: true })
      })()
      this.visualWork = work
      try {
        await work
      } finally {
        if (this.visualWork === work) this.visualWork = undefined
      }
      return
    }
    const match = /^visual:z:(front|back|forward|backward)$/.exec(command)
    if (!match) throw new Error(`Visual command is not connected: ${command}`)
    const source = this.currentFileVisuals().find((v) => v.id === id)
    const part = source
      ? this.currentFileVisuals().filter(
          (v) => v.drawingPath === source.drawingPath && v.drawingIndex !== undefined,
        )
      : this.visuals.filter((v) => this.isSessionVisual(v.id))
    const order = part.map((v) => v.id),
      after = reorderZ(order, id, match[1] as ZOrderOp)
    if (order.join() === after.join()) return
    if (this.visualWork || this.chartWork)
      throw new Error('Wait for the current visual edit to finish.')
    const work = (async () => {
      if (!(await this.flushCellEditor()))
        throw new Error('Finish editing before changing a chart.')
      await this.enqueue(
        [
          workbookOperationSchema.parse({
            op: 'edit_visual',
            visualId: id,
            zIndex: after.indexOf(id),
          }),
        ],
        frozen,
        { failureInDialog: true },
      )
    })()
    this.visualWork = work
    try {
      await work
    } finally {
      if (this.visualWork === work) this.visualWork = undefined
    }
  }
  private editVisual(
    sheetId: string,
    id: string | readonly string[],
    changes: ShapeEditChanges,
    frozen: { stateId: string; revision: number },
  ) {
    if (this.visualWork)
      return Promise.reject(new Error('Wait for the current visual edit to finish.'))
    const work = (async () => {
      if (!(await this.flushCellEditor()))
        throw new Error('Finish editing before changing a chart.')
      if (
        this.closed ||
        this.state?.stateId !== frozen.stateId ||
        this.revision !== frozen.revision
      )
        throw new Error('The document changed during the visual edit. Try again.')
      const ids = typeof id === 'string' ? [id] : [...id]
      if (ids.length === 0 || ids.length > 100) throw new Error('Select up to 100 visuals.')
      if (ids.length > 1 && !changes.remove) throw new Error('Select one visual before moving it.')
      const operations = ids.flatMap((visualId) => {
        if (changes.remove)
          return [workbookOperationSchema.parse({ op: 'delete_visual', visualId })]
        const metadata = Object.fromEntries(
          Object.entries(changes).filter(
            ([key, value]) => !['anchor', 'remove', 'text'].includes(key) && value !== undefined,
          ),
        )
        return [
          ...(changes.text !== undefined
            ? [workbookOperationSchema.parse({ op: 'edit_shape', visualId, text: changes.text })]
            : []),
          ...(changes.anchor
            ? [
                workbookOperationSchema.parse({
                  op: 'set_visual_anchor',
                  sheetId,
                  visualId,
                  anchor: changes.anchor,
                }),
              ]
            : []),
          ...(Object.keys(metadata).length
            ? [workbookOperationSchema.parse({ op: 'edit_visual', visualId, ...metadata })]
            : []),
        ]
      })
      if (!operations.length) return
      return this.enqueue(operations, frozen, { failureInDialog: true })
    })()
    this.visualWork = work
    const clear = () => {
      if (this.visualWork === work) this.visualWork = undefined
    }
    void work.then(clear, clear)
    return work
  }
  get chartRibbonTarget() {
    const id = selectedVisualIds().at(-1)
    const visual = this.charts.find((v) => v.id === id)
    if (!visual) return null
    const chart = visual.chart
    const kind = convertibleType(chart)
    const isPie = chart.chartTypes.some((type) => /pie|doughnut/i.test(type))
    const owner = this.state?.sheets.find((sheet) => sheet.id === visual.sheetId)
    const echo: SelectedChartRibbon = {
      title: chart.title,
      convertible: kind,
      currentType: kind,
      canEdit: !!owner && !owner.protected && !owner.readOnly,
      isPie,
      hasAxes: kind !== null && !isPie,
      canLabel: chartSupportsDataLabels(chart.chartTypes),
      seriesCount: chart.series.length,
      categoryCount: chart.series[0]?.categories.length ?? 0,
      series: chart.series,
      legend: chart.legend,
      axisTitles: chart.axisTitles,
      dataLabels: chart.dataLabels,
      grouping: chart.grouping,
    }
    return { visual, echo }
  }
  private chartText: {
    target: ChartTextTarget
    initial: string
    key: string
    stateId: string
    revision: number
  } | null = null
  get chartTextTarget() {
    return this.chartText
  }
  closeChartText() {
    this.chartText = null
    this.changed()
  }
  editChartText(command: string) {
    const target = this.chartText
    const edit = chartRibbonEdit(command, this.chartRibbonTarget?.echo ?? null)
    if (!target || !edit) return Promise.reject(new Error(t('appChartNotEditable')))
    return this.editInlineChart(target.key, edit, target)
  }
  openChartDialog(id: string, kind: 'format' | 'select-data') {
    if (kind === 'select-data') {
      const visual = this.charts.find((v) => v.id === id || v.chartPath === id)
      if (!visual || !this.state) return
      this.chartSelection = {
        key: id,
        visual: structuredClone(visual),
        stateId: this.state.stateId,
        revision: this.revision,
      }
    } else this.chartFormatKey = id
    this.changed()
  }
  chartRibbonCommand(command: string): Promise<void> {
    const target = this.chartRibbonTarget
    if (!target || !this.state) return Promise.reject(new Error(t('appChartNotEditable')))
    const { visual, echo } = target
    const key = visual.chartPath ?? visual.id
    const frozen = { stateId: this.state.stateId, revision: this.revision }
    const textTargets = {
      'chart-element-title': 'title',
      'chart-element-axis-cat': 'axis-category',
      'chart-element-axis-val': 'axis-value',
    } as const
    const textTarget = textTargets[command as keyof typeof textTargets]
    if (textTarget) {
      this.chartText = {
        target: textTarget,
        initial:
          textTarget === 'title'
            ? echo.title
            : ((textTarget === 'axis-category'
                ? echo.axisTitles?.category
                : echo.axisTitles?.value) ?? ''),
        key,
        ...frozen,
      }
      this.changed()
      return Promise.resolve()
    }
    if (command === 'chart-select-data' || command === 'chart-format-pane') {
      this.openChartDialog(key, command === 'chart-select-data' ? 'select-data' : 'format')
      return Promise.resolve()
    }
    if (command === 'chart-delete')
      return this.editVisual(visual.sheetId, visual.id, { remove: true }, frozen).then(() => {
        clearVisualSelection()
      })
    if (command === 'chart-switch-row-col')
      return this.applyChartSelection(
        key,
        {
          mode: 'transpose',
          defaultNames: Array.from({ length: 24 }, (_, i) => t('appSeriesN', { n: i + 1 })),
        },
        frozen,
      )
    const edit = chartRibbonEdit(command, echo)
    return edit
      ? this.editInlineChart(key, edit, frozen)
      : Promise.reject(new Error(t('appChartNotEditable')))
  }
  private chartSelection: {
    key: string
    visual: SheetChartVisual
    stateId: string
    revision: number
  } | null = null
  get chartSelectionTarget() {
    return this.chartSelection
  }
  closeChartSelection() {
    this.chartSelection = null
    this.changed()
  }
  editChartSelection(intent: ChartSeriesSelection): Promise<void> {
    const target = this.chartSelection
    if (!target) return Promise.reject(new Error(t('appChartNotEditable')))
    return this.applyChartSelection(target.key, intent, target)
  }
  private applyChartSelection(
    key: string,
    intent: ChartSeriesSelection,
    frozen: { stateId: string; revision: number },
  ): Promise<void> {
    if (this.chartWork || this.visualWork)
      return Promise.reject(new Error(t('appChartNotEditable')))
    const work = (async () => {
      if (!(await this.flushCellEditor()))
        throw new Error('Finish editing before changing a chart.')
      const operation = workbookOperationSchema.parse({
        op: 'edit_chart',
        chartPath: key,
        seriesSelection: intent,
      })
      await this.enqueue([operation], frozen, { failureInDialog: true })
    })()
    this.chartWork = work
    const clear = () => {
      if (this.chartWork === work) this.chartWork = undefined
    }
    void work.then(clear, clear)
    return work
  }
  private chartFormatKey: string | null = null
  get chartFormatTarget() {
    const visual = this.charts.find(
      (item) => item.id === this.chartFormatKey || item.chartPath === this.chartFormatKey,
    )
    if (!visual) return null
    const selected = getChartElementSelection()
    return { visual, element: selected?.visualId === visual.id ? selected.element : null }
  }
  closeChartFormat() {
    this.chartFormatKey = null
    this.changed()
  }
  editChartFormat(edit: ChartEditData) {
    if (!this.state || !this.chartFormatKey)
      return Promise.reject(new Error(t('appChartNotEditable')))
    return this.editInlineChart(this.chartFormatKey, edit, {
      stateId: this.state.stateId,
      revision: this.revision,
    })
  }
  private chartWork: Promise<unknown> | undefined
  private async readChartVector(
    id: string,
    range: string,
    frozen: { stateId: string; revision: number },
  ) {
    const current = () =>
      !this.closed && this.state?.stateId === frozen.stateId && this.revision === frozen.revision
    if (!current()) throw new Error('chart_state_changed')
    const visual = this.charts.find((item) => item.id === id || item.chartPath === id)
    if (!visual || !this.state) throw new Error(t('appChartNotEditable'))
    const resolved = resolveChartRange(range, this.state.sheets, visual.sheetId)
    const result = (await this.host.read(resolved.sheetId, resolved.bounds, frozen.stateId)) as {
      cells: ViewCell[]
    }
    if (!current()) throw new Error('chart_state_changed')
    const cells = new Map(result.cells.map((cell) => [cell.address, cell]))
    const vector = []
    for (let row = resolved.bounds.startRow; row <= resolved.bounds.endRow; row++)
      for (
        let column = resolved.bounds.startColumn;
        column <= resolved.bounds.endColumn;
        column++
      ) {
        const cell = cells.get(formatAddress(row, column))
        if (!cell) throw new Error('chart_source_incomplete')
        if (cell.formula && cell.calculationState !== 'ready' && cell.calculationState !== 'error')
          throw new Error('chart_formula_pending')
        vector.push(
          cell.formula && Object.hasOwn(cell, 'cachedValue') ? cell.cachedValue : cell.rawValue,
        )
      }
    return { vector, ref: resolved.ref }
  }
  private editInlineChart(
    id: string,
    edit: ChartEditData,
    frozen: { stateId: string; revision: number },
  ): Promise<void> {
    if (this.chartWork || this.visualWork)
      return Promise.reject(new Error(t('appChartNotEditable')))
    const work = (async () => {
      if (!(await this.flushCellEditor()))
        throw new Error('Finish editing before changing a chart.')
      if (!this.charts.some((visual) => visual.id === id || visual.chartPath === id))
        throw new Error(t('appChartNotEditable'))
      // Resolve range reads back into canonical references; the host derives
      // caches again rather than trusting browser-supplied point values.
      const supported = new Set([
        'title',
        'chartType',
        'seriesColors',
        'pointColors',
        'series',
        'legend',
        'dataLabels',
        'dataLabelPosition',
        'dataLabelFormat',
        'gridlines',
        'valueAxis',
        'gapWidthPct',
        'holeSizePct',
        'explosionPct',
        'pointExplosions',
        'grouping',
        'axisTitles',
      ])
      if (Object.keys(edit).some((key) => !supported.has(key)))
        throw new Error(t('appChartNotEditable'))
      const { series, ...properties } = edit
      const owner = this.charts.find((visual) => visual.id === id || visual.chartPath === id)!
      const seriesData = series?.flatMap((entry) => {
        const values = entry.valuesRef
          ? resolveChartRange(entry.valuesRef, this.state!.sheets, owner.sheetId)
          : undefined
        const categories = entry.categoriesRef
          ? resolveChartRange(entry.categoriesRef, this.state!.sheets, owner.sheetId)
          : undefined

        if ((entry.values && !values) || (entry.categories && !categories))
          throw new Error('chart_reference_required')
        if (values && categories && values.sheetId !== categories.sheetId)
          return [
            {
              index: entry.index,
              ...(entry.name === undefined ? {} : { name: entry.name }),
              valuesRange: values.range,
              sheetId: values.sheetId,
            },
            { index: entry.index, categoriesRange: categories.range, sheetId: categories.sheetId },
          ]
        return [
          {
            index: entry.index,
            ...(entry.name === undefined ? {} : { name: entry.name }),
            ...(values ? { valuesRange: values.range } : {}),
            ...(categories ? { categoriesRange: categories.range } : {}),
            sheetId: values?.sheetId ?? categories?.sheetId ?? owner.sheetId,
          },
        ]
      })
      const operation = workbookOperationSchema.parse({
        op: 'edit_chart',
        chartPath: id,
        ...properties,
        ...(seriesData?.length ? { seriesData } : {}),
      })
      await this.enqueue([operation], frozen, { failureInDialog: true })
    })()
    this.chartWork = work
    const clear = () => {
      if (this.chartWork === work) this.chartWork = undefined
    }
    void work.then(clear, clear)
    return work
  }
  private async readChartValues(frozen: {
    target: RibbonTarget
    stateId: string
    revision: number
  }) {
    const current = () =>
      !this.closed && this.state?.stateId === frozen.stateId && this.revision === frozen.revision
    if (!current()) throw new Error('The document changed. Open the chart command again.')
    const bounds = parseRange(frozen.target.address)
    if (rangeCellCount(bounds) > 2000)
      throw new Error('Chart selections support at most 2000 cells.')
    const result = (await this.host.read(frozen.target.sheetId, bounds, frozen.stateId)) as {
      cells: ViewCell[]
    }
    if (!current()) throw new Error('The document changed while reading chart data. Try again.')
    const cells = new Map(result.cells.map((cell) => [cell.address, cell]))
    const values: ChartGridValue[][] = []
    for (let row = bounds.startRow; row <= bounds.endRow; row++) {
      const line: ChartGridValue[] = []
      for (let column = bounds.startColumn; column <= bounds.endColumn; column++) {
        const cell = cells.get(formatAddress(row, column))
        if (!cell)
          throw new Error('Chart source data is incomplete. Try again after indexing finishes.')
        if (cell.formula && cell.calculationState !== 'ready' && cell.calculationState !== 'error')
          throw new Error('Chart source formulas are still pending. No chart was inserted.')
        line.push(
          cell.formula && Object.hasOwn(cell, 'cachedValue') ? cell.cachedValue : cell.rawValue,
        )
      }
      values.push(line)
    }
    return values
  }
  async recommendCharts(frozen: { target: RibbonTarget; stateId: string; revision: number }) {
    const recommendations = recommendCharts(await this.readChartValues(frozen))
    if (!recommendations) throw new Error(t('appChartNeedsNumericColumn'))
    return recommendations
  }
  insertChart(
    kind: RecommendedKind,
    frozen: { target: RibbonTarget; stateId: string; revision: number },
  ) {
    if (this.chartWork)
      return Promise.reject(new Error('Wait for the current chart insertion to finish.'))
    const work = this.applyChart(kind, frozen)
    this.chartWork = work
    const clear = () => {
      if (this.chartWork === work) this.chartWork = undefined
    }
    void work.then(clear, clear)
    return work
  }
  private async applyChart(
    kind: RecommendedKind,
    frozen: { target: RibbonTarget; stateId: string; revision: number },
  ) {
    if (!(await this.flushCellEditor())) throw new Error('Finish editing before inserting a chart.')
    const values = await this.readChartValues(frozen)
    const sheetName = this.state?.sheets.find((sheet) => sheet.id === frozen.target.sheetId)?.name
    if (!sheetName) throw new Error('The selected sheet no longer exists.')
    buildChartVisual({
      id: 'validate-only',
      sheetId: frozen.target.sheetId,
      sheetName,
      chartType: kind,
      dataRange: frozen.target.address,
      values,
    })
    const operation = workbookOperationSchema.parse({
      op: 'add_chart',
      sheetId: frozen.target.sheetId,
      chartType: kind,
      dataRange: frozen.target.address,
    })
    return this.enqueue([operation], frozen)
  }
  private symbolWork: Promise<{ stateId: string; revision: number }> | undefined
  insertSymbol(
    symbol: string,
    frozen: { target: RibbonTarget; stateId: string; revision: number },
  ) {
    if (this.symbolWork)
      return Promise.reject(new Error('Wait for the current symbol insertion to finish.'))
    const work = this.applySymbol(symbol, frozen)
    this.symbolWork = work
    const clear = () => {
      if (this.symbolWork === work) this.symbolWork = undefined
    }
    void work.then(clear, clear)
    return work
  }
  private async applySymbol(
    symbol: string,
    frozen: { target: RibbonTarget; stateId: string; revision: number },
  ) {
    if (!(await this.flushCellEditor())) throw new Error('Finish editing before inserting symbols.')
    if (this.closed || this.state?.stateId !== frozen.stateId || this.revision !== frozen.revision)
      throw new Error(
        'The document changed while the dialog was open. Reopen the dialog and review the cell.',
      )
    const bounds = parseRange(frozen.target.address)
    const result = (await this.host.read(
      frozen.target.sheetId,
      {
        startRow: bounds.startRow,
        endRow: bounds.startRow,
        startColumn: bounds.startColumn,
        endColumn: bounds.startColumn,
      },
      frozen.stateId,
    )) as { cells: ViewCell[]; merges?: GridRange[] }
    if (this.closed || this.state?.stateId !== frozen.stateId || this.revision !== frozen.revision)
      throw new Error('The document changed while reading the cell. Reopen the dialog.')
    const merged = result.merges?.find(
      (range) =>
        bounds.startRow >= range.startRow &&
        bounds.startRow <= range.endRow &&
        bounds.startColumn >= range.startColumn &&
        bounds.startColumn <= range.endColumn,
    )
    const address = formatAddress(
      merged?.startRow ?? bounds.startRow,
      merged?.startColumn ?? bounds.startColumn,
    )
    const cell = result.cells.find((cell) => cell.address === address)
    if (!cell) throw new Error('The selected cell could not be read. No symbols were inserted.')
    const operations = symbolIntents(frozen.target.sheetId, address, cell, symbol)
    const receipt = receiptSchema.parse(await this.enqueue(operations, frozen))
    return { stateId: receipt.stateId, revision: receipt.revision }
  }
  get selectionInPivot() {
    const selected = this.selection
    if (!selected) return false
    const features = this.selectedFeatures
    const original = features.originalMetadata as { pivotRanges?: GridRange[] } | undefined
    const ranges = [
      ...((original?.pivotRanges ?? []) as GridRange[]),
      ...((features.pivotAdditions ?? []) as { location: GridRange }[]).map(
        (pivot) => pivot.location,
      ),
    ]
    return ranges.some(
      (range) =>
        selected.range.startRow >= range.startRow &&
        selected.range.startRow <= range.endRow &&
        selected.range.startColumn >= range.startColumn &&
        selected.range.startColumn <= range.endColumn,
    )
  }
  async readHeaderFooter(target: RibbonTarget, stateId: string): Promise<HeaderFooterResult> {
    if (this.closed || this.state?.stateId !== stateId)
      throw new Error('The document changed. Open the dialog again.')
    const result = (await this.host.read(
      target.sheetId,
      {
        startRow: 0,
        endRow: 0,
        startColumn: 0,
        endColumn: 0,
      },
      stateId,
    )) as { pageSetup?: { oddHeader?: string; oddFooter?: string } | null }
    if (this.closed || this.state?.stateId !== stateId)
      throw new Error('The document changed while reading page settings. Open the dialog again.')
    const patch = this.state.sheets.find((sheet) => sheet.id === target.sheetId)?.features
      ?.pageSetupEdits as Partial<HeaderFooterResult> | undefined
    return {
      header:
        patch?.header !== undefined
          ? patch.header
          : decodeHeaderFooter(result.pageSetup?.oddHeader ?? ''),
      footer:
        patch?.footer !== undefined
          ? patch.footer
          : decodeHeaderFooter(result.pageSetup?.oddFooter ?? ''),
    }
  }
  async applyHeaderFooter(
    result: HeaderFooterResult,
    frozen: { target: RibbonTarget; stateId: string; revision: number },
  ) {
    if (!(await this.flushEditor()))
      throw new Error('Finish editing before changing page settings.')
    if (this.closed || this.state?.stateId !== frozen.stateId || this.revision !== frozen.revision)
      throw new Error('The document changed while the dialog was open. Open the dialog again.')
    const operation = workbookOperationSchema.parse({
      op: 'set_page_setup',
      sheetId: frozen.target.sheetId,
      ...result,
    })
    return this.enqueue([operation], frozen)
  }
  async readHyperlink(target: RibbonTarget, stateId: string) {
    if (this.closed || this.state?.stateId !== stateId)
      throw new Error('The document changed. Open the dialog again.')
    const bounds = parseRange(target.address)
    const address = formatAddress(bounds.startRow, bounds.startColumn)
    const result = (await this.host.read(
      target.sheetId,
      {
        startRow: bounds.startRow,
        endRow: bounds.startRow,
        startColumn: bounds.startColumn,
        endColumn: bounds.startColumn,
      },
      stateId,
    )) as { hyperlinks?: { address: string; target: string }[] }
    if (this.closed || this.state?.stateId !== stateId)
      throw new Error('The document changed. Open the dialog again.')
    return result.hyperlinks?.find((link) => link.address === address)?.target ?? null
  }
  get ribbonEcho() {
    const sheet = this.state?.sheets.find((item) => item.id === this.selection?.sheetId)
    const original = sheet?.features?.originalMetadata as
      { rowCount?: number; columnCount?: number } | undefined
    return {
      hasContent:
        !!sheet &&
        (Object.values(sheet.cells).some((cell) => cell.rawValue !== null || !!cell.formula) ||
          (!!original?.rowCount && !!original.columnCount)),
      protected: sheet ? (this.protectionEcho(sheet.id) ?? sheet.protected ?? null) : null,
    }
  }
  private protectionEchoes = new Map<string, boolean>()
  private protectionEchoState: string | undefined
  /** Effective sheet protection from the host (file element or session edit), cached per root. */
  private protectionEcho(sheetId: string) {
    const state = this.state
    if (!state) return undefined
    if (this.protectionEchoState !== state.stateId) {
      this.protectionEchoState = state.stateId
      this.protectionEchoes.clear()
    }
    if (!this.protectionEchoes.has(sheetId)) {
      this.protectionEchoes.set(
        sheetId,
        state.sheets.find((item) => item.id === sheetId)?.protected ?? false,
      )
      void this.host
        .read(sheetId, { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, state.stateId)
        .then((read) => {
          const protection = (read as { protection?: { protected: boolean } | null }).protection
          if (this.closed || this.state?.stateId !== state.stateId || !protection) return
          if (this.protectionEchoes.get(sheetId) !== protection.protected) {
            this.protectionEchoes.set(sheetId, protection.protected)
            this.changed()
          }
        })
        .catch(() => {})
    }
    return this.protectionEchoes.get(sheetId)
  }
  async ribbonCommands(
    commands: readonly string[],
    frozen?: {
      target: RibbonTarget
      stateId: string
      revision: number
    },
  ) {
    if (!(await this.flushEditor())) throw new Error('Finish editing before using the ribbon.')
    if (
      frozen &&
      (frozen.stateId !== this.selectionContext.stateId || frozen.revision !== this.revision)
    )
      throw new Error(
        'The document changed while the dialog was open. Reopen the dialog and review its settings.',
      )
    if (
      commands.length === 1 &&
      (commands[0] === 'autofit-row-height' || commands[0] === 'autofit-col-width')
    ) {
      await this.autoFit(commands[0] === 'autofit-row-height' ? 'row' : 'column')
      return
    }
    const target = frozen?.target ?? this.ribbonTarget
    if (!target) throw new Error('Select cells before using this ribbon command.')
    const needsIndent = commands.some((command) => command.startsWith('indent-step:'))
    const needsDecimals = commands.some(
      (command) => command === 'decimal-inc' || command === 'decimal-dec',
    )
    if (needsIndent || needsDecimals) {
      const bounds = parseRange(target.address)
      if (rangeCellCount(bounds) > 10_000) throw new Error('range_limit_exceeded')
      const sheet = this.runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
      if (!sheet || sheet.getSheetId() !== target.sheetId) throw new Error('view_unavailable')
      const indents: { address: string; indent: number }[] = []
      const decimals: DecimalCell[] = []
      for (let row = bounds.startRow; row <= bounds.endRow; row++)
        for (let column = bounds.startColumn; column <= bounds.endColumn; column++) {
          const style = sheet.getRange(row, column).getCellStyleData() ?? {}
          if (needsIndent)
            indents.push({ address: formatAddress(row, column), indent: indentStepsOf(style) })
          if (needsDecimals)
            decimals.push({
              row,
              col: column,
              pattern: style.n?.pattern,
              raw: sheet.getSheet().getCellRaw(row, column) ?? null,
            })
        }
      if (needsIndent) target.cellIndents = indents
      if (needsDecimals) target.decimalCells = decimals
    }
    if (commands.some((command) => command === 'clear-comments' || command === 'clear-all')) {
      const bounds = parseRange(target.address)
      const features = this.selectedFeatures
      const original = features.originalMetadata as
        { comments?: { row: number; column: number; text: string }[] } | undefined
      const notes = new Map(
        (original?.comments ?? []).map((note) => [formatAddress(note.row, note.column), note.text]),
      )
      for (const [address, text] of Object.entries(features.noteEdits ?? {})) {
        if (text === null) notes.delete(address)
        else notes.set(address, String(text))
      }
      target.noteAddresses = [...notes.keys()].filter((address) => {
        const cell = parseAddress(address)
        return (
          cell.row >= bounds.startRow &&
          cell.row <= bounds.endRow &&
          cell.column >= bounds.startColumn &&
          cell.column <= bounds.endColumn
        )
      })
    }
    if (commands.some((command) => command.startsWith('clear-'))) {
      const bounds = parseRange(target.address)
      const sheet = this.runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
      if (!sheet || sheet.getSheetId() !== target.sheetId) throw new Error('view_unavailable')
      const filteredRows: number[] = []
      for (let row = bounds.startRow; row <= bounds.endRow; row++)
        if (sheet.getSheet().getRowFiltered(row)) filteredRows.push(row)
      target.filteredRows = filteredRows
      const features = this.selectedFeatures
      const links = new Map(
        ((features.hyperlinks ?? []) as { row: number; column: number; target: string }[]).map(
          (link) => [formatAddress(link.row, link.column), link.target],
        ),
      )
      for (const [address, value] of Object.entries(features.hyperlinkEdits ?? {})) {
        if (value === null) links.delete(address)
        else links.set(address, String(value))
      }
      target.hyperlinkAddresses = [...links.keys()].filter((address) => {
        const cell = parseAddress(address)
        return (
          cell.row >= bounds.startRow &&
          cell.row <= bounds.endRow &&
          cell.column >= bounds.startColumn &&
          cell.column <= bounds.endColumn
        )
      })
    }
    const operations = commands.flatMap((command) => ribbonIntents(command, target))
    if (operations.length) await this.enqueue(operations)
  }
  private async autoFit(axis: 'row' | 'column') {
    const target = this.ribbonTarget
    const state = this.state
    const book = this.runtime.univerAPI.getActiveWorkbook()
    const sheet = book?.getActiveSheet()
    if (!target || !state || !sheet || !book) throw new Error('Select cells before AutoFit.')
    const frozen = { stateId: state.stateId, revision: this.revision }
    const bounds = parseRange(target.address)
    const canonical = state.sheets.find((item) => item.id === target.sheetId)!
    const metadata = canonical.features?.originalMetadata as
      { rowCount?: number; columnCount?: number } | undefined
    let rows = Math.max(canonical.rows, metadata?.rowCount ?? 0, 1)
    let columns = Math.max(canonical.columns, metadata?.columnCount ?? 0, 1)
    for (const address of Object.keys(canonical.cells)) {
      const cell = parseAddress(address)
      rows = Math.max(rows, cell.row + 1)
      columns = Math.max(columns, cell.column + 1)
    }
    const needed =
      axis === 'row'
        ? { ...bounds, startColumn: 0, endColumn: columns - 1 }
        : { ...bounds, startRow: 0, endRow: rows - 1 }
    const pages: { row: number; column: number }[] = []
    for (let row = Math.floor(needed.startRow / 50) * 50; row <= needed.endRow; row += 50)
      for (
        let column = Math.floor(needed.startColumn / 40) * 40;
        column <= needed.endColumn;
        column += 40
      ) {
        if (pages.length >= 100)
          throw new Error('AutoFit range limit: select fewer rows or columns.')
        pages.push({ row, column })
      }
    const assertCurrent = () => {
      if (
        this.closed ||
        this.state?.stateId !== frozen.stateId ||
        this.revision !== frozen.revision ||
        this.runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId() !==
          target.sheetId
      )
        throw new Error('The document changed during AutoFit. Try again.')
    }
    for (const page of pages) {
      assertCurrent()
      await this.loadPage(page.row, page.column)
    }
    assertCurrent()
    const skeleton = this.runtime.univer
      .__getInjector()
      .get(SheetSkeletonService)
      .getSkeleton(book.getId(), sheet.getSheetId())
    if (!skeleton) throw new Error('AutoFit measurement is unavailable.')
    installAutofitWrapBudget()
    installAutofitLinePitch()
    const ranges =
      axis === 'row'
        ? fullRowSpans([bounds], sheet.getMaxColumns())
        : fullColumnSpans([bounds], sheet.getMaxRows())
    const operations: WorkbookOperation[] = []
    if (axis === 'column') {
      for (const value of skeleton.calculateAutoWidthInRange(ranges)) {
        if (typeof value.width !== 'number' || !Number.isFinite(value.width)) continue
        operations.push(
          ...dimensionIntents(
            target.sheetId,
            axis,
            [{ startRow: 0, endRow: 0, startColumn: value.col, endColumn: value.col }],
            value.width,
          ),
        )
      }
    } else {
      // The desktop command enables automatic rows before measuring, including
      // previously forced heights. Scope that flag to this read-only measure
      // and restore every entry before submitting authoritative operations.
      const data = sheet.getSheet().getRowManager().getRowData()
      const originals = new Map<number, (typeof data)[number] | undefined>()
      try {
        for (let row = bounds.startRow; row <= bounds.endRow; row++) {
          originals.set(row, data[row])
          data[row] = { ...data[row], ia: 1 }
        }
        for (const value of skeleton.calculateAutoHeightInRange(ranges)) {
          if (typeof value.autoHeight !== 'number' || !Number.isFinite(value.autoHeight)) continue
          operations.push(
            ...dimensionIntents(
              target.sheetId,
              axis,
              [{ startRow: value.row, endRow: value.row, startColumn: 0, endColumn: 0 }],
              value.autoHeight,
            ),
          )
        }
      } finally {
        for (const [row, original] of originals) {
          if (original === undefined) delete data[row]
          else data[row] = original
        }
      }
    }
    assertCurrent()
    if (operations.length) await this.enqueue(operations, frozen)
  }
  get goToNames(): GoToNameEntry[] {
    const names: DefinedNameDefinition[] = this.state
      ? sessionDefinedNames(this.state).names.map((name) => ({
          name: name.name,
          formula: name.formula,
          ...(name.scopeSheetId
            ? {
                sheetIndex: this.state!.sheets.findIndex((sheet) => sheet.id === name.scopeSheetId),
              }
            : {}),
        }))
      : []
    const seen = new Set<string>()
    return orderDefinedNames(names)
      .filter((name) => {
        if (name.sheetIndex !== undefined && !this.state?.sheets[name.sheetIndex]) return false
        if (seen.has(name.name)) return false
        seen.add(name.name)
        return true
      })
      .map((name) => ({ name: name.name, ref: name.formula }))
  }
  async goToReference(
    input: string,
    options: { focus?: boolean; isCurrent?(): boolean } = {},
  ): Promise<string | null> {
    if (!(await this.flushEditor())) return t('appGoToNotReady')
    if (options.isCurrent?.() === false) return null
    const workbook = this.runtime.univerAPI.getActiveWorkbook()
    const active = workbook?.getActiveSheet()
    if (!workbook || !active) return t('appGoToNotReady')
    const text = input.trim()
    if (!text) return t('appGoToEmpty')
    const resolved = resolveGoToRef(text, this.goToNames)
    if (!resolved) return t('appGoToUnresolved', { ref: text })
    try {
      const parsed = deserializeRangeWithSheet(resolved)
      const sheet = parsed.sheetName
        ? workbook
            .getSheets()
            .find((item) => item.getSheetName().toLowerCase() === parsed.sheetName.toLowerCase())
        : active
      if (!sheet) return t('appGoToSheetNotFound', { ref: text })
      const raw = parsed.range
      // Expand the display before constructing FRange, which rejects cells
      // outside its currently allocated coordinate space.
      this.internal++
      try {
        if (Number.isFinite(raw.endRow) && raw.endRow >= sheet.getMaxRows())
          sheet.setRowCount(raw.endRow + 1)
        if (Number.isFinite(raw.endColumn) && raw.endColumn >= sheet.getMaxColumns())
          sheet.setColumnCount(raw.endColumn + 1)
        this.runtime.univer.__getInjector().get(IUndoRedoService).clearUndoRedo(workbook.getId())
      } finally {
        this.internal--
      }
      const bounds = {
        startRow: raw.rangeType === RANGE_TYPE.COLUMN ? 0 : raw.startRow,
        endRow: raw.rangeType === RANGE_TYPE.COLUMN ? sheet.getMaxRows() - 1 : raw.endRow,
        startColumn: raw.rangeType === RANGE_TYPE.ROW ? 0 : raw.startColumn,
        endColumn: raw.rangeType === RANGE_TYPE.ROW ? sheet.getMaxColumns() - 1 : raw.endColumn,
      }
      const range = sheet.getRange(
        bounds.startRow,
        bounds.startColumn,
        bounds.endRow - bounds.startRow + 1,
        bounds.endColumn - bounds.startColumn + 1,
      )
      workbook.setActiveRange(range)
      await this.loadPage(
        Math.floor(bounds.startRow / 50) * 50,
        Math.floor(bounds.startColumn / 40) * 40,
      )
      if (options.isCurrent?.() === false) return null
      await revealCellBelowFreeze(sheet, bounds.startRow, bounds.startColumn)
      if (options.focus !== false)
        requestAnimationFrame(() => {
          if (!this.closed && options.isCurrent?.() !== false) {
            this.ribbonFocusBoundary?.release()
            this.runtime.univer.__getInjector().get(ILayoutService).focus()
          }
        })
      return null
    } catch (error) {
      const detail = error instanceof Error ? error.message : String(error)
      if (detail.includes('Range not found')) return t('appGoToSheetNotFound', { ref: text })
      if (detail.includes('out of bounds')) return t('appGoToOutOfBounds', { ref: text })
      return t('appGoToFailed', { ref: text, detail })
    }
  }
  async ribbonSummary(fn: 'SUM' | 'AVERAGE' | 'COUNT' | 'MIN' | 'MAX') {
    if (!(await this.flushEditor())) throw new Error('Finish editing before inserting a summary.')
    await this.insertSummary(fn)
  }
  async ribbonChat(instruction = '') {
    if (!this.host.action) throw new Error('Native chat is unavailable in this host.')
    await this.host.action({
      action: 'chat',
      instruction,
      scope: instruction ? 'selection' : 'document',
    })
  }
  private runtime: ReturnType<typeof createUniver>
  private state: CandidateState | undefined
  private revision = 0
  private internal = 0
  private installing = 0
  private closed = false
  private composing = false
  private pending: Promise<unknown> = Promise.resolve()
  private pendingError = false
  private disposables: { dispose(): void }[] = []
  private loaded = new Set<string>()
  private loadedDataValidations = new Set<string>()
  /** File conditional formats per sheet, kept for projection after session edits are undone. */
  private fileCfRules = new Map<string, WorkbookRangeResult['conditionalRules']>()
  /** What each sheet's Univer CF model currently projects ('file' or the session rule JSON). */
  private cfSignatures = new Map<string, string>()
  private cfCaptures = new Map<string, ReturnType<typeof setTimeout>>()
  /**
   * Projects a sheet's conditional formats into Univer: the session rule set
   * when one exists, otherwise the file rules in the desktop's precedence order.
   */
  private projectConditionalFormats(sheetId: string) {
    const workbook = this.runtime.univerAPI.getActiveWorkbook()
    const sheet = workbook?.getSheetBySheetId(sheetId)
    const canonical = this.state?.sheets.find((item) => item.id === sheetId)
    if (!workbook || !sheet || !canonical) return
    const session = (
      canonical.features?.cfEdits as { rules: IConditionFormattingRule[] } | undefined
    )?.rules
    const file = this.fileCfRules.get(sheetId)
    if (!session && !file) return
    const signature = session ? JSON.stringify(session) : 'file'
    if (this.cfSignatures.get(sheetId) === signature) return
    this.cfSignatures.set(sheetId, signature)
    const injector = this.runtime.univer.__getInjector()
    const commands = injector.get(ICommandService)
    const unitId = workbook.getId()
    this.internal++
    try {
      for (const rule of [
        ...(injector.get(ConditionalFormattingRuleModel).getSubunitRules(unitId, sheetId) ?? []),
      ])
        commands.syncExecuteCommand(DeleteConditionalRuleMutation.id, {
          unitId,
          subUnitId: sheetId,
          cfId: rule.cfId,
        })
      // Univer's addRule unshifts: the rule added last is evaluated first.
      const rules: IConditionFormattingRule[] = session
        ? session.map((rule, index) => ({ ...rule, cfId: `session-cf-${index}` }))
        : [...dropShadowedPaintOnceRules(file!)]
            .sort((a, b) => a.priority - b.priority)
            .flatMap((rule) => {
              try {
                const built = buildConditionalRule(
                  sheet as never,
                  ((this.state?.features?.originalMetadata as { dxfStyles?: never[] } | undefined)
                    ?.dxfStyles ?? []) as never,
                  rule,
                )
                return built ? [built as IConditionFormattingRule] : []
              } catch {
                // An unsupported rule must not cost the sheet its other rules.
                return []
              }
            })
      for (const rule of [...rules].reverse())
        commands.syncExecuteCommand(AddConditionalRuleMutation.id, {
          unitId,
          subUnitId: sheetId,
          rule,
        })
    } finally {
      this.internal--
    }
  }
  /** Home → Conditional Formatting: Univer's rule manager, as on the desktop. */
  async openConditionalFormatting() {
    if (!(await this.flushEditor()))
      throw new Error('Finish editing before changing conditional formatting.')
    // value 2 = the manage-rules list; a missing params object silently no-ops.
    await this.runtime.univerAPI.executeCommand(
      'sheet.operation.open.conditional.formatting.panel',
      { value: 2 },
    )
  }
  /** Home → Conditional Formatting panel edits become one declarative session edit. */
  private captureConditionalFormats(sheetId: string) {
    const pending = this.cfCaptures.get(sheetId)
    if (pending) clearTimeout(pending)
    this.cfCaptures.set(
      sheetId,
      setTimeout(() => {
        this.cfCaptures.delete(sheetId)
        const workbook = this.runtime.univerAPI.getActiveWorkbook()
        if (!workbook || this.closed) return
        const rules = (
          this.runtime.univer
            .__getInjector()
            .get(ConditionalFormattingRuleModel)
            .getSubunitRules(workbook.getId(), sheetId) ?? []
        ).map(({ ranges, stopIfTrue, rule }) => ({
          ranges: ranges.map(({ startRow, endRow, startColumn, endColumn }) => ({
            startRow,
            endRow,
            startColumn,
            endColumn,
          })),
          stopIfTrue: Boolean(stopIfTrue),
          rule: JSON.parse(JSON.stringify(rule)) as Record<string, unknown>,
        }))
        this.cfSignatures.set(sheetId, JSON.stringify(rules))
        void this.enqueue([{ op: 'set_conditional_formats', sheetId, rules }]).catch((error) => {
          // The view returns to the authoritative rules.
          this.cfSignatures.delete(sheetId)
          this.projectConditionalFormats(sheetId)
          this.report(error)
        })
      }, 0),
    )
  }
  private pageWork: Promise<void> | undefined
  private filePageBreaks = new Map<string, { rowBreaks: number[]; colBreaks: number[] }>()
  get pageLayoutEcho(): PageLayoutEcho {
    const view = this.projectedView(),
      sheetId = this.selection?.sheetId
    const sheet = this.state?.sheets.find((item) => item.id === sheetId)
    const file = (sheetId ? this.filePageSetup.get(sheetId) : undefined) as
      Partial<PageLayoutEcho> | null | undefined
    const saved = sheet?.features?.originalMetadata as
      { printArea?: string; printTitles?: string } | undefined
    const prior = sheet?.features?.pageSetupEdits as PageSettingsPatch | undefined
    return {
      ...file,
      printArea: saved?.printArea,
      printTitles: saved?.printTitles,
      ...prior,
      ...(prior?.scale !== undefined ? { fitToWidth: 0, fitToHeight: 0 } : {}),
      // Saved margins are numeric; only a newly selected preset has a preset label.
      margins: prior?.margins,
      showGridlines: view.gridlines,
      showHeadings: view.headings,
      rightToLeft: view.rightToLeft,
      pageBreakPreview: sheetId ? this.pageBreakPreviews.has(sheetId) : false,
    }
  }
  private formulaBarHidden = false
  private crossHighlightOn = loadCrossHighlightPreference()
  private crossHighlight: CrossHighlightHandle | undefined
  private pageBreakPreviews = new Map<string, { dispose(): void }[]>()
  get formulaBarVisible() {
    return !this.formulaBarHidden
  }
  get crossHighlightVisible() {
    return this.crossHighlightOn
  }
  /** Display-only overlays (cross highlight, page-break preview) follow each projection. */
  private threadAffordances: { dispose(): void } | undefined
  private refreshViewOverlays() {
    this.threadAffordances ??= installThreadedCommentAffordances(this.runtime, {
      theme: () => (this.darkTheme() ? 'dark' : 'light'),
      onHover: (hover) => {
        this.threadHover = hover
        this.changed()
      },
    })
    if (!this.crossHighlight && this.crossHighlightOn) {
      this.crossHighlight = installCrossHighlight(this.runtime, {
        theme: () => (this.darkTheme() ? 'dark' : 'light'),
      })
      this.crossHighlight.setVisible(true)
    } else this.crossHighlight?.refresh()
    for (const sheetId of [...this.pageBreakPreviews.keys()]) void this.installPageBreaks(sheetId)
  }
  private async installPageBreaks(sheetId: string) {
    const state = this.state
    const sheet = state?.sheets.find((item) => item.id === sheetId)
    const worksheet = this.runtime.univerAPI.getActiveWorkbook()?.getSheetBySheetId(sheetId)
    if (!state || !sheet || !worksheet) return
    if (!this.filePageBreaks.has(sheetId) || !this.filePageSetup.has(sheetId)) {
      const read = (await this.host.read(
        sheetId,
        { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
        state.stateId,
      )) as { pageSetup?: unknown; pageBreaks?: { rowBreaks: number[]; colBreaks: number[] } }
      this.filePageSetup.set(sheetId, read.pageSetup ?? null)
      this.filePageBreaks.set(sheetId, read.pageBreaks ?? { rowBreaks: [], colBreaks: [] })
    }
    if (!this.pageBreakPreviews.has(sheetId) || this.state !== state) return
    const prior = (sheet.features?.pageSetupEdits ?? {}) as PageSettingsPatch
    const raw = this.filePageBreaks.get(sheetId)!
    const mapped = (axis: 'row' | 'column', values: number[]) =>
      values
        .map((value) => currentAxisIndex(state.structuralLog, sheetId, value, axis))
        .filter((value): value is number => value !== null && value > 0)
    for (const layer of this.pageBreakPreviews.get(sheetId) ?? []) layer.dispose()
    this.pageBreakPreviews.set(
      sheetId,
      installPageBreakPreview(
        this.runtime,
        worksheet,
        {
          ...((this.filePageSetup.get(sheetId) ?? {}) as object),
          ...prior,
        } as PageSetupJournalState,
        {
          rowBreaks: prior.rowBreaks ?? mapped('row', raw.rowBreaks),
          colBreaks: prior.colBreaks ?? mapped('column', raw.colBreaks),
        },
        { rows: sheet.rows, columns: sheet.columns },
        `page-break-${sheetId}`,
      ),
    )
  }
  get workbookProtected(): boolean | null {
    return this.state ? workbookStructureLocked(this.state.features) : null
  }
  /** Review: the active sheet's protection and allow-edit ranges, read from the host. */
  async protectionContext() {
    if (!(await this.flushEditor())) throw new Error(t('appClipboardBusy'))
    const target = this.ribbonTarget,
      state = this.state
    if (!target || !state) throw new Error(t('appProtectionNeedsFile'))
    const read = (await this.host.read(
      target.sheetId,
      { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
      state.stateId,
    )) as {
      protection?: { protected: boolean; hasPassword: boolean } | null
      protectedRanges?: { name: string; sqref: string; hasPassword: boolean }[]
    }
    return {
      sheetId: target.sheetId,
      activeCell: target.address.split(':')[0]!,
      frozen: { stateId: state.stateId, revision: this.revision },
      protection: read.protection ?? null,
      ranges: read.protectedRanges ?? [],
    }
  }
  /** Submits a Review protection change; returns a localized message for the dialog or null. */
  async protectionChange(
    operation: WorkbookOperation,
    frozen?: { stateId: string; revision: number },
  ): Promise<string | null> {
    try {
      await this.enqueue([operation], frozen, { failureInDialog: true })
      return null
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error)
      const code = /protection:(\w+)/.exec(message)?.[1]
      const key = (
        {
          password_required: 'appProtectedWithPassword',
          password_unavailable: 'appProtectedWithPassword',
          wrong_password: 'appProtectionWrongPassword',
          not_indexed: 'appProtectionNeedsIndexed',
          workbook_password: 'appWorkbookProtectedWithPassword',
          range_password: 'appRangesPasswordBlocked',
        } as const
      )[code as 'wrong_password']
      return key ? t(key) : message
    }
  }
  /** Review > Workbook Statistics over the canonical root (read through the host). */
  async workbookStatistics() {
    if (!(await this.flushEditor())) throw new Error(t('appClipboardBusy'))
    const state = this.state
    if (!state) throw new Error(t('appOpenXlsxFirst'))
    let cells = 0,
      formulas = 0,
      budget = 200_000,
      partial = false
    for (const sheet of state.sheets) {
      const columns = Math.max(1, Math.min(sheet.columns, 2000)),
        step = Math.max(1, Math.floor(2000 / columns))
      if (sheet.columns > columns) partial = true
      for (let row = 0; row < sheet.rows; row += step) {
        if (budget <= 0) {
          partial = true
          break
        }
        const page = (await this.host.read(
          sheet.id,
          {
            startRow: row,
            endRow: Math.min(sheet.rows - 1, row + step - 1),
            startColumn: 0,
            endColumn: columns - 1,
          },
          state.stateId,
        )) as { cells: ViewCell[] }
        if (this.state?.stateId !== state.stateId)
          throw new Error(t('appDocumentChangedDuringAction'))
        budget -= step * columns
        for (const cell of page.cells) {
          if (cell.formula) formulas++
          const value = cell.formula ? cell.cachedValue : cell.rawValue
          if (value !== null && value !== undefined && value !== '') cells++
        }
      }
    }
    return t('appWorkbookStats', {
      sheets: state.sheets.length,
      cells: cells.toLocaleString(),
      formulas: formulas.toLocaleString(),
      note: partial ? t('appLoadedRegionOnly') : '',
    })
  }
  private notesApplying = 0
  private dirtyNotes = new Map<string, { sheetId: string; row: number; column: number }>()
  private notePopupWatch: { unsubscribe(): void } | undefined
  /** File notes plus session edits, in the desktop's "Author:" display form. */
  private projectedNotes(sheet: NonNullable<typeof this.state>['sheets'][number]) {
    const original = (
      (
        sheet.features?.originalMetadata as
          | {
              comments?: {
                row: number
                column: number
                text: string
                author?: string
                thread?: unknown
              }[]
            }
          | undefined
      )?.comments ?? []
    ).filter((comment) => !comment.thread)
    const notes = new Map(
      original.map((comment) => [
        formatAddress(comment.row, comment.column),
        { author: comment.author ?? '', text: comment.text },
      ]),
    )
    for (const [address, thread] of Object.entries(
      (sheet.features?.threadEdits ?? {}) as Record<string, unknown>,
    ))
      if (thread) notes.delete(address)
    for (const [address, text] of Object.entries(
      (sheet.features?.noteEdits ?? {}) as Record<string, string | null>,
    ))
      if (text === null) notes.delete(address)
      else notes.set(address, { author: notes.get(address)?.author ?? '', text })
    return notes
  }
  /** File threads plus session thread edits, keyed by address. */
  private projectedThreads(sheet: NonNullable<typeof this.state>['sheets'][number]) {
    const threads = new Map<string, Omit<CellThread, 'row' | 'column'>>()
    for (const comment of (
      sheet.features?.originalMetadata as
        | {
            comments?: {
              row: number
              column: number
              thread?: Omit<CellThread, 'row' | 'column'>
            }[]
          }
        | undefined
    )?.comments ?? [])
      if (comment.thread) threads.set(formatAddress(comment.row, comment.column), comment.thread)
    for (const [address, thread] of Object.entries(
      (sheet.features?.threadEdits ?? {}) as Record<
        string,
        Omit<CellThread, 'row' | 'column'> | null
      >,
    ))
      if (thread) threads.set(address, thread)
      else threads.delete(address)
    for (const [address, text] of Object.entries(
      (sheet.features?.noteEdits ?? {}) as Record<string, string | null>,
    ))
      if (text !== null) threads.delete(address)
    return threads
  }
  private threadsLoading = false
  threadHover: ThreadHover | null = null
  /** The Univer runtime for shared desktop panes (comments pane). */
  get univerRuntime() {
    return this.runtime
  }
  private applyProjectedThreads() {
    const state = this.state
    if (!state) return
    this.threadsLoading = true
    try {
      threadStore.load({
        sheets: state.sheets.map((sheet) => ({
          id: sheet.id,
          comments: [...this.projectedThreads(sheet)].map(([address, thread]) => {
            const { row, column } = parseAddress(address)
            return { row, column, author: thread.author, text: thread.text, thread }
          }),
        })),
      } as unknown as Parameters<typeof threadStore.load>[0])
    } finally {
      this.threadsLoading = false
    }
    threadStore.onChange = (sheetId) => {
      if (this.threadsLoading || this.closed) return
      void this.submitThreads(sheetId).catch(this.report)
    }
  }
  /** Comments pane edits: the changed threads become canonical set_thread operations. */
  private async submitThreads(sheetId: string, extra: WorkbookOperation[] = []) {
    const sheet = this.state?.sheets.find((item) => item.id === sheetId)
    if (!sheet) return
    const projected = this.projectedThreads(sheet)
    const current = new Map(
      threadStore
        .sheetThreads(sheetId)
        .map(({ row, column, ...thread }) => [formatAddress(row, column), thread] as const),
    )
    const operations: WorkbookOperation[] = []
    for (const [address, thread] of current)
      if (JSON.stringify(projected.get(address)) !== JSON.stringify(thread))
        operations.push({
          op: 'set_thread',
          sheetId,
          address,
          thread: { ...thread, replies: thread.replies.map((reply) => ({ ...reply })) },
        })
    for (const address of projected.keys())
      if (!current.has(address))
        operations.push({ op: 'set_thread', sheetId, address, thread: null })
    if (operations.length || extra.length) await this.enqueue([...operations, ...extra])
  }
  /** Review > Comments (threaded) and the shared note navigation. */
  async commentCommand(command: string): Promise<string | null> {
    const target = this.ribbonTarget
    if (!target) throw new Error(t('appSelectCellFirst'))
    const sheet = this.state?.sheets.find((item) => item.id === target.sheetId)
    if (!sheet) throw new Error(t('appSelectCellFirst'))
    const { row, column } = parseAddress(target.address.split(':')[0]!)
    const thread = threadStore.get(sheet.id, row, column)
    const note = this.projectedNotes(sheet).has(formatAddress(row, column))
    switch (command) {
      case 'comment-new':
        if (thread) threadPane.openAt(sheet.id, row, column, false)
        else if (note) return t('appThreadCellHasNote')
        else threadPane.openAt(sheet.id, row, column, true)
        return null
      case 'comment-reply':
        if (!thread) return t('appThreadNoneAtSelection')
        threadPane.openAt(sheet.id, row, column, false)
        return null
      case 'comment-resolve':
        if (!thread) return t('appThreadNoneAtSelection')
        threadStore.setDone(sheet.id, row, column, !thread.done)
        return null
      case 'comment-delete':
        if (thread) {
          threadStore.remove(sheet.id, row, column)
          return null
        }
        return this.noteCommand('note-delete')
      case 'comments-pane-toggle':
        threadPane.toggle()
        return null
      case 'notes-convert': {
        const threads = threadStore.sheetThreads(sheet.id)
        if (!threads.length) return t('appNoCommentsOnSheet')
        this.threadsLoading = true
        try {
          for (const entry of threads) threadStore.convertToNote(sheet.id, entry.row, entry.column)
        } finally {
          this.threadsLoading = false
        }
        await this.submitThreads(
          sheet.id,
          threads.map((entry) => ({
            op: 'set_note' as const,
            sheetId: sheet.id,
            address: formatAddress(entry.row, entry.column),
            text: [entry.text, ...entry.replies.map((reply) => reply.text)].join('\n\n'),
          })),
        )
        return t('appThreadConverted', { count: threads.length })
      }
      case 'comment-prev':
      case 'comment-next': {
        const worksheet = this.runtime.univerAPI.getActiveWorkbook()?.getSheetBySheetId(sheet.id)
        if (!worksheet) return null
        const stops = [
          ...threadStore.sheetThreads(sheet.id).map(({ row, column }) => ({ row, column })),
          ...[...this.projectedNotes(sheet).keys()].map((address) => parseAddress(address)),
        ].sort((a, b) => a.row - b.row || a.column - b.column)
        if (!stops.length) return t('appNoNotesOnSheet')
        const after = stops.find(
          (stop) => stop.row > row || (stop.row === row && stop.column > column),
        )
        const before = [...stops]
          .reverse()
          .find((stop) => stop.row < row || (stop.row === row && stop.column < column))
        const next = command === 'comment-next' ? (after ?? stops[0]!) : (before ?? stops.at(-1)!)
        worksheet.getRange(next.row, next.column, 1, 1).activate()
        await revealCellBelowFreeze(worksheet, next.row, next.column)
        if (threadStore.get(sheet.id, next.row, next.column))
          threadPane.openAt(sheet.id, next.row, next.column, false)
        else await this.runtime.univerAPI.executeCommand('sheet.operation.add-note-popup')
        return null
      }
    }
    throw new Error(`Unknown comment command: ${command}`)
  }
  private applyProjectedNotes() {
    const state = this.state,
      workbook = this.runtime.univerAPI.getActiveWorkbook()
    if (!state || !workbook) return
    this.notesApplying++
    try {
      for (const sheet of state.sheets) {
        const worksheet = workbook.getSheetBySheetId(sheet.id)
        if (!worksheet) continue
        const desired = this.projectedNotes(sheet)
        for (const note of worksheet.getNotes() as { row: number; col: number }[])
          if (!desired.has(formatAddress(note.row, note.col)))
            worksheet.getRange(note.row, note.col).deleteNote()
        for (const [address, note] of desired) {
          const { row, column } = parseAddress(address)
          worksheet.getRange(row, column).createOrUpdateNote({
            id: `note-${sheet.id}-${row}-${column}`,
            row,
            col: column,
            width: 220,
            height: 90,
            note: note.author ? `${note.author}:\n${note.text}` : note.text,
          })
        }
      }
    } finally {
      this.notesApplying--
    }
    if (!this.notePopupWatch)
      this.notePopupWatch = this.runtime.univer
        .__getInjector()
        .get(SheetsNotePopupService)
        .activePopup$.subscribe((popup) => {
          if (!popup || popup.temp) void this.flushNotes().catch(this.report)
        })
  }
  /** Commits note popup edits as canonical set_note operations once the popup closes. */
  private async flushNotes() {
    if (!this.dirtyNotes.size || !this.state) return
    const dirty = [...this.dirtyNotes.values()]
    this.dirtyNotes.clear()
    const workbook = this.runtime.univerAPI.getActiveWorkbook()
    const model = this.runtime.univer.__getInjector().get(SheetsNoteModel)
    const operations: WorkbookOperation[] = []
    for (const { sheetId, row, column } of dirty) {
      const sheet = this.state.sheets.find((item) => item.id === sheetId)
      if (!sheet || !workbook) continue
      const address = formatAddress(row, column)
      const current = model.getNote(workbook.getId(), sheetId, { row, col: column }) as
        { note?: string } | undefined
      const previous = this.projectedNotes(sheet).get(address)
      let text = current?.note ?? ''
      if (previous?.author && text.startsWith(`${previous.author}:\n`))
        text = text.slice(previous.author.length + 2)
      const next = text.trim() === '' ? null : text
      if ((previous?.text ?? null) === next) continue
      operations.push({ op: 'set_note', sheetId, address, text: next })
    }
    if (operations.length) await this.enqueue(operations)
  }
  /** Review > Notes commands on the selection. */
  async noteCommand(command: string) {
    const target = this.ribbonTarget
    if (!target) throw new Error(t('appSelectCellFirst'))
    const sheet = this.state?.sheets.find((item) => item.id === target.sheetId)
    const workbook = this.runtime.univerAPI.getActiveWorkbook()
    const worksheet = workbook?.getSheetBySheetId(target.sheetId)
    if (!sheet || !workbook || !worksheet) throw new Error(t('appSelectCellFirst'))
    const cell = parseAddress(target.address.split(':')[0]!)
    const notes = worksheet.getNotes() as { row: number; col: number; show?: boolean }[]
    if (command === 'note-open') {
      await this.runtime.univerAPI.executeCommand('sheet.operation.add-note-popup')
      return null
    }
    if (command === 'note-show-toggle') {
      await this.runtime.univerAPI.executeCommand('sheet.command.toggle-note-popup')
      return null
    }
    if (command === 'note-show-all') {
      const hidden = notes.filter((note) => !note.show)
      const targets = hidden.length ? hidden : notes
      if (!targets.length) return t('appNoNotesOnSheet')
      for (const note of targets)
        await this.runtime.univerAPI.executeCommand('sheet.mutation.toggle-note-popup', {
          unitId: workbook.getId(),
          sheetId: target.sheetId,
          row: note.row,
          col: note.col,
        })
      return null
    }
    if (command === 'note-delete' || command === 'comment-delete') {
      const address = formatAddress(cell.row, cell.column)
      if (!this.projectedNotes(sheet).has(address)) return t('appNoNotesOnSheet')
      await this.enqueue([{ op: 'set_note', sheetId: target.sheetId, address, text: null }])
      return null
    }
    if (command === 'comment-prev' || command === 'comment-next') {
      const stops = notes
        .map((note) => ({ row: note.row, column: note.col }))
        .sort((a, b) => a.row - b.row || a.column - b.column)
      if (!stops.length) return t('appNoNotesOnSheet')
      const after = stops.find(
        (stop) => stop.row > cell.row || (stop.row === cell.row && stop.column > cell.column),
      )
      const before = [...stops]
        .reverse()
        .find((stop) => stop.row < cell.row || (stop.row === cell.row && stop.column < cell.column))
      const next =
        command === 'comment-next' ? (after ?? stops[0]!) : (before ?? stops[stops.length - 1]!)
      worksheet.getRange(next.row, next.column, 1, 1).activate()
      await this.runtime.univerAPI.executeCommand('sheet.operation.add-note-popup')
      return null
    }
    throw new Error(`Unknown note command: ${command}`)
  }
  /** Insert > PivotChart: a chart beside the file PivotTable, its series built like the desktop's. */
  async insertPivotChart(kind: string) {
    if (!(await this.flushEditor())) throw new Error(t('appClipboardBusy'))
    const state = this.state
    const hit = this.pivotAtSelection()
    if (!state || !hit) throw new Error(t('appCursorNotInPivot'))
    const definition = hit.file?.definition
    if (!definition || hit.relayout) throw new Error(t('appPivotDefNotLoadedSave'))
    const frozen = { stateId: state.stateId, revision: this.revision }
    const bounds = parseRange(definition.outputRef)
    const width = bounds.endColumn - bounds.startColumn + 1
    const cells = new Map<string, ViewCell>()
    const step = Math.max(1, Math.floor(2000 / width))
    for (let row = bounds.startRow; row <= bounds.endRow; row += step) {
      const page = (await this.host.read(
        hit.sheet.id,
        { ...bounds, startRow: row, endRow: Math.min(bounds.endRow, row + step - 1) },
        frozen.stateId,
      )) as { cells: ViewCell[] }
      for (const cell of page.cells) cells.set(cell.address, cell)
    }
    const values = Array.from({ length: bounds.endRow - bounds.startRow + 1 }, (_, r) =>
      Array.from({ length: width }, (_, c) => {
        const cell = cells.get(formatAddress(bounds.startRow + r, bounds.startColumn + c))
        return (cell?.formula ? cell.cachedValue : cell?.rawValue) as
          string | number | boolean | null | undefined
      }),
    )
    const chart = buildPivotChartData(definition, hit.sheet.name, bounds, values)
    if (!chart.series.length) throw new Error(t('appPivotNoChartData'))
    const chartType = (['column', 'bar', 'line', 'pie', 'doughnut', 'radar'] as const).find(
      (type) => type === kind,
    )
    if (!chartType) throw new Error(`Unknown PivotChart type: ${kind}`)
    // The session numbers new charts per sheet; the series edit targets that id.
    const existing = (
      (hit.sheet.features?.visuals ?? hit.sheet.features?.charts ?? []) as unknown[]
    ).length
    const chartPath = `demo-chart-${hit.sheet.id}-${existing + 1}`
    try {
      await this.enqueue(
        [
          {
            op: 'add_chart',
            sheetId: hit.sheet.id,
            chartType,
            dataRange: definition.outputRef,
            title: chart.title,
            anchorCell: formatAddress(bounds.startRow, bounds.endColumn + 2),
          },
          {
            op: 'edit_chart',
            chartPath,
            seriesSelection: {
              mode: 'rows',
              rows: chart.series.slice(0, 24).map((series) => ({
                name: series.name.slice(0, 255),
                ...(series.valuesRef ? { valuesRef: series.valuesRef } : {}),
                ...(series.categoriesRef ? { categoriesRef: series.categoriesRef } : {}),
              })),
            },
          },
        ],
        frozen,
        { failureInDialog: true },
      )
    } catch (error) {
      throw new Error(error instanceof Error ? error.message : String(error), { cause: error })
    }
    return t(chart.truncated ? 'appPivotChartInsertedTruncated' : 'appPivotChartInserted')
  }
  private sparklineLayers: { dispose(): void }[] = []
  /** File and session sparklines of the active sheet (shared desktop renderer). */
  private installProjectedSparklines() {
    for (const layer of this.sparklineLayers) layer.dispose()
    this.sparklineLayers = []
    const state = this.state
    const sheetId = this.runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId()
    const sheet = state?.sheets.find((item) => item.id === sheetId)
    if (!sheet) return
    const file = ((sheet.features?.originalMetadata as { sparklines?: unknown[] } | undefined)
      ?.sparklines ?? []) as Parameters<typeof installSparklines>[1]
    const session = (
      (sheet.features?.sparklineAdditions ?? []) as unknown as SessionSparklineGroup[]
    ).map((group) => ({
      type: group.type,
      ...(group.color === undefined ? {} : { color: group.color }),
      cells: sparklineCells(sheet.name, group),
    }))
    const groups = [...file, ...session]
    if (groups.length)
      this.sparklineLayers = installSparklines(
        this.runtime as unknown as Parameters<typeof installSparklines>[0],
        groups,
        sheet.id,
      )
  }
  /** Insert > Sparklines: one sparkline per selected row, right of the selection. */
  async insertSparklines(type: 'line' | 'column' | 'stacked') {
    if (!(await this.flushEditor())) throw new Error(t('appClipboardBusy'))
    const target = this.ribbonTarget
    if (!target || !this.state) throw new Error(t('appSelectCellFirst'))
    const bounds = parseRange(target.address)
    if (bounds.endColumn === bounds.startColumn) throw new Error(t('appSparklineNeedsCols'))
    if (bounds.endColumn + 1 >= MAX_GRID_COLUMNS) throw new Error(t('appSparklineNoSpace'))
    const rows = Math.min(bounds.endRow - bounds.startRow + 1, 200)
    await this.enqueue([
      {
        op: 'add_sparkline',
        sheetId: target.sheetId,
        type,
        dataRange: `${formatAddress(bounds.startRow, bounds.startColumn)}:${formatAddress(bounds.startRow + rows - 1, bounds.endColumn)}`,
      },
    ])
    return t('appSparklinesInserted', { count: rows })
  }
  /** File and session tables render through Univer's table plugin (filters, structured refs, theme). */
  private async registerProjectedTables() {
    const state = this.state,
      workbook = this.runtime.univerAPI.getActiveWorkbook()
    if (!state || !workbook) return
    const themeColors = (state.features?.originalMetadata as { themeColors?: string[] } | undefined)
      ?.themeColors
    this.internal++
    try {
      for (const sheet of state.sheets) {
        const worksheet = workbook.getSheetBySheetId(sheet.id) as unknown as {
          getSubTableInfos?(): { id: string }[]
          removeTable(id: string): Promise<boolean>
          addTable(
            name: string,
            range: GridRange,
            id: string,
            options?: unknown,
          ): Promise<boolean> | boolean
          addTableTheme(id: string, theme: unknown): Promise<boolean>
        } | null
        if (!worksheet) continue
        for (const existing of worksheet.getSubTableInfos?.() ?? [])
          await Promise.resolve(worksheet.removeTable(existing.id)).catch(() => false)
        for (const [index, table] of sessionTablesOf(sheet).entries()) {
          // Univer requires a header row (as on the desktop).
          if (!table.options.headerRow) continue
          const id = `projected-table-${sheet.id}-${index}`
          try {
            await worksheet.addTable(
              table.name,
              { ...table.area, endRow: table.area.endRow - (table.options.totalsRow ? 1 : 0) },
              id,
              {
                columns: table.columnNames.map((name, column) => ({
                  id: `${id}-col-${column}`,
                  displayName: name,
                })),
              },
            )
            const styled =
              table.kind === 'session' || table.style !== table.file?.styleName
                ? builtinTablePalette(table.style ?? DEFAULT_TABLE_STYLE, themeColors)
                : Object.fromEntries(
                    TABLE_PALETTE_KEYS.flatMap((key) =>
                      table.file?.[key] === undefined ? [] : [[key, table.file[key]]],
                    ),
                  )
            await worksheet.addTableTheme(
              id,
              tableThemeJson(`${id}-${state.stateId}`, styled, {
                bandedRows: table.options.bandedRows,
                bandedColumns: table.options.bandedColumns,
                firstColumn: table.options.firstColumn,
                lastColumn: table.options.lastColumn,
              }),
            )
          } catch {
            // Rendering is best effort; the canonical table is unaffected.
          }
        }
      }
    } finally {
      this.internal--
    }
  }
  /** Table Design echo for the table under the selection. */
  get selectedTable(): SelectedTableRibbon | null {
    const selected = this.selection,
      sheet = this.state?.sheets.find((item) => item.id === selected?.sheetId)
    if (!selected || !sheet) return null
    const { startRow: row, startColumn: column } = selected.range
    const table = sessionTablesOf(sheet).find(
      (item) =>
        row >= item.area.startRow &&
        row <= item.area.endRow &&
        column >= item.area.startColumn &&
        column <= item.area.endColumn,
    )
    if (!table) return null
    return {
      key: `${sheet.id}:${table.kind}:${table.originalName}`,
      name: table.name,
      rangeA1: `${formatAddress(table.area.startRow, table.area.startColumn)}:${formatAddress(table.area.endRow, table.area.endColumn)}`,
      style: table.style ?? null,
      isFile: table.kind === 'file',
      ...table.options,
    }
  }
  /** Insert > Table / Format as Table: unique header names are written with the table. */
  async createTable(range: string, hasHeaders: boolean, style = DEFAULT_TABLE_STYLE) {
    if (!(await this.flushEditor())) throw new Error(t('appClipboardBusy'))
    const target = this.ribbonTarget,
      state = this.state
    if (!target || !state) throw new Error(t('appTablesNeedFile'))
    let area = parseRange(range)
    const frozen = { stateId: state.stateId, revision: this.revision }
    const width = area.endColumn - area.startColumn + 1
    if (width > 1000) throw new Error(t('appTableTooWide'))
    const operations: WorkbookOperation[] = []
    if (!hasHeaders) {
      // Excel adds a header row above the data when that row is free.
      if (area.startRow === 0) throw new Error(t('appTableNoRoomForHeader'))
      const above = (await this.host.read(
        target.sheetId,
        { ...area, startRow: area.startRow - 1, endRow: area.startRow - 1 },
        frozen.stateId,
      )) as { cells: ViewCell[] }
      if (above.cells.some((cell) => cell.rawValue !== null || cell.formula))
        throw new Error(t('appTableNoRoomForHeader'))
      area = { ...area, startRow: area.startRow - 1 }
    }
    const header = (await this.host.read(
      target.sheetId,
      { ...area, endRow: area.startRow },
      frozen.stateId,
    )) as { cells: ViewCell[] }
    const byAddress = new Map(header.cells.map((cell) => [cell.address, cell]))
    const used = new Set<string>()
    for (let index = 0; index < width; index++) {
      const address = formatAddress(area.startRow, area.startColumn + index)
      const cell = byAddress.get(address)
      const raw = hasHeaders
        ? String((cell?.formula ? cell.cachedValue : cell?.rawValue) ?? '')
            .trim()
            .slice(0, 255)
        : ''
      const base = raw.length === 0 ? `Column${index + 1}` : raw
      let candidate = base
      for (let suffix = 2; used.has(candidate.toLowerCase()); suffix++)
        candidate = `${base}${suffix}`
      used.add(candidate.toLowerCase())
      if (candidate !== raw || cell?.formula)
        operations.push({ op: 'set_cell', sheetId: target.sheetId, address, value: candidate })
    }
    operations.push({
      op: 'add_table',
      sheetId: target.sheetId,
      range: `${formatAddress(area.startRow, area.startColumn)}:${formatAddress(area.endRow, area.endColumn)}`,
      style,
      bandedRows: true,
    })
    try {
      await this.enqueue(operations, frozen, { failureInDialog: true })
    } catch (error) {
      throw new Error(tableErrorMessage(error), { cause: error })
    }
    return t('appTableCreated')
  }
  /** Header/totals row toggles: the desktop's cell writes plus the table area change. */
  private async tableBandOperations(
    table: SelectedTableRibbon,
    key: 'headerRow' | 'totalsRow',
    on: boolean,
    stateId: string,
  ): Promise<WorkbookOperation[]> {
    const sheetId = this.ribbonTarget!.sheetId
    const area = parseRange(table.rangeA1)
    const width = area.endColumn - area.startColumn + 1
    if (table[key] === on) return []
    const rangeOf = (bounds: GridRange) =>
      `${formatAddress(bounds.startRow, bounds.startColumn)}:${formatAddress(bounds.endRow, bounds.endColumn)}`
    const free = async (row: number) => {
      const read = (await this.host.read(
        sheetId,
        { ...area, startRow: row, endRow: row },
        stateId,
      )) as { cells: ViewCell[] }
      return !read.cells.some((cell) => cell.rawValue !== null || cell.formula)
    }
    const base = { op: 'edit_table' as const, sheetId, table: table.name }
    if (key === 'totalsRow') {
      if (on) {
        const row = area.endRow + 1
        if (!(await free(row))) throw new Error(t('appTableTotalsRowBlocked'))
        const dataStart = area.startRow + (table.headerRow ? 1 : 0)
        const operations: WorkbookOperation[] = []
        if (width > 1)
          operations.push({
            op: 'set_cell',
            sheetId,
            address: formatAddress(row, area.startColumn),
            value: t('appTableTotalLabel'),
          })
        if (area.endRow >= dataStart) {
          const column = (await this.host.read(
            sheetId,
            {
              startRow: dataStart,
              endRow: area.endRow,
              startColumn: area.endColumn,
              endColumn: area.endColumn,
            },
            stateId,
          )) as { cells: ViewCell[] }
          const values = column.cells.map((cell) =>
            cell.formula ? cell.cachedValue : cell.rawValue,
          )
          operations.push({
            op: 'set_formula',
            sheetId,
            address: formatAddress(row, area.endColumn),
            formula: `=SUBTOTAL(${totalsFunctionFor(values)},${rangeOf({ ...area, startRow: dataStart, startColumn: area.endColumn })})`,
          })
        }
        operations.push(
          {
            op: 'format_range',
            sheetId,
            range: rangeOf({ ...area, startRow: row, endRow: row }),
            format: { bold: true },
          },
          { ...base, area: rangeOf({ ...area, endRow: row }), totalsRow: true },
        )
        return operations
      }
      const totals = { ...area, startRow: area.endRow }
      return [
        { op: 'clear_range', sheetId, range: rangeOf(totals) },
        { op: 'format_range', sheetId, range: rangeOf(totals), format: { bold: false } },
        { ...base, area: rangeOf({ ...area, endRow: area.endRow - 1 }), totalsRow: false },
      ]
    }
    if (on) {
      const row = area.startRow - 1
      if (row < 0 || !(await free(row))) throw new Error(t('appTableHeaderRowBlocked'))
      const sheet = this.state!.sheets.find((item) => item.id === sheetId)!
      const current = sessionTablesOf(sheet).find((item) => item.name === table.name)
      const names = uniqueColumnNames(
        Array.from({ length: width }, (_, index) => current?.columnNames[index] ?? ''),
      )
      return [
        {
          op: 'set_range',
          sheetId,
          start: formatAddress(row, area.startColumn),
          values: [names],
        },
        { ...base, area: rangeOf({ ...area, startRow: row }), headerRow: true },
      ]
    }
    return [
      { op: 'clear_range', sheetId, range: rangeOf({ ...area, endRow: area.startRow }) },
      { ...base, area: rangeOf({ ...area, startRow: area.startRow + 1 }), headerRow: false },
    ]
  }
  /** Insert > Table / Home > Format as Table: a single cell expands to its current region. */
  async formatAsTable(style = DEFAULT_TABLE_STYLE) {
    const target = this.ribbonTarget
    const worksheet = this.runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
    if (!target || !worksheet) throw new Error(t('appTableSelectRange'))
    let bounds = parseRange(target.address)
    if (bounds.startRow === bounds.endRow && bounds.startColumn === bounds.endColumn) {
      const inferred = inferContinuousRegion(
        worksheet as unknown as Parameters<typeof inferContinuousRegion>[0],
        bounds.startRow,
        bounds.startColumn,
      )
      if (inferred) bounds = inferred
    }
    return this.createTable(
      `${formatAddress(bounds.startRow, bounds.startColumn)}:${formatAddress(bounds.endRow, bounds.endColumn)}`,
      true,
      style,
    )
  }
  /** Table Design commands on the selected table. */
  async tableCommand(command: string): Promise<string | null> {
    if (!(await this.flushEditor())) throw new Error(t('appClipboardBusy'))
    const table = this.selectedTable,
      target = this.ribbonTarget,
      state = this.state
    if (!table || !target || !state) return null
    const frozen = { stateId: state.stateId, revision: this.revision }
    const base = { op: 'edit_table' as const, sheetId: target.sheetId, table: table.name }
    let operations: WorkbookOperation[]
    let message: string | null = null
    if (command.startsWith('table-style:'))
      operations = [{ ...base, style: command.slice('table-style:'.length) }]
    else if (command.startsWith('table-rename:')) {
      const name = command.slice('table-rename:'.length).trim()
      if (name === table.name) return null
      operations = [{ ...base, name }]
    } else if (command === 'table-convert-range') {
      operations = [{ ...base, remove: true }]
      message = t('appTableConverted')
    } else if (command.startsWith('table-resize:')) {
      operations = [{ ...base, area: command.slice('table-resize:'.length).toUpperCase() }]
      message = t('appTableResized')
    } else if (command.startsWith('table-option:')) {
      const [, key = '', value = ''] = command.split(':')
      if (key === 'headerRow' || key === 'totalsRow') {
        operations = await this.tableBandOperations(table, key, value === '1', frozen.stateId)
        if (!operations.length) return null
        try {
          await this.enqueue(operations, frozen, { failureInDialog: true })
        } catch (error) {
          throw new Error(tableErrorMessage(error), { cause: error })
        }
        return null
      }
      if (
        !['bandedRows', 'bandedColumns', 'firstColumn', 'lastColumn', 'filterButton'].includes(key)
      )
        throw new Error(`Unknown table option: ${key}`)
      operations = [{ ...base, [key]: value === '1' }]
    } else throw new Error(`Unknown Table Design command: ${command}`)
    try {
      await this.enqueue(operations, frozen, { failureInDialog: true })
    } catch (error) {
      throw new Error(tableErrorMessage(error), { cause: error })
    }
    return message
  }
  /** Insert > Icons / Equation: the rendered PNG becomes a canonical picture at the selection. */
  async insertRenderedPicture(dataUrl: string, fileName: string) {
    if (!(await this.flushEditor())) throw new Error('Finish editing before inserting a picture.')
    const target = this.ribbonTarget
    if (!target || !this.state || this.closed) throw new Error(t('appSelectCellFirst'))
    const frozen = { stateId: this.state.stateId, revision: this.revision }
    // The webview CSP does not allow fetching data: URLs; decode the PNG directly.
    const base64 = /^data:image\/png;base64,(.+)$/.exec(dataUrl)?.[1]
    if (!base64) throw new Error(t('appPictureBadType'))
    const bytes = Uint8Array.from(atob(base64), (char) => char.charCodeAt(0))
    await this.insertPictureFile(new File([bytes], fileName, { type: 'image/png' }), target, frozen)
  }
  /** Insert > Checkbox: a checkbox validation rule on the selection, as on the desktop. */
  async insertCheckbox() {
    if (!(await this.flushEditor())) throw new Error(t('appClipboardBusy'))
    const target = this.ribbonTarget
    if (!target) throw new Error(t('appSelectCellFirst'))
    await this.enqueue([
      {
        op: 'set_data_validation',
        sheetId: target.sheetId,
        range: target.address,
        validation: { kind: 'checkbox' },
      },
    ])
  }
  /** View tab commands (workbook views, show, zoom, window). */
  async viewCommand(command: string): Promise<string | null> {
    const target = this.ribbonTarget,
      state = this.state
    if (!target || !state) throw new Error(t('appSelectCellFirst'))
    if (command === 'toggle-formula-bar') {
      this.formulaBarHidden = !this.formulaBarHidden
      this.container.classList.toggle('formula-bar-hidden', this.formulaBarHidden)
      this.changed()
      return t(this.formulaBarHidden ? 'appFormulaBarHidden' : 'appFormulaBarShown')
    }
    if (command === 'toggle-cross-highlight') {
      this.crossHighlightOn = !this.crossHighlightOn
      storeCrossHighlightPreference(this.crossHighlightOn)
      if (this.crossHighlightOn) this.refreshViewOverlays()
      this.crossHighlight?.setVisible(this.crossHighlightOn)
      this.changed()
      return t(this.crossHighlightOn ? 'appCrossHighlightOn' : 'appCrossHighlightOff')
    }
    if (command === 'toggle-page-break-preview') {
      const layers = this.pageBreakPreviews.get(target.sheetId)
      if (layers) {
        for (const layer of layers) layer.dispose()
        this.pageBreakPreviews.delete(target.sheetId)
        this.changed()
        return t('appPageBreakPreviewOff')
      }
      this.pageBreakPreviews.set(target.sheetId, [])
      await this.installPageBreaks(target.sheetId)
      this.changed()
      return t('appPageBreakPreviewOn')
    }
    if (!(await this.flushEditor())) throw new Error(t('appClipboardBusy'))
    const frozen = { stateId: state.stateId, revision: this.revision }
    const view = this.projectedView()
    const bounds = parseRange(target.address)
    let operation: WorkbookOperation
    let message: string | null
    if (command === 'toggle-right-to-left') {
      operation = { op: 'set_page_setup', sheetId: target.sheetId, rightToLeft: !view.rightToLeft }
      message = t(view.rightToLeft ? 'appRightToLeftOff' : 'appRightToLeftOn')
    } else if (command.startsWith('zoom:') || command === 'zoom-reset') {
      const percent = command === 'zoom-reset' ? 100 : Number(command.slice('zoom:'.length))
      if (!Number.isInteger(percent) || percent < 10 || percent > 400)
        throw new Error('Unsupported zoom level.')
      operation = { op: 'set_page_setup', sheetId: target.sheetId, zoomScale: percent }
      message = t('appZoom', { percent })
    } else if (command === 'zoom-to-selection') {
      const workbook = this.runtime.univerAPI.getActiveWorkbook()
      const render = workbook
        ? this.runtime.univer
            .__getInjector()
            .get(IRenderManagerService)
            .getRenderById(workbook.getId())
        : null
      const skeleton = render?.with(SheetSkeletonManagerService).getCurrentSkeleton()
      if (!render || !skeleton) throw new Error(t('appSelectCellFirst'))
      const rows = skeleton.rowHeightAccumulation,
        columns = skeleton.columnWidthAccumulation
      const top = bounds.startRow > 0 ? (rows[bounds.startRow - 1] ?? 0) : 0
      const bottom = rows[Math.min(bounds.endRow, rows.length - 1)] ?? 0
      const left = bounds.startColumn > 0 ? (columns[bounds.startColumn - 1] ?? 0) : 0
      const right = columns[Math.min(bounds.endColumn, columns.length - 1)] ?? 0
      const width = render.engine.width - skeleton.rowHeaderWidthAndMarginLeft,
        height = render.engine.height - skeleton.columnHeaderHeightAndMarginTop
      if (right <= left || bottom <= top || width <= 0 || height <= 0) return null
      const ratio = clampZoomRatio(Math.min(width / (right - left), height / (bottom - top)))
      const percent = Math.max(10, Math.min(400, Math.round(ratio * 100)))
      operation = { op: 'set_page_setup', sheetId: target.sheetId, zoomScale: percent }
      message = t('appZoom', { percent })
    } else if (
      command === 'freeze-here' ||
      command === 'freeze-top-row' ||
      command === 'freeze-first-col' ||
      command === 'unfreeze'
    ) {
      const [rows, columns] =
        command === 'freeze-here'
          ? [bounds.startRow, bounds.startColumn]
          : command === 'freeze-top-row'
            ? [1, 0]
            : command === 'freeze-first-col'
              ? [0, 1]
              : [0, 0]
      operation = { op: 'set_freeze', sheetId: target.sheetId, rows, columns }
      message =
        command === 'freeze-here' ? t('appFrozenAtSelection') : command === 'unfreeze' ? null : null
    } else throw new Error(`Unknown View command: ${command}`)
    await this.enqueue([operation], frozen, { failureInDialog: true })
    if (command === 'zoom-to-selection')
      void this.runtime.univerAPI.executeCommand('sheet.command.scroll-to-cell', {
        range: {
          startRow: bounds.startRow,
          startColumn: bounds.startColumn,
          endRow: bounds.startRow,
          endColumn: bounds.startColumn,
        },
      })
    return message
  }
  async pageLayoutCommand(command: string) {
    if (!(await this.flushEditor())) throw new Error(t('appClipboardBusy'))
    const target = this.ribbonTarget
    if (!target || !this.state) throw new Error(t('appActiveSheetUnavailable'))
    const view = this.projectedView()
    const frozen = { stateId: this.state.stateId, revision: this.revision }
    const bounds = parseRange(target.address)
    const dimensions = target.dimensions ?? { rows: 1048576, columns: 16384 }
    const work = (async () => {
      const read = (await this.host.read(
        target.sheetId,
        {
          startRow: bounds.startRow,
          endRow: bounds.startRow,
          startColumn: bounds.startColumn,
          endColumn: bounds.startColumn,
        },
        frozen.stateId,
      )) as { pageSetup?: unknown; pageBreaks?: { rowBreaks: number[]; colBreaks: number[] } }
      if (
        this.closed ||
        this.state?.stateId !== frozen.stateId ||
        this.revision !== frozen.revision
      )
        throw new Error(t('appDocumentChangedDuringAction'))
      if (Object.hasOwn(read, 'pageSetup')) this.filePageSetup.set(target.sheetId, read.pageSetup)
      if (read.pageBreaks) this.filePageBreaks.set(target.sheetId, read.pageBreaks)
      const prior = (this.state.sheets.find((sheet) => sheet.id === target.sheetId)?.features
        ?.pageSetupEdits ?? {}) as PageSettingsPatch
      const raw = this.filePageBreaks.get(target.sheetId) ?? { rowBreaks: [], colBreaks: [] }
      const mapped = (axis: 'row' | 'column', values: number[]) =>
        values
          .map((value) => currentAxisIndex(this.state!.structuralLog, target.sheetId, value, axis))
          .filter((value): value is number => value !== null && value > 0)
      const breaks = {
        rowBreaks: prior.rowBreaks ?? mapped('row', raw.rowBreaks),
        colBreaks: prior.colBreaks ?? mapped('column', raw.colBreaks),
      }
      let patch: PageSettingsPatch | null
      if (command === 'toggle-gridlines') patch = { showGridlines: !view.gridlines }
      else if (command === 'toggle-headings') patch = { showHeadings: !view.headings }
      else
        patch = pageSettingsCommand(command.replace(/^page-layout:/, ''), {
          sheetId: target.sheetId,
          address: target.address,
          prior,
          breaks,
          dimensions,
        })
      if (!patch)
        throw new Error(
          t(
            command.endsWith('breaks:remove')
              ? 'appBreakNoneHere'
              : command.endsWith('breaks:insert')
                ? 'appBreaksNeedCell'
                : 'appActiveSheetUnavailable',
          ),
        )
      await this.enqueue([{ op: 'set_page_setup', sheetId: target.sheetId, ...patch }], frozen, {
        failureInDialog: true,
      })
    })()
    this.pageWork = work
    try {
      await work
    } finally {
      if (this.pageWork === work) this.pageWork = undefined
    }
  }
  private filePageSetup = new Map<string, unknown>()
  private viewports = new Map<string, { row: number; column: number }>()
  private projectionEpoch = 0
  private projectionDisposables: { dispose(): void }[] = []
  private loadTimer: ReturnType<typeof setTimeout> | undefined
  private themeObserver: MutationObserver
  private systemTheme = matchMedia('(prefers-color-scheme: dark)')
  private baseTheme!: ReturnType<ThemeService['getCurrentTheme']>
  constructor(
    private readonly container: HTMLElement,
    private readonly host: EditorHost,
    private readonly report: (error: unknown) => void,
    private readonly changed: () => void = () => {},
    locale = navigator.language,
    private appearance: WorkbookAppearance = 'vscode',
    private readonly clearError: () => void = () => {},
  ) {
    this.runtime = createUniver({
      locale: /^zh[-_](TW|HK|MO|Hant)/i.test(locale)
        ? LocaleType.ZH_TW
        : /^zh/i.test(locale)
          ? LocaleType.ZH_CN
          : LocaleType.EN_US,
      locales: {
        [LocaleType.EN_US]: {
          ...enUS,
          ...filterEnUS,
          ...sortEnUS,
          ...noteEnUS,
          ...tableEnUS,
          ...cfEnUS,
        },
        [LocaleType.ZH_CN]: {
          ...zhCN,
          ...filterZhCN,
          ...sortZhCN,
          ...noteZhCN,
          ...tableZhCN,
          ...cfZhCN,
        },
        [LocaleType.ZH_TW]: {
          ...zhTW,
          ...filterZhTW,
          ...sortZhTW,
          ...noteZhTW,
          ...tableZhTW,
          ...cfZhTW,
        },
      },
      darkMode: this.darkTheme(),
      presets: [
        UniverSheetsCorePreset({
          container,
          header: true,
          toolbar: false,
          contextMenu: true,
          formulaBar: true,
          footer: { sheetBar: true, statisticBar: true, menus: false, zoomSlider: false },
        }),
        UniverSheetsDrawingPreset(),
        UniverSheetsDataValidationPreset(),
        UniverSheetsFilterPreset(),
        UniverSheetsSortPreset(),
        UniverSheetsNotePreset(),
        UniverSheetsTablePreset(),
        UniverSheetsConditionalFormattingPreset(),
      ],
    })
    // The desktop's Excel keyboard layer: data-edge jumps, editor word jumps,
    // sheet-tab switching, Ctrl+Home/End, whole row/column selection,
    // hide/unhide and Insert/Delete Cells (forwarded below as session edits).
    registerExcelJumpNav(this.runtime)
    registerEditorWordJump(this.runtime)
    registerExcelShortcuts(this.runtime, (sheetId) => {
      const sheet = this.state?.sheets.find((item) => item.id === sheetId)
      return sheet && sheet.rows > 0 && sheet.columns > 0
        ? { row: sheet.rows - 1, column: sheet.columns - 1 }
        : null
    })
    // Visibility and Insert/Delete Cells chords act through session edits and
    // the shared dialog instead of Univer's local commands.
    const keyCommands = this.runtime.univer.__getInjector().get(ICommandService)
    const keyHandlers: Record<string, () => boolean> = {
      'genoffice.command.hide-selected-rows': () => this.axisVisibility('rows', true),
      'genoffice.command.unhide-selected-rows': () => this.axisVisibility('rows', false),
      'genoffice.command.hide-selected-cols': () => this.axisVisibility('cols', true),
      'genoffice.command.unhide-selected-cols': () => this.axisVisibility('cols', false),
      'genoffice.command.insert-cells': () => (this.commandSink?.('insert-cells'), true),
      'genoffice.command.delete-cells': () => (this.commandSink?.('delete-cells'), true),
    }
    for (const [id, handler] of Object.entries(keyHandlers)) {
      keyCommands.unregisterCommand(id)
      keyCommands.registerCommand({ id, type: CommandType.COMMAND, handler })
    }
    // The desktop's cell / header context menu; dialog entries reach the ribbon dispatcher.
    this.disposables.push(installCellContextMenu(this.runtime))
    setContextMenuCommandSink((command) => this.commandSink?.(command))
    this.disposables.push({
      dispose: registerPasteSpecialOpener(() =>
        this.commandSink?.(PASTE_SPECIAL_OPEN_SHELL_COMMAND),
      ),
    })
    const gutterHost = container.parentElement
    if (gutterHost?.classList.contains('workbook-content')) {
      const runGutter = (
        action: OutlineAction,
        axis: 'rows' | 'cols',
        detail: { group?: OutlineGroup; level?: number },
      ) => {
        const data = this.outlineProjection?.data
        if (data)
          void this.outlineCommand(action, axis, {
            stateId: data.stateId,
            sheetId: data.sheetId,
            ...detail,
          }).catch(this.report)
      }
      this.outlineGutter = installOutlineGutter({
        runtime: this.runtime,
        host: gutterHost,
        outline: (sheetId) =>
          !this.installing &&
          this.outlineProjection?.data.sheetId === sheetId &&
          this.outlineProjection.data.stateId === this.state?.stateId
            ? this.outlineProjection.model
            : null,
        groups: (worksheet, axis) => {
          const projection = this.outlineProjection
          if (!projection || projection.data.sheetId !== worksheet.getSheetId()) return []
          const entries = axis === 'rows' ? projection.model.rows : projection.model.cols
          return outlineGroups(
            entries,
            axis === 'rows' ? projection.model.summaryBelow : projection.model.summaryRight,
            (index) => entries.get(index)?.fileHidden === true,
          )
        },
        onToggleGroup: (axis, group, collapsed) =>
          runGutter(collapsed ? 'hide-detail' : 'show-detail', axis, { group }),
        onLevel: (axis, level) => runGutter('level', axis, { level }),
      })
      this.disposables.push(this.outlineGutter)
    }
    setVisualCommandListener((command, id) => {
      void this.visualCommand(command, id).catch(this.report)
    })
    this.disposables.push({ dispose: () => setVisualCommandListener(null) })
    setChartDialogListener((id, kind) => {
      if (kind === 'format' || kind === 'select-data') this.openChartDialog(id, kind)
    })
    const unsubscribeVisual = subscribeVisualSelection(() => this.changed())
    const unsubscribeChartElement = subscribeChartElementSelection(() => this.changed())
    this.disposables.push({
      dispose: () => {
        setChartDialogListener(null)
        unsubscribeChartElement()
        unsubscribeVisual()
      },
    })
    this.ribbonFocusBoundary = installRibbonFocusBoundary(
      this.runtime.univer.__getInjector().get(IRenderManagerService),
      container,
    )
    this.disposables.push(this.ribbonFocusBoundary)
    const disposePictureTransfer = installPictureTransfer(container, {
      pasteCells: (html, text) => {
        void this.pasteNativeCells(html, text).catch(this.report)
      },
      isCellEditing: () =>
        this.composing || this.runtime.univerAPI.getActiveWorkbook()?.isCellEditing() === true,
      isPasteTarget: (target) =>
        target instanceof Element &&
        container.contains(target) &&
        isGridKeyTarget(target, container),
      cellAtPoint: (x, y) => cellAtClientPoint(this.runtime, x, y),
      insert: (file, anchor) => this.insertTransferredPicture(file, anchor),
    })
    this.disposables.push({ dispose: disposePictureTransfer })
    const filterKey = (event: KeyboardEvent) => {
      if (
        event.defaultPrevented ||
        event.isComposing ||
        this.composing ||
        event.code !== 'KeyL' ||
        !(event.ctrlKey || event.metaKey) ||
        !event.shiftKey ||
        event.altKey ||
        this.runtime.univerAPI.getActiveWorkbook()?.isCellEditing() ||
        this.container.ownerDocument.querySelector('.dialog-backdrop')
      )
        return
      if (
        !(event.target instanceof Element) ||
        !this.container.closest('.workbook-editor')?.contains(event.target) ||
        !isGridKeyTarget(event.target, this.container)
      )
        return
      event.preventDefault()
      event.stopImmediatePropagation()
      void this.sortFilterCommand('filter-toggle').catch(this.report)
    }
    // Format Painter paints when the target selection is released, as Univer's painter does.
    const paintUp = (event: PointerEvent) => {
      if (
        !this.painter ||
        event.button !== 0 ||
        !(event.target instanceof Element) ||
        !this.container.contains(event.target) ||
        !isGridKeyTarget(event.target, this.container)
      )
        return
      setTimeout(() => void this.paintFormats().catch(this.report), 0)
    }
    window.addEventListener('pointerup', paintUp, true)
    this.disposables.push({ dispose: () => window.removeEventListener('pointerup', paintUp, true) })
    window.addEventListener('keydown', filterKey, true)
    this.disposables.push({ dispose: () => window.removeEventListener('keydown', filterKey, true) })
    const copyKey = (event: KeyboardEvent) => {
      if (
        !event.defaultPrevented &&
        !event.isComposing &&
        !this.composing &&
        event.code === 'Backquote' &&
        (event.ctrlKey || event.metaKey) &&
        !event.altKey &&
        !event.shiftKey &&
        event.target instanceof Element &&
        container.contains(event.target) &&
        isGridKeyTarget(event.target, container)
      ) {
        event.preventDefault()
        event.stopImmediatePropagation()
        void this.auditCommand('toggle-show-formulas').catch(this.report)
        return
      }
      if (
        !event.defaultPrevented &&
        !event.isComposing &&
        !this.composing &&
        event.code === 'F3' &&
        (event.ctrlKey || event.metaKey) &&
        !event.shiftKey &&
        !event.altKey &&
        event.target instanceof Element &&
        container.contains(event.target) &&
        isGridKeyTarget(event.target, container)
      ) {
        event.preventDefault()
        event.stopImmediatePropagation()
        void this.openNameManager().catch(this.report)
        return
      }
      if (
        !event.defaultPrevented &&
        !event.isComposing &&
        !this.composing &&
        event.code === 'F3' &&
        event.shiftKey &&
        !event.ctrlKey &&
        !event.metaKey &&
        !event.altKey &&
        event.target instanceof Element &&
        container.contains(event.target) &&
        isGridKeyTarget(event.target, container)
      ) {
        event.preventDefault()
        event.stopImmediatePropagation()
        void this.openInsertFunction('All').catch(this.report)
        return
      }

      if (
        (event.ctrlKey || event.metaKey) &&
        event.altKey &&
        event.key.toLowerCase() === 'v' &&
        !event.isComposing &&
        !this.composing &&
        event.target instanceof Element &&
        container.contains(event.target) &&
        isGridKeyTarget(event.target, container) &&
        !this.runtime.univerAPI.getActiveWorkbook()?.isCellEditing()
      ) {
        event.preventDefault()
        event.stopImmediatePropagation()
        void this.openPasteSpecial().catch(this.report)
        return
      }
      if (
        event.key === 'Escape' &&
        this.painter &&
        event.target instanceof Element &&
        container.contains(event.target) &&
        isGridKeyTarget(event.target, container) &&
        !this.runtime.univerAPI.getActiveWorkbook()?.isCellEditing()
      ) {
        event.preventDefault()
        event.stopImmediatePropagation()
        this.painter = null
        this.changed()
        return
      }
      if (
        event.key === 'Escape' &&
        this.copiedCut &&
        event.target instanceof Element &&
        container.contains(event.target) &&
        isGridKeyTarget(event.target, container) &&
        !this.runtime.univerAPI.getActiveWorkbook()?.isCellEditing()
      ) {
        const clipboard = this.runtime.univer.__getInjector().get(ISheetClipboardService)
        this.consumedCutId = this.copiedCut.id
        clipboard.copyContentCache().del(this.copiedCut.id)
        clipboard.removeMarkSelection()
        this.copiedCut = undefined
        this.copiedBlock = undefined
        return
      }
      if (
        event.defaultPrevented ||
        event.isComposing ||
        this.composing ||
        !(event.ctrlKey || event.metaKey) ||
        event.altKey ||
        !['c', 'x'].includes(event.key.toLowerCase()) ||
        this.runtime.univerAPI.getActiveWorkbook()?.isCellEditing() ||
        selectedVisualIds().length
      )
        return
      if (
        !(event.target instanceof Element) ||
        !container.contains(event.target) ||
        !isGridKeyTarget(event.target, container)
      )
        return
      event.preventDefault()
      event.stopImmediatePropagation()
      void this.copyClipboard(event.key.toLowerCase() === 'x' ? 'CUT' : 'COPY').catch(this.report)
    }
    window.addEventListener('keydown', copyKey, true)
    this.disposables.push({ dispose: () => window.removeEventListener('keydown', copyKey, true) })
    const calculationKey = (event: KeyboardEvent) => {
      if (
        event.code !== 'F9' ||
        event.defaultPrevented ||
        event.ctrlKey ||
        event.metaKey ||
        event.altKey ||
        this.runtime.univerAPI.getActiveWorkbook()?.isCellEditing()
      )
        return
      const target = event.target
      if (
        !(target instanceof Element) ||
        !this.container.closest('.workbook-editor')?.contains(target) ||
        target.closest(
          'input,textarea,select,[contenteditable="true"],[role="dialog"],.dialog-overlay',
        )
      )
        return
      event.preventDefault()
      event.stopImmediatePropagation()
      try {
        if (event.shiftKey) this.recalculateSheet()
        else this.recalculate()
      } catch (error) {
        this.report(error)
      }
    }
    window.addEventListener('keydown', calculationKey, true)
    this.disposables.push({
      dispose: () => window.removeEventListener('keydown', calculationKey, true),
    })
    this.findReplace = new HostFindReplaceService({
      search: (request) => {
        if (!this.host.search)
          return Promise.reject(new Error('Indexed document search is unavailable.'))
        return this.host.search(request)
      },
      getContext: () => {
        const book = this.runtime.univerAPI.getActiveWorkbook(),
          sheet = book?.getActiveSheet(),
          range = sheet?.getActiveRange()?.getRange()
        return book && sheet && this.state
          ? {
              stateId: this.state.stateId,
              sheetId: sheet.getSheetId(),
              unitId: book.getId(),
              ...(range ? { address: formatAddress(range.startRow, range.startColumn) } : {}),
            }
          : null
      },
      flush: () => this.flushEditor(),
      jump: async (sheetId, address, isCurrent) => {
        const sheet = this.state?.sheets.find((sheet) => sheet.id === sheetId)
        if (!sheet) throw new Error('The search result sheet no longer exists.')
        const error = await this.goToReference(
          "'" + sheet.name.replace(/'/g, "''") + "'!" + address,
          { focus: false, isCurrent },
        )
        if (error) throw new Error(error)
      },
      report: this.report,
    })
    this.baseTheme = this.runtime.univer.__getInjector().get(ThemeService).getCurrentTheme()
    this.setAppearance(appearance)
    this.themeObserver = new MutationObserver(this.updateTheme)
    this.themeObserver.observe(document.body, {
      attributes: true,
      attributeFilter: ['class', 'data-vscode-theme-kind'],
    })
    this.systemTheme.addEventListener('change', this.updateTheme)
    void this.runtime.univer
      .__getInjector()
      .get(LifecycleService)
      .onStage(LifecycleStages.Rendered)
      .then(() => {
        if (!this.closed) {
          this.installFilterPopupCommands()
          this.installContextMenuCommands()
        }
      })
      .catch((error) => {
        if (!this.closed) this.report(error)
      })
    const commands = this.runtime.univer.__getInjector().get(ICommandService)
    this.disposables.push(
      commands.onCommandExecuted((command) => {
        if (
          command.id.startsWith('sheet.operation.set-cell-edit-visible') &&
          (command.params as { visible?: boolean } | undefined)?.visible
        )
          this.ribbonFocusBoundary?.editStarted()
      }),
      commands.beforeCommandExecuted((command, options) => {
        if (this.internal) return
        if (
          [
            DeltaColumnWidthCommand.id,
            DeltaRowHeightCommand.id,
            SetColWidthCommand.id,
            SetRowHeightCommand.id,
          ].includes(command.id)
        ) {
          const book = this.runtime.univerAPI.getActiveWorkbook()
          const params = command.params as {
            unitId?: string
            subUnitId?: string
            anchorCol?: number
            anchorRow?: number
            deltaX?: number
            deltaY?: number
            value?: number
            ranges?: GridRange[]
          }
          const sheet = params.subUnitId
            ? book?.getSheetBySheetId(params.subUnitId)
            : book?.getActiveSheet()
          if (
            !sheet ||
            (params.unitId && params.unitId !== book?.getId()) ||
            !this.state?.sheets.some((item) => item.id === sheet.getSheetId())
          )
            throw new Error('invalid_manual_target')
          const row =
            command.id === DeltaRowHeightCommand.id || command.id === SetRowHeightCommand.id
          const delta =
            command.id === DeltaRowHeightCommand.id || command.id === DeltaColumnWidthCommand.id
          let ranges = params.ranges ?? [sheet.getActiveRange()!.getRange()]
          let pixels = params.value!
          if (delta) {
            const anchor = row ? params.anchorRow! : params.anchorCol!
            if (!Number.isInteger(anchor) || anchor < 0) throw new Error('invalid_manual_target')
            pixels = row
              ? sheet.getSheet().getRowHeight(anchor) + params.deltaY!
              : sheet.getSheet().getColumnWidth(anchor) + params.deltaX!
            const selections =
              this.runtime.univer
                .__getInjector()
                .get(SheetsSelectionsService)
                .getCurrentSelections() ?? []
            const axisRanges = selections
              .map((item) => item.range)
              .filter(
                (range) =>
                  range.rangeType === (row ? RANGE_TYPE.ROW : RANGE_TYPE.COLUMN) ||
                  range.rangeType === RANGE_TYPE.ALL,
              )
            const includesAnchor = axisRanges.some((range) =>
              row
                ? range.startRow <= anchor && range.endRow >= anchor
                : range.startColumn <= anchor && range.endColumn >= anchor,
            )
            ranges = includesAnchor
              ? axisRanges
              : [
                  {
                    startRow: row ? anchor : 0,
                    endRow: row ? anchor : 0,
                    startColumn: row ? 0 : anchor,
                    endColumn: row ? 0 : anchor,
                  },
                ]
          }
          const operations = dimensionIntents(
            sheet.getSheetId(),
            row ? 'row' : 'column',
            ranges,
            pixels,
          )
          void this.enqueue(operations).catch(this.report)
          throw new Error('manual_intent_forwarded')
        }
        if (
          command.id === 'sheet.mutation.set-range-values' &&
          options?.fromFormula &&
          options?.applyFormulaCalculationResult
        ) {
          const params = command.params as {
              unitId?: string
              subUnitId?: string
              cellValue?: Record<string, Record<string, unknown>>
            },
            book = this.runtime.univerAPI.getActiveWorkbook(),
            sheet = params.subUnitId ? book?.getSheetBySheetId(params.subUnitId) : undefined
          if (!sheet || params.unitId !== book?.getId() || !params.cellValue)
            throw new Error('invalid_formula_projection')
          for (const [row, columns] of Object.entries(params.cellValue))
            for (const column of Object.keys(columns))
              if (
                !/^\d+$/.test(row) ||
                !/^\d+$/.test(column) ||
                !sheet.getRange(Number(row), Number(column)).getFormula()
              )
                throw new Error('invalid_formula_projection')
          return
        }
        if (HIDE_COMMANDS.has(command.id)) {
          const params = command.params as { subUnitId?: string; ranges?: GridRange[] }
          const sheetId =
            params.subUnitId ??
            this.runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId()
          if (!sheetId || !this.state?.sheets.some((sheet) => sheet.id === sheetId))
            throw new Error('invalid_manual_target')
          const rows = command.id.includes('row')
          const hidden = command.id.endsWith('-hidden')
          const operations: WorkbookOperation[] = (params.ranges ?? []).map((range) =>
            rows
              ? {
                  op: 'set_rows_hidden',
                  sheetId,
                  row: range.startRow + 1,
                  count: range.endRow - range.startRow + 1,
                  hidden,
                }
              : {
                  op: 'set_cols_hidden',
                  sheetId,
                  column: columnLabel(range.startColumn),
                  count: range.endColumn - range.startColumn + 1,
                  hidden,
                },
          )
          void this.enqueue(operations).catch(this.report)
          throw new Error('manual_intent_forwarded')
        }
        if (command.id === 'sheet.command.set-style') {
          const workbook = this.runtime.univerAPI.getActiveWorkbook(),
            sheet = workbook?.getActiveSheet(),
            params = command.params as {
              unitId?: string
              subUnitId?: string
              range?: GridRange
              style: { type: string; value: unknown }
            }
          const targetSheetId = params.subUnitId ?? sheet?.getSheetId(),
            range = params.range ?? sheet?.getActiveRange()?.getRange()
          if (
            !range ||
            !targetSheetId ||
            (params.unitId && params.unitId !== workbook?.getId()) ||
            !this.state?.sheets.some((candidate) => candidate.id === targetSheetId)
          )
            throw new Error('invalid_manual_target')
          const op = manualStyleIntent(targetSheetId, range, params.style)
          void this.enqueue([op]).catch(this.report)
          throw new Error('manual_intent_forwarded')
        }
        if (command.id === 'sheet.command.set-range-values') {
          const workbook = this.runtime.univerAPI.getActiveWorkbook(),
            sheet = workbook?.getActiveSheet()
          if (!sheet) throw new Error('view_unavailable')
          const params = command.params as {
            value: unknown
            range?: GridRange
            subUnitId?: string
            unitId?: string
          }
          const targetSheetId = params.subUnitId ?? sheet.getSheetId()
          if (
            (params.unitId && params.unitId !== workbook!.getId()) ||
            !this.state?.sheets.some((candidate) => candidate.id === targetSheetId)
          )
            throw new Error('invalid_manual_target')
          const operations = manualValueIntents(
            targetSheetId,
            {
              ...params,
              range: params.range ?? sheet.getActiveRange()?.getRange(),
            },
            workbook!.getSheetBySheetId(targetSheetId)!.getSheet().getMergeData(),
          )
          void this.enqueue(operations).catch(this.report)
          throw new Error('manual_intent_forwarded')
        }
        if (command.id === UndoCommandId || command.id === RedoCommandId) {
          void this.requestHistory(command.id === UndoCommandId ? 'undo' : 'redo').catch(
            this.report,
          )
          throw new Error('history_forwarded')
        }
        if (
          command.id.includes('set-formula-calculation') ||
          command.id.includes('set-formula-result')
        )
          return
        // Note popups edit locally while open; the final text commits on close.
        if (
          command.id === 'sheet.mutation.update-note' ||
          command.id === 'sheet.mutation.remove-note'
        ) {
          const params = command.params as { sheetId?: string; row?: number; col?: number }
          if (
            !this.notesApplying &&
            params.sheetId !== undefined &&
            params.row !== undefined &&
            params.col !== undefined
          )
            this.dirtyNotes.set(`${params.sheetId}!${params.row}:${params.col}`, {
              sheetId: params.sheetId,
              row: params.row,
              column: params.col,
            })
          return
        }
        // Showing or pinning a note popup is display state, as on the desktop.
        if (
          command.id === 'sheet.mutation.toggle-note-popup' ||
          command.id === 'sheet.mutation.update-note-position'
        )
          return
        if (CF_MUTATIONS.has(command.id)) {
          const subUnitId = (command.params as { subUnitId?: string } | undefined)?.subUnitId
          if (!subUnitId || !this.state?.sheets.some((sheet) => sheet.id === subUnitId))
            throw new Error('invalid_manual_target')
          this.captureConditionalFormats(subUnitId)
          return
        }
        if (command.type === CommandType.MUTATION && command.id.startsWith('sheet.'))
          throw new Error('unsupported_manual_command')
      }),
    )
    this.disposables.push(
      this.runtime.univerAPI.addEvent(this.runtime.univerAPI.Event.Scroll, ({ worksheet }) => {
        // This event is emitted by the mounted scroll manager. A newly created
        // workbook has no manager yet, and starts at the origin instead.
        if (this.closed || this.internal) return
        const scroll = worksheet.getScrollState()
        this.viewports.set(worksheet.getSheetId(), {
          row: scroll.sheetViewStartRow,
          column: scroll.sheetViewStartColumn,
        })
        this.scheduleLoad()
      }),
    )
    this.disposables.push(
      this.runtime.univerAPI.addEvent(this.runtime.univerAPI.Event.ActiveSheetChanged, () => {
        if (this.internal || this.installing) return
        this.syncFormulaView()
        this.installProjectedSparklines()
        this.outlineProjection = undefined
        this.outlineGutter?.refresh()
        void this.loadOutlineProjection().catch(this.report)
        this.scheduleLoad()
        this.changed()
      }),
    )
    this.container.addEventListener('compositionstart', this.onCompositionStart)
    this.container.addEventListener('compositionend', this.onCompositionEnd)
  }
  private onCompositionStart = () => {
    this.composing = true
  }
  private darkTheme() {
    if (this.appearance === 'excel') return false
    const classes = document.body.classList
    if (classes.contains('vscode-dark') || classes.contains('vscode-high-contrast')) return true
    if (classes.contains('vscode-light') || classes.contains('vscode-high-contrast-light'))
      return false
    return this.systemTheme.matches
  }
  private updateTheme = () => {
    this.runtime.univerAPI.toggleDarkMode(this.darkTheme())
  }
  setAppearance(appearance: WorkbookAppearance) {
    this.appearance = appearance
    const service = this.runtime.univer.__getInjector().get(ThemeService)
    const accent = getComputedStyle(this.container).getPropertyValue('--workbook-accent').trim()
    service.setTheme(
      appearance === 'excel'
        ? {
            ...this.baseTheme,
            primary: {
              ...this.baseTheme.green,
              ...(accent ? { 500: accent, 600: accent, 700: accent } : {}),
            },
          }
        : this.baseTheme,
    )
    this.updateTheme()
  }
  get appearanceState() {
    return {
      appearance: this.appearance,
      darkMode: this.runtime.univer.__getInjector().get(ThemeService).darkMode,
    }
  }
  private onCompositionEnd = () => {
    this.composing = false
  }
  async install(stateId: string, revision: number, data: unknown) {
    while (
      !this.closed &&
      (isVisualDragActive() || isChartEditorOpen() || (isShapeTextEditorOpen() && !this.visualWork))
    )
      await new Promise<void>((resolve) => setTimeout(resolve, 16))
    const state = workbookStateSchema.parse(data)
    if (state.stateId !== stateId || this.closed) throw new Error('invalid_projection')
    this.installing++
    this.container.setAttribute('aria-busy', 'true')
    try {
      const active = this.runtime.univerAPI.getActiveWorkbook(),
        sheetId = active?.getActiveSheet()?.getSheetId(),
        selected = active?.getActiveSheet()?.getActiveRange()?.getRange()
      const viewport = sheetId ? this.viewports.get(sheetId) : undefined
      this.closeFilterPopup()
      this.projectionEpoch++
      this.cfSignatures.clear()
      this.clearChartVisuals()
      this.clearTraceArrows()
      for (const disposable of this.projectionDisposables) disposable.dispose()
      this.projectionDisposables = []
      this.outlineProjection = undefined
      this.outlineGutter?.refresh()
      this.state = state
      // The saved calculation mode (an unsaved Calculation Options edit wins over the file's calcPr).
      const savedMode =
        (state.features?.calcModeEdits as { mode?: string } | undefined)?.mode ??
        (state.features?.originalMetadata as { calcMode?: string } | undefined)?.calcMode
      setManualCalculation(this.runtime, savedMode === 'manual')
      this.watchValues.clear()
      this.revision = revision
      this.loaded.clear()
      this.viewports.clear()
      this.internal++
      try {
        if (active)
          this.runtime.univer
            .__getInjector()
            .get(IUniverInstanceService)
            .disposeUnit(active.getId())
        const sheets: IWorkbookData['sheets'] = Object.fromEntries(
          state.sheets.map((sheet) => {
            const metadata = sheet.features?.originalMetadata as
              | {
                  columnWidths?: {
                    startColumn: number
                    endColumn: number
                    width?: number
                    hidden: boolean
                  }[]
                  defaultRowHeight?: number | null
                  defaultColumnWidth?: number | null
                  freeze?: { frozenRows: number; frozenColumns: number } | null
                  showGridLines?: boolean
                  showRowColHeaders?: boolean
                  zoomScale?: number
                }
              | undefined
            const page = sheet.features?.pageSetupEdits as
              | {
                  showGridlines?: boolean
                  showHeadings?: boolean
                  zoomScale?: number
                  rightToLeft?: boolean
                }
              | undefined
            const freeze = (sheet.features?.freezeEdits ?? metadata?.freeze) as
              { frozenRows: number; frozenColumns: number } | undefined
            const originalColumns = new Map<number, { w?: number; hd?: 0 | 1 }>()
            for (const span of metadata?.columnWidths ?? [])
              for (
                let column = span.startColumn;
                column <= Math.min(16383, span.endColumn);
                column++
              )
                originalColumns.set(column, {
                  ...(span.width === undefined
                    ? {}
                    : {
                        w: axisColumnWidthToPixels(
                          span.width,
                          measureNormalFontMdw(
                            (state.features?.originalMetadata ?? {}) as NormalFontDescriptor,
                          ),
                        ),
                      }),
                  hd: span.hidden ? 1 : 0,
                })
            const columnData: NonNullable<IWorkbookData['sheets'][string]['columnData']> = {}
            const rowData: NonNullable<IWorkbookData['sheets'][string]['rowData']> =
              Object.fromEntries(
                Object.entries((sheet.features?.rowHeights ?? {}) as Record<string, number>).map(
                  ([row, points]) => [Number(row) - 1, { h: (points * 4) / 3, ia: 0 }],
                ),
              )
            for (
              let column = 0;
              column <
              Math.min(
                16384,
                Math.max(
                  20,
                  sheet.columns + 5,
                  sheet.id === sheetId ? (selected?.endColumn ?? -1) + 1 : 0,
                ),
              );
              column++
            ) {
              const original = originalAxisIndex(state.structuralLog, sheet.id, column, 'column')
              if (original !== null && originalColumns.has(original))
                columnData[column] = originalColumns.get(original)!
            }
            for (const [column, width] of Object.entries(
              (sheet.features?.colWidths ?? {}) as Record<string, number>,
            ))
              columnData[parseAddress(`${column}1`).column] = {
                ...columnData[parseAddress(`${column}1`).column],
                w: width,
              }
            const visibility = visibilityEditsSchema.parse(sheet.features?.visibilityEdits ?? [])
            for (const edit of visibility)
              for (let index = edit.start; index < edit.start + edit.count; index++) {
                if (edit.axis === 'rows')
                  rowData[index] = { ...rowData[index], hd: edit.hidden ? 1 : 0 }
                else columnData[index] = { ...columnData[index], hd: edit.hidden ? 1 : 0 }
              }
            const cellData: NonNullable<IWorkbookData['sheets'][string]['cellData']> = {}
            for (const [address, cell] of Object.entries(sheet.cells)) {
              const point = parseAddress(address)
              ;(cellData[point.row] ??= {})[point.column] = gridCell(cell)
            }
            return [
              sheet.id,
              {
                id: sheet.id,
                name: sheet.name,
                hidden: (sheet.features?.originalMetadata as { hidden?: boolean } | undefined)
                  ?.hidden
                  ? 1
                  : 0,
                rowCount: Math.min(
                  1_048_576,
                  Math.max(
                    1000,
                    sheet.rows + 100,
                    sheet.id === sheetId ? (selected?.endRow ?? -1) + 1 : 0,
                  ),
                ),
                columnCount: Math.min(16_384, Math.max(20, sheet.columns + 5)),
                rowData,
                columnData,
                defaultRowHeight: ((metadata?.defaultRowHeight || 15) * 4) / 3,
                showGridlines: (page?.showGridlines ?? metadata?.showGridLines ?? true) ? 1 : 0,
                rowHeader: {
                  width: 46,
                  hidden: (page?.showHeadings ?? metadata?.showRowColHeaders ?? true) ? 0 : 1,
                },
                columnHeader: {
                  height: 20,
                  hidden: (page?.showHeadings ?? metadata?.showRowColHeaders ?? true) ? 0 : 1,
                },
                zoomRatio: (page?.zoomScale ?? metadata?.zoomScale ?? 100) / 100,
                rightToLeft:
                  (page?.rightToLeft ??
                  (metadata as { rightToLeft?: boolean } | undefined)?.rightToLeft)
                    ? 1
                    : 0,
                ...(freeze
                  ? {
                      freeze: {
                        xSplit: freeze.frozenColumns,
                        ySplit: freeze.frozenRows,
                        startRow: freeze.frozenRows,
                        startColumn: freeze.frozenColumns,
                      },
                    }
                  : {}),
                ...(metadata?.defaultColumnWidth
                  ? {
                      defaultColumnWidth: axisColumnWidthToPixels(
                        metadata.defaultColumnWidth,
                        measureNormalFontMdw(
                          (state.features?.originalMetadata ?? {}) as NormalFontDescriptor,
                        ),
                      ),
                    }
                  : {}),
                cellData,
              },
            ]
          }),
        )
        const workbook = this.runtime.univerAPI.createWorkbook({
          id: 'projection',
          sheetOrder: state.sheets.map((sheet) => sheet.id),
          sheets,
        })
        this.projectionDisposables.push(
          installFormulaViewInterceptor(this.runtime, { current: null }, (sheetId) =>
            this.showFormulasFor(sheetId),
          ),
        )
        const defined = sessionDefinedNames(state).names.map((name) => ({
          name: name.name,
          formula: name.formula,
          ...(name.scopeSheetId
            ? { sheetIndex: state.sheets.findIndex((sheet) => sheet.id === name.scopeSheetId) }
            : {}),
        }))
        installNameDefinitions(
          workbook,
          orderDefinedNames(defined).map((name) => ({
            name: name.name,
            formula: name.formula,
            ...(name.sheetIndex === undefined
              ? {}
              : { sheetId: state.sheets[name.sheetIndex]!.id }),
          })),
        )
        this.syncFormulaView()
        // Clipboard controllers are registered by the workbook's UI plugin.
        // Hidden toolbars do not instantiate their default copy/paste hooks.
        this.runtime.univer.__getInjector().get(SheetClipboardController)
        if (sheetId && state.sheets.some((sheet) => sheet.id === sheetId)) {
          const sheet = workbook.setActiveSheet(sheetId)
          if (selected)
            sheet.setActiveRange(
              sheet.getRange(
                selected.startRow,
                selected.startColumn,
                selected.endRow - selected.startRow + 1,
                selected.endColumn - selected.startColumn + 1,
              ),
            )
        }
        this.runtime.univer.__getInjector().get(IUndoRedoService).clearUndoRedo(workbook.getId())
        this.projectionDisposables.push(
          workbook.onSelectionChange((ranges) => {
            if (this.internal || this.installing) return
            const sheet = workbook.getActiveSheet()
            if (sheet) this.host.selection(sheet.getSheetId(), ranges[0] ?? null)
            this.findReplace.contextChanged()
            this.changed()
          }),
        )
      } finally {
        this.internal--
      }
      await new Promise<void>((resolve) =>
        requestAnimationFrame(() => requestAnimationFrame(() => resolve())),
      )
      if (this.closed || this.state?.stateId !== stateId) return
      this.projectionDisposables.push(
        installFxButtonHook(this.runtime, () => {
          void this.openInsertFunction('All').catch(this.report)
        }),
      )
      if (viewport && sheetId && state.sheets.some((sheet) => sheet.id === sheetId)) {
        this.viewports.set(sheetId, viewport)
        this.runtime.univerAPI
          .getActiveWorkbook()
          ?.getActiveSheet()
          ?.scrollToCell(viewport.row, viewport.column, 0)
      }
      this.installFilterPopupCommands()
      await this.loadVisible()
      this.installChartVisuals()
      void this.loadOutlineProjection().catch(this.report)
      this.findReplace.contextChanged()
      this.applyProjectedThreads()
      this.applyProjectedNotes()
      await this.registerProjectedTables()
      this.installProjectedSparklines()
      this.refreshViewOverlays()
      this.changed()
    } finally {
      this.installing--
      if (!this.installing) this.container.setAttribute('aria-busy', 'false')
    }
  }
  private scheduleLoad() {
    if (this.loadTimer) clearTimeout(this.loadTimer)
    this.loadTimer = setTimeout(() => {
      void this.loadVisible()
        .then(() => {
          const id = this.runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId()
          if (id !== this.chartVisualSheet) this.installChartVisuals()
        })
        .catch(this.report)
    }, 80)
  }
  private async loadVisible() {
    const epoch = this.projectionEpoch
    const sheet = this.runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
    if (!sheet || this.closed || !this.state) return
    for (let pass = 0; pass < 100; pass++) {
      if (epoch !== this.projectionEpoch || this.closed) return
      const visible = sheet.getVisibleRangesOfAllViewports()
      const pages = new Map<string, { row: number; column: number }>()
      for (const range of visible?.values() ?? []) {
        for (let row = Math.floor(range.startRow / 50) * 50; row <= range.endRow; row += 50)
          for (
            let column = Math.floor(range.startColumn / 40) * 40;
            column <= range.endColumn;
            column += 40
          ) {
            pages.set(`${row}:${column}`, { row, column })
            if (pages.size > 100)
              throw new Error('viewport_limit: increase zoom or row/column sizes')
          }
      }
      if (!pages.size) {
        const scroll = this.viewports.get(sheet.getSheetId())
        pages.set('origin', { row: scroll?.row ?? 0, column: scroll?.column ?? 0 })
      }
      const pending = [...pages.values()].filter(
        (page) =>
          !this.loaded.has(
            `${this.state!.stateId}:${sheet.getSheetId()}:${page.row}:${page.column}`,
          ),
      )
      if (!pending.length) {
        if (!this.installing) this.changed()
        return
      }
      for (const page of pending) {
        if (epoch !== this.projectionEpoch || this.closed) return
        await this.loadPage(page.row, page.column)
      }
      // Loading original row heights / hidden rows can change the viewport.
      await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()))
    }
    throw new Error('viewport_limit: increase zoom or row/column sizes')
  }
  private async loadPage(pageRow: number, pageColumn: number) {
    const epoch = this.projectionEpoch
    const state = this.state,
      sheet = this.runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
    if (!state || !sheet || this.closed) return
    const startRow = Math.max(0, pageRow),
      startColumn = Math.max(0, pageColumn)
    const range = {
      startRow,
      endRow: Math.min(sheet.getMaxRows() - 1, startRow + 49),
      startColumn,
      endColumn: Math.min(sheet.getMaxColumns() - 1, startColumn + 39),
    }
    const key = `${state.stateId}:${sheet.getSheetId()}:${startRow}:${startColumn}`
    if (this.loaded.has(key)) return
    let result: {
      dataValidations?: WorkbookRangeResult['dataValidations']
      pageSetup?: unknown
      pageBreaks?: { rowBreaks: number[]; colBreaks: number[] }
      autoFilter?: WorkbookRangeResult['autoFilter']
      autoFilterColumns?: WorkbookRangeResult['autoFilterColumns']
      cells: ViewCell[]
      merges?: GridRange[]
      rows?: { row: number; height?: number; hidden?: boolean; customHeight?: boolean }[]
    }
    const started = Date.now()
    for (;;) {
      try {
        result = (await this.host.read(sheet.getSheetId(), range, state.stateId)) as {
          cells: ViewCell[]
          dataValidations?: WorkbookRangeResult['dataValidations']
          pageSetup?: unknown
          merges?: GridRange[]
          rows?: { row: number; height?: number; hidden?: boolean; customHeight?: boolean }[]
        }
        break
      } catch (error) {
        if (this.closed || epoch !== this.projectionEpoch) return
        if (
          !(
            error &&
            typeof error === 'object' &&
            'code' in error &&
            error.code === 'index_pending'
          ) ||
          Date.now() - started >= 30_000
        )
          throw error
        await new Promise((resolve) => setTimeout(resolve, 80))
      }
    }
    if (epoch !== this.projectionEpoch || this.state?.stateId !== state.stateId || this.closed)
      return
    if (Object.hasOwn(result, 'pageSetup'))
      this.filePageSetup.set(sheet.getSheetId(), result.pageSetup)
    if (result.pageBreaks) this.filePageBreaks.set(sheet.getSheetId(), result.pageBreaks)
    const cells: Record<number, Record<number, ICellData>> = {}
    for (const cell of result.cells) {
      const point = parseAddress(cell.address)
      ;(cells[point.row] ??= {})[point.column] = gridCell(cell)
    }
    this.internal++
    try {
      if (!this.loadedDataValidations.has(sheet.getSheetId())) {
        const commandService = this.runtime.univer.__getInjector().get(ICommandService)
        for (const [index, sourceRule] of (result.dataValidations ?? []).entries()) {
          const rule = workbookDvRuleToUniver(
            sourceRule,
            `file-dv-${sheet.getSheetId()}-${index}`,
            this.dvListSourceContext(state, sheet.getSheetId()),
          )
          if (rule)
            commandService.syncExecuteCommand('data-validation.mutation.addRule', {
              unitId: sheet.getWorkbook().getUnitId(),
              subUnitId: sheet.getSheetId(),
              rule,
            })
        }
        this.loadedDataValidations.add(sheet.getSheetId())
      }
      const cfResult = result as {
        conditionalRules?: WorkbookRangeResult['conditionalRules']
        sessionConditionalFormats?: unknown[] | null
      }
      if (
        !this.fileCfRules.has(sheet.getSheetId()) &&
        cfResult.sessionConditionalFormats === null &&
        cfResult.conditionalRules
      )
        this.fileCfRules.set(sheet.getSheetId(), cfResult.conditionalRules)
      this.projectConditionalFormats(sheet.getSheetId())
      const canonicalSheet = state.sheets.find((item) => item.id === sheet.getSheetId())!
      const heights = (canonicalSheet.features?.rowHeights ?? {}) as Record<string, number>
      const filter = canonicalSheet.features?.filterEdits as
        { hiddenRows: number[]; visibilityRange: GridRange } | undefined
      const visibility = visibilityEditsSchema.parse(canonicalSheet.features?.visibilityEdits ?? [])
      const filterRange = result.autoFilter
      const filterService = this.runtime.univer.__getInjector().get(SheetsFilterService)
      const unitId = sheet.getWorkbook().getUnitId()
      if (filterRange && !filterService.getFilterModel(unitId, sheet.getSheetId()))
        this.runtime.univer
          .__getInjector()
          .get(ICommandService)
          .syncExecuteCommand(SetSheetsFilterRangeMutation.id, {
            unitId,
            subUnitId: sheet.getSheetId(),
            range: filterRange,
          })
      const filterModel = filterService.getFilterModel(unitId, sheet.getSheetId())
      if (filterModel && filterRange) {
        filterModel.setRange(filterRange)
        for (const column of result.autoFilterColumns ?? []) {
          const criteria = filterColumnToUniver(filterRange, column)
          if (criteria) filterModel.setCriteria(criteria.colId, criteria, false)
        }
      }
      const filterHidden = filterModel ? new Set(filterModel.filteredOutRows) : new Set<number>()
      if (filterModel && filter) {
        filterHidden.clear()
        for (const row of filter.hiddenRows)
          if (axisHiddenOverride(visibility, 'rows', row) === undefined) filterHidden.add(row)
      }
      for (const row of result.rows ?? []) {
        const override = axisHiddenOverride(visibility, 'rows', row.row)
        const claimed =
          filterRange && row.row > filterRange.startRow && row.row <= filterRange.endRow
        if (filterModel && claimed && !filter) {
          if (row.hidden && override === undefined) filterHidden.add(row.row)
          else filterHidden.delete(row.row)
        }
        const hidden = override ?? (claimed ? false : row.hidden === true)
        // Filter visibility belongs to the filter model; hd is reserved for
        // manual/outline visibility so clearing criteria can reveal its rows.
        const manuallyHidden = sheet.getSheet().getRowRawVisible(row.row) === false
        if (hidden && !manuallyHidden) sheet.hideRows(row.row, 1)
        if (!hidden && manuallyHidden) sheet.showRows(row.row, 1)
        const height = heights[String(row.row + 1)] ?? row.height
        if (height && height > 0) {
          if (heights[String(row.row + 1)] !== undefined || row.customHeight)
            sheet.setRowHeightsForced(row.row, 1, (height * 4) / 3)
          else sheet.setRowHeights(row.row, 1, (height * 4) / 3)
        }
      }
      if (filterModel) filterModel.filteredOutRows = filterHidden
      const mergeKey = (merge: GridRange) =>
        [merge.startRow, merge.startColumn, merge.endRow, merge.endColumn].join(':')
      const installed = new Set(sheet.getSheet().getMergeData().map(mergeKey))
      const merges = (result.merges ?? []).filter((merge) => !installed.has(mergeKey(merge)))
      if (merges.length)
        this.runtime.univer
          .__getInjector()
          .get(ICommandService)
          .syncExecuteCommand(AddWorksheetMergeMutation.id, {
            unitId: sheet.getWorkbook().getUnitId(),
            subUnitId: sheet.getSheetId(),
            ranges: merges,
          })
      // Include off-page anchors without expanding a sparse page to the grid.
      const startRow = Math.min(
        range.startRow,
        ...result.cells.map((cell) => parseAddress(cell.address).row),
      )
      const startColumn = Math.min(
        range.startColumn,
        ...result.cells.map((cell) => parseAddress(cell.address).column),
      )
      sheet
        .getRange(
          startRow,
          startColumn,
          range.endRow - startRow + 1,
          range.endColumn - startColumn + 1,
        )
        .setValues(cells)
      this.syncFormulaView()
      this.loaded.add(key)
      this.runtime.univer
        .__getInjector()
        .get(IUndoRedoService)
        .clearUndoRedo(sheet.getWorkbook().getUnitId())
      this.runtime.univerAPI.getFormula().executeCalculation()
    } finally {
      this.internal--
    }
  }
  get selection() {
    const sheet = this.runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
    const range = sheet?.getActiveRange()?.getRange()
    return sheet && range
      ? {
          sheetId: sheet.getSheetId(),
          range,
          address: `${formatAddress(range.startRow, range.startColumn)}:${formatAddress(range.endRow, range.endColumn)}`,
          style: sheet.getActiveRange()?.getCellStyleData() ?? {},
        }
      : null
  }
  get selectionContext() {
    const selection = this.selection
    return {
      stateId: this.state?.stateId ?? '',
      revision: this.revision,
      selection: selection ? { sheetId: selection.sheetId, address: selection.address } : null,
    }
  }
  get consolidateDialogContext() {
    const selection = this.selection,
      sheet = selection
        ? this.runtime.univerAPI.getActiveWorkbook()?.getSheetBySheetId(selection.sheetId)
        : undefined
    if (!selection || !sheet || !this.state) return null
    const name = sheet.getSheetName()
    const prefix = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replaceAll("'", "''")}'`
    return {
      target: {
        sheetId: selection.sheetId,
        address: `${formatAddress(selection.range.startRow, selection.range.startColumn)}:${formatAddress(selection.range.endRow, selection.range.endColumn)}`,
      },
      stateId: this.state.stateId,
      revision: this.revision,
      defaultReference: `${prefix}!${selection.address}`,
      targetLabel: `${name}!${formatAddress(selection.range.startRow, selection.range.startColumn)}`,
    }
  }
  /** Read the installed grid projection for presentation verification. */
  projectedMerges() {
    return (
      this.runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheet().getMergeData() ?? []
    )
  }
  projectedDimensions(range: string) {
    const bounds = parseRange(range)
    if (rangeCellCount(bounds) > 2000) throw new Error('presentation_range_limit')
    const sheet = this.runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheet()
    if (!sheet) throw new Error('view_unavailable')
    return {
      rows: Array.from({ length: bounds.endRow - bounds.startRow + 1 }, (_, offset) => ({
        index: bounds.startRow + offset,
        pixels: sheet.getRowHeight(bounds.startRow + offset),
        hidden: !sheet.getRowVisible(bounds.startRow + offset),
      })),
      columns: Array.from({ length: bounds.endColumn - bounds.startColumn + 1 }, (_, offset) => ({
        index: bounds.startColumn + offset,
        pixels: sheet.getColumnWidth(bounds.startColumn + offset),
      })),
    }
  }

  projectedValues(range: string) {
    if (rangeCellCount(parseRange(range)) > 2000) throw new Error('presentation_range_limit')
    return this.runtime.univerAPI
      .getActiveWorkbook()
      ?.getActiveSheet()
      ?.getRange(range)
      .getRawValues()
  }
  projectedStyles(range: string) {
    const bounds = parseRange(range)
    if (rangeCellCount(bounds) > 2000) throw new Error('presentation_range_limit')
    const sheet = this.runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
    if (!sheet) throw new Error('view_unavailable')
    const styles: Record<string, unknown> = {}
    for (let row = bounds.startRow; row <= bounds.endRow; row++)
      for (let column = bounds.startColumn; column <= bounds.endColumn; column++)
        styles[formatAddress(row, column)] = sheet.getRange(row, column).getCellStyleData() ?? {}
    return JSON.parse(JSON.stringify(styles))
  }
  private currentFileVisuals(): WorkbookVisualObject[] {
    const sheetId = this.runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId()
    const sheet = this.state?.sheets.find((s) => s.id === sheetId)
    if (!sheet) return []
    const metadata = this.state?.features?.originalMetadata as
      { visuals?: WorkbookVisualObject[] } | undefined
    const edits = (sheet.features?.visualEdits ?? {}) as Record<
      string,
      Partial<WorkbookVisualObject> & {
        remove?: true
        zIndex?: number
        frameSize?: { width: number; height: number }
      }
    >
    const originals = (metadata?.visuals ?? [])
      .filter((v) => v.sheetId === sheet.originalSheetId && !edits[v.id]?.remove)
      .map((v) => {
        const { remove: _remove, zIndex: _z, frameSize, ...patch } = edits[v.id] ?? {}
        return {
          ...v,
          ...patch,
          ...(patch.fillColor ? { fillGradient: undefined, fillMediaPath: undefined } : {}),
          ...(frameSize ? { frameWidth: frameSize.width, frameHeight: frameSize.height } : {}),
          sheetId: sheet.id,
        }
      })
    return projectVisualOrder(originals, (v) => edits[v.id]?.zIndex)
  }
  get visuals(): readonly WorkbookVisualObject[] {
    const sheetId = this.runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId(),
      sheet = this.state?.sheets.find((s) => s.id === sheetId)
    const chartEdits = (sheet?.features?.chartEdits ?? {}) as Record<string, object>
    return [
      ...this.currentFileVisuals()
        .filter(
          (v) => v.kind === 'shape' || v.kind === 'image' || (v.kind === 'chart' && !!v.chart),
        )
        .map((v) =>
          v.kind === 'chart'
            ? { ...v, chart: applyChartStateEdit(v.chart!, chartEdits[v.chartPath ?? '']) }
            : v,
        ),
      ...((sheet?.features?.visuals ?? sheet?.features?.charts ?? []) as unknown as SheetVisual[]),
    ]
  }
  get charts(): readonly SheetChartVisual[] {
    return this.visuals.filter((v) => v.kind === 'chart' && !!v.chart) as SheetChartVisual[]
  }

  isSessionChart(id: string) {
    return (
      this.state?.sheets.some((sheet) =>
        (
          (sheet.features?.visuals ?? sheet.features?.charts ?? []) as unknown as SheetVisual[]
        ).some((visual) => visual.id === id && visual.kind === 'chart'),
      ) ?? false
    )
  }
  get selectedFeatures(): Record<string, unknown> {
    const sheetId = this.selection?.sheetId
    return {
      ...this.state?.sheets.find((item) => item.id === sheetId)?.features,
      filePageSetup: sheetId ? this.filePageSetup.get(sheetId) : undefined,
    }
  }
  projectedView() {
    const sheet = this.runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
    const config = sheet?.getSheet().getConfig()
    return {
      freeze: sheet?.getFreeze(),
      zoom: sheet?.getZoom(),
      gridlines: config?.showGridlines === 1,
      headings: config?.rowHeader.hidden !== 1,
      rightToLeft: config?.rightToLeft === 1,
      showFormulas: sheet ? this.showFormulasFor(sheet.getSheetId()) : false,
    }
  }
  get selectedNote() {
    const selected = this.selection
    if (!selected) return ''
    const address = formatAddress(selected.range.startRow, selected.range.startColumn)
    const edits = (this.selectedFeatures.noteEdits ?? {}) as Record<string, string | null>
    if (Object.hasOwn(edits, address)) return edits[address] ?? ''
    const original = this.selectedFeatures.originalMetadata as
      { comments?: { row: number; column: number; text: string }[] } | undefined
    return (
      original?.comments?.find(
        (note) =>
          note.row === selected.range.startRow && note.column === selected.range.startColumn,
      )?.text ?? ''
    )
  }
  async insertSummary(functionName: 'SUM' | 'AVERAGE' | 'COUNT' | 'MIN' | 'MAX') {
    const selected = this.selection
    if (!selected || !this.state) throw new Error('missing_selection')
    const row = selected.range.endRow + 1
    if (row >= 1048576) throw new Error('invalid_summary_target')
    const address = formatAddress(row, selected.range.startColumn)
    const result = (await this.host.read(
      selected.sheetId,
      {
        startRow: row,
        endRow: row,
        startColumn: selected.range.startColumn,
        endColumn: selected.range.startColumn,
      },
      this.state.stateId,
    )) as { cells: ViewCell[] }
    if (
      result.cells.some(
        (cell) => cell.address === address && (cell.rawValue !== null || cell.formula),
      )
    )
      throw new Error(
        'summary_target_not_empty: select an empty result cell and enter the formula explicitly',
      )
    return this.enqueue([
      {
        op: 'set_formula',
        sheetId: selected.sheetId,
        address,
        formula: `=${functionName}(${selected.address})`,
      },
    ])
  }
  recalculate() {
    calculateWorkbookNow(this.runtime)
  }
  recalculateSheet() {
    calculateActiveSheet(this.runtime)
  }
  private sortFilterWork: Promise<void> | undefined
  sortFilterCommand(
    command: 'filter-toggle' | 'filter-clear' | 'filter-reapply' | 'sort:asc' | 'sort:desc',
  ) {
    if (this.sortFilterWork)
      return Promise.reject(new Error('A sort/filter operation is already running.'))
    const work = this.runSortFilterCommand(command)
    this.sortFilterWork = work
    const clear = () => {
      if (this.sortFilterWork === work) this.sortFilterWork = undefined
    }
    void work.then(clear, clear)
    return work
  }
  private async runSortFilterCommand(
    command: 'filter-toggle' | 'filter-clear' | 'filter-reapply' | 'sort:asc' | 'sort:desc',
  ) {
    if (!(await this.flushCellEditor()))
      throw new Error('Finish editing before sorting or filtering.')
    const state = this.state,
      sheetId = this.runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId()
    if (!state || !sheetId) throw new Error(t('appWorkbookNotReady'))
    const frozen = { stateId: state.stateId, revision: this.revision }
    const current = () =>
      !this.closed && this.state?.stateId === frozen.stateId && this.revision === frozen.revision
    if (command === 'sort:asc' || command === 'sort:desc') {
      const target = this.ribbonTarget
      if (!target) throw new Error(t('appSelectCellFirst'))
      const bounds = parseRange(target.address)
      if (bounds.startRow === bounds.endRow) throw new Error(t('appSortSelectRows'))
      await this.enqueue(
        [
          {
            op: 'sort_range',
            sheetId: target.sheetId,
            range: target.address,
            byColumn: columnLabel(bounds.startColumn),
            order: command === 'sort:asc' ? 'asc' : 'desc',
            hasHeader: false,
          },
        ],
        frozen,
        { failureInDialog: true },
      )
      return
    }
    const metadata = workbookRangeResultSchema
      .pick({ autoFilter: true, autoFilterColumns: true })
      .strip()
      .parse(
        await this.host.read(
          sheetId,
          { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
          frozen.stateId,
        ),
      )
    if (!current()) throw new Error('The document changed while reading the filter. Try again.')
    const operations: WorkbookOperation[] = []
    if (command === 'filter-toggle') {
      if (metadata.autoFilter) operations.push({ op: 'clear_filter', sheetId })
      else {
        const selection = this.selection
        if (!selection || selection.sheetId !== sheetId) throw new Error(t('appSelectCellFirst'))
        const range = await smartFilterRange(
          selection.range,
          async (range) =>
            (await this.host.read(sheetId, range, frozen.stateId)) as {
              cells: ViewCell[]
              merges?: GridRange[]
            },
          current,
        )
        operations.push({
          op: 'set_filter',
          sheetId,
          range: `${formatAddress(range.startRow, range.startColumn)}:${formatAddress(range.endRow, range.endColumn)}`,
        })
      }
    } else {
      if (!metadata.autoFilter) return
      if (command === 'filter-clear' && !metadata.autoFilterColumns.length) return
      operations.push({
        op: command === 'filter-clear' ? 'clear_filter_criteria' : 'reapply_filter',
        sheetId,
      })
    }
    await this.enqueue(operations, frozen, { failureInDialog: true })
  }
  private filterCommandsInstalled = false
  private filterPopupContext:
    | {
        stateId: string
        revision: number
        sheetId: string
        unitId: string
        col: number
        range: GridRange
      }
    | undefined
  private filterPopupWork: Promise<unknown> | undefined
  private filterPopupRejected = false
  private closeFilterPopup() {
    const commands = this.runtime.univer.__getInjector().get(ICommandService)
    this.internal++
    try {
      if (commands.hasCommand(CloseFilterPanelOperation.id))
        commands.syncExecuteCommand(CloseFilterPanelOperation.id)
      this.filterPopupContext = undefined
      this.filterPopupRejected = false
    } finally {
      this.internal--
    }
  }
  private installFilterPopupCommands() {
    if (this.filterCommandsInstalled) return
    const commands = this.runtime.univer.__getInjector().get(ICommandService)
    if (!commands.hasCommand(OpenFilterPanelOperation.id)) return
    this.filterCommandsInstalled = true
    for (const command of [
      OpenFilterPanelOperation,
      CloseFilterPanelOperation,
      ChangeFilterByOperation,
      SetSheetsFilterCriteriaCommand,
      SortRangeCommand,
    ])
      commands.unregisterCommand(command.id)
    this.disposables.push(
      commands.registerCommand({
        ...OpenFilterPanelOperation,
        handler: async (accessor, raw) => {
          if (this.filterPopupWork) return false
          this.clearError()
          const params = raw as IOpenFilterPanelOperationParams
          const work = this.prepareFilterPopup(params)
          this.filterPopupWork = work
          try {
            const context = await work
            this.filterPopupContext = context
            this.filterPopupRejected = false
            return OpenFilterPanelOperation.handler(accessor, params)
          } catch (error) {
            this.report(error)
            return false
          } finally {
            if (this.filterPopupWork === work) this.filterPopupWork = undefined
          }
        },
      }),
      commands.registerCommand({
        ...ChangeFilterByOperation,
        handler: async (accessor, raw) => {
          if (this.filterPopupWork) return false
          const params = raw as { filterBy: FilterBy }
          if (params.filterBy !== FilterBy.VALUES)
            return ChangeFilterByOperation.handler(accessor, params)
          const context = this.filterPopupContext
          if (!context) return false
          this.clearError()
          const work = this.prepareFilterPopup({
            unitId: context.unitId,
            subUnitId: context.sheetId,
            col: context.col,
          })
          this.filterPopupWork = work
          try {
            this.filterPopupContext = await work
            return ChangeFilterByOperation.handler(accessor, params)
          } catch (error) {
            this.report(error)
            return false
          } finally {
            if (this.filterPopupWork === work) this.filterPopupWork = undefined
          }
        },
      }),
      commands.registerCommand({
        ...CloseFilterPanelOperation,
        handler: (accessor) => {
          if (!this.internal && this.filterPopupWork) return false
          // The stock panel closes after any resolved apply. A host rejection
          // must keep its draft; consume only that automatic close, then allow Cancel.
          if (!this.internal && this.filterPopupRejected) {
            this.filterPopupRejected = false
            return false
          }
          this.filterPopupContext = undefined
          return CloseFilterPanelOperation.handler(accessor, {})
        },
      }),
      commands.registerCommand({
        ...SortRangeCommand,
        handler: async (_accessor, raw) => {
          if (this.filterPopupWork) return false
          this.clearError()
          const params = raw as ISortRangeCommandParams
          const context = this.filterPopupContext
          const work = (async () => {
            if (
              !context ||
              context.stateId !== this.state?.stateId ||
              context.revision !== this.revision ||
              context.unitId !== params.unitId ||
              context.sheetId !== params.subUnitId
            )
              throw new Error('The document changed while the filter was open. Reopen the filter.')
            const [rule] = params.orderRules
            const range = context.range
            if (
              params.hasTitle ||
              params.orderRules.length !== 1 ||
              rule?.colIndex !== context.col ||
              !['asc', 'desc'].includes(rule.type) ||
              params.range.startRow !== range.startRow + 1 ||
              params.range.endRow !== range.endRow ||
              params.range.startColumn !== range.startColumn ||
              params.range.endColumn !== range.endColumn
            )
              throw new Error('Invalid filter popup sort target.')
            await this.enqueue(
              [
                {
                  op: 'sort_range',
                  sheetId: context.sheetId,
                  range: `${formatAddress(range.startRow, range.startColumn)}:${formatAddress(range.endRow, range.endColumn)}`,
                  byColumn: columnLabel(context.col),
                  order: rule.type,
                  hasHeader: true,
                },
              ],
              context,
              { failureInDialog: true },
            )
            return true
          })()
          this.filterPopupWork = work
          try {
            return await work
          } catch (error) {
            this.report(error)
            return false
          } finally {
            if (this.filterPopupWork === work) this.filterPopupWork = undefined
          }
        },
      }),
      commands.registerCommand({
        ...SetSheetsFilterCriteriaCommand,
        handler: async (_accessor, raw) => {
          if (this.filterPopupWork) return false
          this.clearError()
          const params = raw as ISetSheetsFilterCriteriaCommandParams
          const context = this.filterPopupContext
          const work = (async () => {
            if (
              !context ||
              context.stateId !== this.state?.stateId ||
              context.revision !== this.revision ||
              context.unitId !== params.unitId ||
              context.sheetId !== params.subUnitId ||
              context.col !== params.col
            )
              throw new Error('The document changed while the filter was open. Reopen the filter.')
            const operation = filterPopupOperation(
              context.sheetId,
              context.range,
              params.col,
              params.criteria ?? null,
            )
            await this.enqueue([operation], context, { failureInDialog: true })
            return true
          })()
          this.filterPopupWork = work
          try {
            return await work
          } catch (error) {
            this.filterPopupRejected = true
            this.report(error)
            return false
          } finally {
            if (this.filterPopupWork === work) this.filterPopupWork = undefined
          }
        },
      }),
    )
  }
  private async prepareFilterPopup(params: IOpenFilterPanelOperationParams) {
    if (!(await this.flushCellEditor())) throw new Error('Finish editing before filtering.')
    const state = this.state
    const book = this.runtime.univerAPI.getActiveWorkbook()
    const sheet = book?.getActiveSheet()
    if (
      !state ||
      !sheet ||
      params.unitId !== book?.getId() ||
      params.subUnitId !== sheet.getSheetId()
    )
      throw new Error('invalid_manual_target')
    const context = {
      stateId: state.stateId,
      revision: this.revision,
      sheetId: params.subUnitId,
      unitId: params.unitId,
      col: params.col,
    }
    const current = () =>
      !this.closed &&
      this.state?.stateId === context.stateId &&
      this.revision === context.revision &&
      this.runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId() === context.sheetId
    const metadata = workbookRangeResultSchema
      .pick({ autoFilter: true })
      .strip()
      .parse(
        await this.host.read(
          context.sheetId,
          { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
          context.stateId,
        ),
      )
    if (!current()) throw new Error('The document changed while reading the filter. Try again.')
    const range = metadata.autoFilter
    if (
      !range ||
      !Number.isInteger(params.col) ||
      params.col < range.startColumn ||
      params.col > range.endColumn
    )
      throw new Error('filter_column_outside_range')
    // The host reads at most 2000 cells at a time; larger tables load in row bands up to the host's
    // 200,000-cell candidate bound.
    if (rangeCellCount(range) > 200_000)
      throw new Error('Filter popups support tables of at most 200,000 cells.')
    const width = range.endColumn - range.startColumn + 1
    if (width > 2000) throw new Error('Filter popups support tables of at most 2000 columns.')
    const bandRows = Math.max(1, Math.floor(2000 / width))
    const result: { cells: ViewCell[] } = { cells: [] }
    for (let row = range.startRow; row <= range.endRow; row += bandRows) {
      const band = (await this.host.read(
        context.sheetId,
        { ...range, startRow: row, endRow: Math.min(range.endRow, row + bandRows - 1) },
        context.stateId,
      )) as { cells: ViewCell[] }
      result.cells.push(...band.cells)
      if (!current())
        throw new Error('The document changed while loading filter values. Try again.')
    }
    const cells: Record<number, Record<number, ICellData>> = {}
    for (const cell of result.cells) {
      if (cell.formula && cell.calculationState !== 'ready' && cell.kind !== 'error')
        throw new Error('Wait for filter values to finish calculating.')
      const point = parseAddress(cell.address)
      // The popup consumes authoritative formula results rather than recalculating
      // a partially materialized worksheet in the renderer.
      ;(cells[point.row] ??= {})[point.column] = gridCell({
        ...cell,
        rawValue: cell.formula ? (cell.cachedValue ?? cell.rawValue) : cell.rawValue,
      })
    }
    this.internal++
    try {
      sheet
        .getRange(
          range.startRow,
          range.startColumn,
          range.endRow - range.startRow + 1,
          range.endColumn - range.startColumn + 1,
        )
        .setValues(cells)
    } finally {
      this.internal--
    }
    return { ...context, range }
  }
  async advancedFilterContext(createIfMissing = true): Promise<{
    stateId: string
    revision: number
    sheetId: string
    range: GridRange
    columns: { colId: number; label: string }[]
  }> {
    if (!(await this.flushEditor())) throw new Error('Finish editing before filtering.')
    const sheetId = this.runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId(),
      state = this.state
    if (!sheetId || !state) throw new Error(t('appSelectCellFirst'))
    const frozen = { stateId: state.stateId, revision: this.revision }
    const current = () =>
      !this.closed && this.state?.stateId === frozen.stateId && this.revision === frozen.revision
    const anchor = {
      startRow: 0,
      endRow: 0,
      startColumn: 0,
      endColumn: 0,
    }
    const metadata = workbookRangeResultSchema
      .pick({ autoFilter: true, autoFilterColumns: true })
      .strip()
      .parse(await this.host.read(sheetId, anchor, frozen.stateId))
    if (!current()) throw new Error('The document changed while reading the filter. Try again.')
    const range = metadata.autoFilter
    if (!range) {
      if (!createIfMissing) throw new Error(t('appAdvFilterNeedsFilter'))
      await this.sortFilterCommand('filter-toggle')
      return this.advancedFilterContext(false)
    }
    const width = Math.min(range.endColumn - range.startColumn + 1, 26)
    const result = (await this.host.read(
      sheetId,
      { ...range, endRow: range.startRow, endColumn: range.startColumn + width - 1 },
      frozen.stateId,
    )) as { cells: ViewCell[] }
    if (!current()) throw new Error('The document changed while reading filter headers. Try again.')
    const cells = new Map(result.cells.map((cell) => [cell.address, cell]))
    const columns = Array.from({ length: width }, (_, colId) => {
      const cell = cells.get(formatAddress(range.startRow, range.startColumn + colId))
      if (cell?.formula && cell.calculationState !== 'ready')
        throw new Error('Wait for filter headers to finish calculating.')
      const value = cell?.formula ? cell.cachedValue : cell?.rawValue
      return {
        colId,
        label:
          value === null || value === undefined || value === ''
            ? t('appColumnLabel', { col: columnLabel(range.startColumn + colId) })
            : String(value),
      }
    })
    return { ...frozen, sheetId: sheetId, range, columns }
  }
  private advancedFilterWork: Promise<void> | undefined
  advancedFilter(
    criteria: AdvancedFilterCriteria,
    context: Awaited<ReturnType<WorkbookRuntime['advancedFilterContext']>>,
  ) {
    if (this.advancedFilterWork)
      return Promise.reject(new Error('An advanced filter is already being applied.'))
    const work = this.runAdvancedFilter(criteria, context)
    this.advancedFilterWork = work
    const clear = () => {
      if (this.advancedFilterWork === work) this.advancedFilterWork = undefined
    }
    void work.then(clear, clear)
    return work
  }
  private async runAdvancedFilter(
    criteria: AdvancedFilterCriteria,
    context: Awaited<ReturnType<WorkbookRuntime['advancedFilterContext']>>,
  ) {
    if (!(await this.flushCellEditor())) throw new Error('Finish editing before filtering.')
    if (
      this.closed ||
      this.state?.stateId !== context.stateId ||
      this.revision !== context.revision
    )
      throw new Error('The document changed while Advanced Filter was open. Reopen the dialog.')
    if (
      !Number.isInteger(criteria.colId) ||
      !context.columns.some((column) => column.colId === criteria.colId)
    )
      throw new Error('Invalid filter column.')
    await this.enqueue(
      [
        {
          op: 'set_filter_criteria',
          sheetId: context.sheetId,
          column: columnLabel(context.range.startColumn + criteria.colId),
          customs: { and: criteria.and, filters: [...criteria.filters] },
        },
      ],
      context,
      { failureInDialog: true },
    )
  }
  async removeDuplicatesContext() {
    if (!(await this.flushEditor())) throw new Error('Finish editing before removing duplicates.')
    const target = this.ribbonTarget,
      state = this.state
    if (!target || !state) throw new Error(t('appSelectCellFirst'))
    const bounds = parseRange(target.address)
    if (bounds.endRow <= bounds.startRow) throw new Error(t('appDedupeSelectRows'))
    if (rangeCellCount(bounds) > 200_000)
      throw new Error('Remove Duplicates supports selections of at most 200,000 cells.')
    return {
      target: { sheetId: target.sheetId, address: target.address },
      stateId: state.stateId,
      revision: this.revision,
    }
  }
  private removeDuplicatesWork: Promise<number> | undefined
  removeDuplicates(
    hasHeader: boolean,
    context: Awaited<ReturnType<WorkbookRuntime['removeDuplicatesContext']>>,
  ) {
    if (this.removeDuplicatesWork)
      return Promise.reject(new Error('Remove Duplicates is already running.'))
    const work = this.runRemoveDuplicates(hasHeader, context)
    this.removeDuplicatesWork = work
    const clear = () => {
      if (this.removeDuplicatesWork === work) this.removeDuplicatesWork = undefined
    }
    void work.then(clear, clear)
    return work
  }
  private async runRemoveDuplicates(
    hasHeader: boolean,
    context: Awaited<ReturnType<WorkbookRuntime['removeDuplicatesContext']>>,
  ) {
    if (!(await this.flushCellEditor()))
      throw new Error('Finish editing before using this command.')
    const target = context.target,
      state = this.state
    if (
      this.closed ||
      !state ||
      state.stateId !== context.stateId ||
      this.revision !== context.revision
    )
      throw new Error('The document changed while Remove Duplicates was open. Reopen the dialog.')
    const bounds = parseRange(target.address)
    if (bounds.endRow <= bounds.startRow) throw new Error(t('appDedupeSelectRows'))
    if (rangeCellCount(bounds) > 200_000)
      throw new Error('Remove Duplicates supports selections of at most 200,000 cells.')
    const frozen = { stateId: context.stateId, revision: context.revision }
    const current = () =>
      !this.closed && this.state?.stateId === frozen.stateId && this.revision === frozen.revision
    // The host reads at most 2000 cells at a time: larger selections load in row bands.
    const result: { cells: ViewCell[]; merges: GridRange[] } = { cells: [], merges: [] }
    const bandRows = Math.max(1, Math.floor(2000 / (bounds.endColumn - bounds.startColumn + 1)))
    for (let row = bounds.startRow; row <= bounds.endRow; row += bandRows) {
      const band = (await this.host.read(
        target.sheetId,
        { ...bounds, startRow: row, endRow: Math.min(bounds.endRow, row + bandRows - 1) },
        frozen.stateId,
      )) as { cells: ViewCell[]; merges?: GridRange[] }
      result.cells.push(...band.cells)
      result.merges.push(...(band.merges ?? []))
      if (!current())
        throw new Error('The document changed while reading the selection. Try again.')
    }
    if (
      result.merges.some(
        (merge) =>
          merge.startRow <= bounds.endRow &&
          merge.endRow >= bounds.startRow &&
          merge.startColumn <= bounds.endColumn &&
          merge.endColumn >= bounds.startColumn,
      )
    )
      throw new Error('Remove merged cells from the selection before removing duplicates.')
    const cells = new Map(result.cells.map((cell) => [cell.address, cell]))
    const original: DedupeValue[][] = []
    for (let row = bounds.startRow; row <= bounds.endRow; row++) {
      const values: DedupeValue[] = []
      for (let column = bounds.startColumn; column <= bounds.endColumn; column++) {
        const cell = cells.get(formatAddress(row, column))
        if (!cell)
          throw new Error('Selection data is incomplete. Wait for loading to finish and retry.')
        if (cell.kind === 'error')
          throw new Error('Remove Duplicates cannot process cells containing formula errors.')
        if (cell.formula && cell.calculationState !== 'ready')
          throw new Error('Wait for selected formulas to finish calculating and retry.')
        if (cell.formula && !Object.hasOwn(cell, 'cachedValue'))
          throw new Error('A selected formula has no current result. Recalculate and retry.')
        const value = cell.formula ? cell.cachedValue : cell.rawValue
        if (
          value !== null &&
          typeof value !== 'string' &&
          typeof value !== 'number' &&
          typeof value !== 'boolean'
        )
          throw new Error(
            'The selection contains a value type that Remove Duplicates cannot process.',
          )
        values.push(value ?? null)
      }
      original.push(values)
    }
    const deduped = dedupeRows(original, hasHeader)
    if (deduped.removed === 0) return 0
    const next = [...deduped.rows]
    while (next.length < original.length) next.push(Array(original[0]!.length).fill(null))
    const operations: WorkbookOperation[] = []
    for (let row = 0; row < original.length; row++) {
      if (original[row]!.every((value, column) => Object.is(value, next[row]![column]))) continue
      for (let column = 0; column < original[row]!.length; column++)
        operations.push({
          op: 'set_cell',
          sheetId: target.sheetId,
          address: formatAddress(bounds.startRow + row, bounds.startColumn + column),
          value: next[row]![column]!,
        })
    }
    if (operations.length > 2000) {
      // One staged commit (a single history entry) carries the writes beyond the batch ceiling. Runs of
      // changed rows are written as value blocks, which the host applies far faster than cell by cell.
      const width = original[0]!.length
      const band = Math.max(1, Math.floor(1000 / width))
      const stages: WorkbookOperation[][] = []
      let row = 0
      while (row < original.length) {
        if (original[row]!.every((value, column) => Object.is(value, next[row]![column]))) {
          row++
          continue
        }
        const first = row
        while (
          row < original.length &&
          row - first < band &&
          !original[row]!.every((value, column) => Object.is(value, next[row]![column]))
        )
          row++
        stages.push([
          {
            op: 'set_range',
            sheetId: target.sheetId,
            start: formatAddress(bounds.startRow + first, bounds.startColumn),
            values: next.slice(first, row),
          },
        ])
      }
      await this.enqueueStaged(stages, frozen)
    } else await this.enqueue(operations, frozen, { failureInDialog: true })
    return deduped.removed
  }
  async flashFill() {
    if (!(await this.flushCellEditor())) throw new Error('Finish editing before using Flash Fill.')
    const target = this.ribbonTarget,
      state = this.state
    if (!target || !state) throw new Error(t('appSelectCellFirst'))
    const selected = parseRange(target.address),
      column = selected.startColumn,
      sheet = state.sheets.find((item) => item.id === target.sheetId)
    if (!sheet) throw new Error('The selected sheet no longer exists.')
    if (column === 0) throw new Error('Flash Fill needs at least one source column to the left.')
    const frozen = { stateId: state.stateId, revision: this.revision }
    const current = () =>
      !this.closed && this.state?.stateId === frozen.stateId && this.revision === frozen.revision
    let endRow = selected.endRow
    if (selected.startRow === selected.endRow) {
      const probeEnd = Math.min(selected.startRow + 999, Math.max(0, sheet.rows - 1))
      const probe = (await this.host.read(
        target.sheetId,
        {
          startRow: selected.startRow,
          endRow: probeEnd,
          startColumn: column - 1,
          endColumn: column - 1,
        },
        frozen.stateId,
      )) as { cells: ViewCell[] }
      if (!current())
        throw new Error('The document changed while reading Flash Fill data. Try again.')
      const cells = new Map(probe.cells.map((cell) => [cell.address, cell]))
      endRow = selected.startRow - 1
      for (let row = selected.startRow; row <= probeEnd; row++) {
        const cell = cells.get(formatAddress(row, column - 1))
        if (!cell) throw new Error('Flash Fill source data is incomplete. Wait and retry.')
        if (cell.kind === 'error' || (cell.formula && cell.calculationState !== 'ready'))
          throw new Error('Flash Fill source formulas must finish calculating first.')
        if (cell.formula && !Object.hasOwn(cell, 'cachedValue'))
          throw new Error('A Flash Fill formula has no current result. Recalculate and retry.')
        const value = cell.formula ? cell.cachedValue : cell.rawValue
        if (value === null || value === undefined || String(value) === '') break
        endRow = row
      }
    }
    const rowCount = endRow - selected.startRow + 1
    if (rowCount < 2) throw new Error('Flash Fill needs at least two rows with source data.')
    if (rowCount > 2000) throw new Error('Flash Fill supports at most 2000 rows at a time.')
    const firstSource = Math.max(0, column - 6),
      width = column - firstSource + 1,
      valuesByAddress = new Map<string, ViewCell>(),
      chunkRows = Math.max(1, Math.floor(2000 / width))
    for (let startRow = selected.startRow; startRow <= endRow; startRow += chunkRows) {
      const chunkEnd = Math.min(endRow, startRow + chunkRows - 1)
      const read = (await this.host.read(
        target.sheetId,
        { startRow, endRow: chunkEnd, startColumn: firstSource, endColumn: column },
        frozen.stateId,
      )) as { cells: ViewCell[]; merges?: GridRange[] }
      if (!current())
        throw new Error('The document changed while reading Flash Fill data. Try again.')
      if (
        read.merges?.some(
          (merge) =>
            merge.startRow <= endRow &&
            merge.endRow >= selected.startRow &&
            merge.startColumn <= column &&
            merge.endColumn >= firstSource,
        )
      )
        throw new Error('Remove merged cells from the Flash Fill range before continuing.')
      for (const cell of read.cells) valuesByAddress.set(cell.address, cell)
    }
    const matrix: { source: string[]; target: string }[] = []
    for (let row = selected.startRow; row <= endRow; row++) {
      const cells: ViewCell[] = []
      for (let col = firstSource; col <= column; col++) {
        const cell = valuesByAddress.get(formatAddress(row, col))
        if (!cell) throw new Error('Flash Fill source data is incomplete. Wait and retry.')
        if (cell.kind === 'error') throw new Error('Flash Fill cannot use cells containing errors.')
        if (cell.formula && cell.calculationState !== 'ready')
          throw new Error('Flash Fill source formulas must finish calculating first.')
        if (cell.formula && !Object.hasOwn(cell, 'cachedValue'))
          throw new Error('A Flash Fill formula has no current result. Recalculate and retry.')
        cells.push(cell)
      }
      const sourceCells = cells.slice(0, -1),
        targetCell = cells[cells.length - 1]!,
        source = sourceCells.map((cell) =>
          String((cell.formula ? cell.cachedValue : cell.rawValue) ?? ''),
        ),
        raw = targetCell.formula ? targetCell.cachedValue : targetCell.rawValue
      matrix.push({ source, target: String(raw ?? '') })
    }
    const examples: FlashFillExample[] = []
    for (const item of matrix) {
      if (item.target !== '') {
        examples.push({ source: item.source, output: item.target })
        if (examples.length === 3) break
      }
    }
    if (!examples.length)
      throw new Error('Enter one or more Flash Fill examples in the target column.')
    const template = inferFlashFillTemplate(examples)
    if (!template) throw new Error('The examples do not show a Flash Fill pattern I can apply.')
    const operations: WorkbookOperation[] = []
    for (const [offset, entry] of matrix.entries()) {
      if (entry.target !== '' || entry.source.every((value) => value.trim() === '')) continue
      const value = applyFlashFillTemplate(template, entry.source)
      if (value === entry.target) continue
      operations.push({
        op: 'set_cell',
        sheetId: target.sheetId,
        address: formatAddress(selected.startRow + offset, column),
        value,
      })
    }
    if (!operations.length) throw new Error('There are no blank target cells to fill.')
    await this.enqueue(operations, frozen)
    return operations.length
  }
  private textToColumnsWork: Promise<number> | undefined
  textToColumns(delimiter: 1 | 2 | 4 | 8) {
    if (this.textToColumnsWork)
      return Promise.reject(new Error('Text to Columns is already being applied.'))
    const work = this.runTextToColumns(delimiter)
    this.textToColumnsWork = work
    const clear = () => {
      if (this.textToColumnsWork === work) this.textToColumnsWork = undefined
    }
    void work.then(clear, clear)
    return work
  }
  private async runTextToColumns(delimiter: 1 | 2 | 4 | 8) {
    if (![1, 2, 4, 8].includes(delimiter)) throw new Error('Invalid Text to Columns delimiter.')
    if (!(await this.flushCellEditor())) throw new Error('Finish editing before splitting text.')
    const target = this.ribbonTarget,
      state = this.state
    if (!target || !state) throw new Error(t('appSelectCellFirst'))
    const bounds = parseRange(target.address)
    if (bounds.startColumn !== bounds.endColumn)
      throw new Error('Text to Columns requires a selection in one column.')
    if (rangeCellCount(bounds) > 100_000)
      throw new Error('Text to Columns supports at most 100,000 selected cells.')
    const sheet = state.sheets.find((item) => item.id === target.sheetId)
    if (!sheet) throw new Error('The selected sheet no longer exists.')
    const frozen = { stateId: state.stateId, revision: this.revision }
    const current = () =>
      !this.closed && this.state?.stateId === frozen.stateId && this.revision === frozen.revision
    // The host reads at most 2000 cells at a time: longer selections load in row bands.
    const readBanded = async (range: GridRange) => {
      const cells: ViewCell[] = [],
        merges: GridRange[] = []
      const bandRows = Math.max(1, Math.floor(2000 / (range.endColumn - range.startColumn + 1)))
      for (let row = range.startRow; row <= range.endRow; row += bandRows) {
        const band = (await this.host.read(
          target.sheetId,
          { ...range, startRow: row, endRow: Math.min(range.endRow, row + bandRows - 1) },
          frozen.stateId,
        )) as { cells: ViewCell[]; merges?: GridRange[] }
        cells.push(...band.cells)
        merges.push(...(band.merges ?? []))
        if (!current())
          throw new Error('The document changed while reading text to split. Try again.')
      }
      return { cells, merges }
    }
    const source = await readBanded(bounds)
    const sourceCells = new Map(source.cells.map((cell) => [cell.address, cell]))
    const delimiterText = { 1: '\t', 2: ',', 4: ';', 8: ' ' }[delimiter],
      separator = new RegExp(`[${delimiterText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}]`),
      rows: (string[] | undefined)[] = []
    let lastRow = -1,
      maxLength = 0
    for (let row = bounds.startRow; row <= bounds.endRow; row++) {
      const cell = sourceCells.get(formatAddress(row, bounds.startColumn))
      if (cell?.kind === 'error')
        throw new Error('Text to Columns cannot split cells containing errors.')
      if (cell?.formula && cell.calculationState !== 'ready')
        throw new Error('Wait for source formulas to finish calculating, then retry.')
      if (cell?.formula && !Object.hasOwn(cell, 'cachedValue'))
        throw new Error('A source formula has no current result. Recalculate and retry.')
      const value = cell?.formula ? cell.cachedValue : cell?.rawValue
      const parts = typeof value === 'string' ? value.split(separator) : undefined
      rows.push(parts)
      if (parts) {
        lastRow = row
        maxLength = Math.max(maxLength, parts.length)
      }
    }
    if (lastRow < bounds.startRow || maxLength === 0)
      throw new Error('The selection contains no text values to split.')
    const endColumn = bounds.startColumn + maxLength,
      outputBounds = {
        startRow: bounds.startRow,
        endRow: lastRow,
        startColumn: bounds.startColumn,
        endColumn,
      }
    if (endColumn >= MAX_GRID_COLUMNS)
      throw new Error('Text to Columns would exceed the worksheet column limit.')
    if (rangeCellCount(outputBounds) > 200_000)
      throw new Error('Text to Columns would write more than 200,000 cells at once.')
    const output = await readBanded(outputBounds)
    if (output.merges.length)
      throw new Error(
        'Remove merged cells from the Text to Columns output range before continuing.',
      )
    const existing = new Map(output.cells.map((cell) => [cell.address, cell]))
    const operations: WorkbookOperation[] = []
    for (let row = bounds.startRow; row <= lastRow; row++) {
      const parts = rows[row - bounds.startRow]
      for (let offset = 0; offset <= maxLength; offset++) {
        if (offset === 0 && parts?.length === 1) continue
        const value = parts?.[offset] || null,
          address = formatAddress(row, bounds.startColumn + offset),
          cell = existing.get(address),
          currentValue = cell?.formula ? cell.cachedValue : cell?.rawValue
        if (cell?.formula || !Object.is(currentValue ?? null, value))
          operations.push({ op: 'set_cell', sheetId: target.sheetId, address, value })
      }
    }
    if (operations.length > 2000) {
      // One staged commit carries the writes beyond the batch ceiling, as value blocks of changed rows.
      const stages: WorkbookOperation[][] = []
      const width = maxLength + 1
      const band = Math.max(1, Math.floor(1000 / width))
      const changedRows = new Set(
        operations.map((op) => parseAddress((op as { address: string }).address).row),
      )
      let row = bounds.startRow
      while (row <= lastRow) {
        if (!changedRows.has(row)) {
          row++
          continue
        }
        // A single-part row keeps its source cell: its block starts one column to the right.
        const skipsFirst = rows[row - bounds.startRow]?.length === 1
        const first = row
        while (
          row <= lastRow &&
          row - first < band &&
          changedRows.has(row) &&
          (rows[row - bounds.startRow]?.length === 1) === skipsFirst
        )
          row++
        const offset = skipsFirst ? 1 : 0
        stages.push([
          {
            op: 'set_range',
            sheetId: target.sheetId,
            start: formatAddress(first, bounds.startColumn + offset),
            values: rows
              .slice(first - bounds.startRow, row - bounds.startRow)
              .map((parts) =>
                Array.from(
                  { length: width - offset },
                  (_, index) => parts?.[index + offset] || null,
                ),
              ),
          },
        ])
      }
      await this.enqueueStaged(stages, frozen)
    } else if (operations.length) await this.enqueue(operations, frozen, { failureInDialog: true })
    return operations.length
  }
  private csvImportWork: Promise<unknown> | undefined
  /** Data → From Text/CSV at the selection's top-left cell, as one canonical transaction. */
  importCsv(lang: string) {
    if (this.csvImportWork) return Promise.reject(new Error('A CSV import is already running.'))
    const work = this.runImportCsv(lang)
    this.csvImportWork = work
    const clear = () => {
      if (this.csvImportWork === work) this.csvImportWork = undefined
    }
    void work.then(clear, clear)
    return work
  }
  private async runImportCsv(lang: string) {
    if (!this.ribbonTarget) throw new Error(t('appSelectCellFirst'))
    const file = await pickCsvFile()
    if (!file) return null
    if (file.size > CSV_IMPORT_MAX_BYTES) throw new Error(t('appCsvTooLarge'))
    let text: string
    try {
      text = decodeCsvBuffer(new Uint8Array(await file.arrayBuffer()), CSV_CHARSET_BY_LANG[lang])
    } catch (error) {
      throw new Error(t('appCsvTooLarge'), { cause: error })
    }
    const matrix = csvImportMatrix(text)
    if ('error' in matrix) throw new Error(t(matrix.error))
    const { rows, columns, values } = matrix
    if (!(await this.flushCellEditor())) throw new Error('Finish editing before importing a CSV.')
    // The pick can take arbitrarily long; bind to the selection and revision as of the write.
    const target = this.ribbonTarget,
      state = this.state
    if (!target || !state || this.closed) throw new Error(t('appSelectCellFirst'))
    const frozen = { stateId: state.stateId, revision: this.revision }
    const origin = parseAddress(target.address.split(':')[0]!)
    if (origin.row + rows > MAX_GRID_ROWS || origin.column + columns > MAX_GRID_COLUMNS)
      throw new Error('The CSV data would extend past the worksheet grid.')
    const destination = {
      startRow: origin.row,
      endRow: origin.row + rows - 1,
      startColumn: origin.column,
      endColumn: origin.column + columns - 1,
    }
    const page = (await this.host.read(
      target.sheetId,
      { ...destination, endRow: origin.row, endColumn: origin.column },
      frozen.stateId,
    )) as { merges?: GridRange[] }
    if (this.closed || this.state?.stateId !== frozen.stateId || this.revision !== frozen.revision)
      throw new Error('The document changed while importing the CSV. Try again.')
    if (
      page.merges?.some(
        (merge) =>
          merge.startRow <= destination.endRow &&
          merge.endRow >= destination.startRow &&
          merge.startColumn <= destination.endColumn &&
          merge.endColumn >= destination.startColumn,
      )
    )
      throw new Error('Remove merged cells from the CSV destination range before importing.')
    // set_range carries at most 500 rows × 100 columns; each stage stays within one ordinary batch.
    const stages: WorkbookOperation[][] = []
    let stage: WorkbookOperation[] = [],
      budget = 0
    for (let column = 0; column < columns; column += 100) {
      const width = Math.min(100, columns - column),
        band = Math.max(1, Math.min(500, Math.floor(2000 / width)))
      for (let row = 0; row < rows; row += band) {
        const block = values
          .slice(row, row + band)
          .map((line) =>
            line.slice(column, column + width).map((value) => (value === '' ? null : value)),
          )
        const cells = block.length * width
        if (stage.length && budget + cells > 2000) {
          stages.push(stage)
          stage = []
          budget = 0
        }
        stage.push({
          op: 'set_range',
          sheetId: target.sheetId,
          start: formatAddress(origin.row + row, origin.column + column),
          values: block,
        })
        budget += cells
      }
    }
    stages.push(stage)
    if (stages.length === 1) await this.enqueue(stages[0]!, frozen)
    else await this.enqueueStaged(stages, frozen)
    return { rows, columns, cell: formatAddress(origin.row, origin.column) }
  }
  /** Reads a PivotTable source through the host: field captions, numeric flags and label grid. */
  private async readPivotSource(
    sheetId: string,
    range: string,
    frozen: { stateId: string; revision: number },
  ) {
    const state = this.state
    if (!state) throw new Error(t('appOpenXlsxFirst'))
    const bounds = parseRange(range)
    const rows = bounds.endRow - bounds.startRow + 1,
      columns = Math.min(PIVOT_SOURCE_COL_LIMIT, bounds.endColumn - bounds.startColumn + 1)
    if (rows > PIVOT_SOURCE_ROW_LIMIT) throw new Error(t('appPivotSourceRowLimit'))
    const cells = new Map<string, ViewCell>()
    const step = Math.max(1, Math.floor(2000 / columns))
    for (let row = bounds.startRow; row <= bounds.endRow; row += step) {
      const page = (await this.host.read(
        sheetId,
        {
          startRow: row,
          endRow: Math.min(bounds.endRow, row + step - 1),
          startColumn: bounds.startColumn,
          endColumn: bounds.startColumn + columns - 1,
        },
        frozen.stateId,
      )) as { cells: ViewCell[] }
      if (
        this.closed ||
        this.state?.stateId !== frozen.stateId ||
        this.revision !== frozen.revision
      )
        throw new Error('The document changed while reading the PivotTable source. Try again.')
      for (const cell of page.cells) cells.set(cell.address, cell)
    }
    const sourceCell = (row: number, column: number) => {
      const cell = cells.get(formatAddress(row, column))
      return {
        value: ((cell?.formula ? cell.cachedValue : cell?.rawValue) ?? null) as
          string | number | boolean | null,
        pattern: cell?.numberFormat ?? cell?.style?.numberFormat,
        static1904: state.dateSystem === '1904' && !cell?.formula,
      }
    }
    const raw = Array.from({ length: rows }, (_, r) =>
      Array.from({ length: columns }, (_, c) =>
        sourceCell(bounds.startRow + r, bounds.startColumn + c),
      ),
    )
    const grid = pivotSourceGrid(raw)
    const fields: PivotField[] = grid[0]!.map((value, offset) => ({
      label:
        value === null || value === ''
          ? t('appColumnLabel', { col: columnLabel(bounds.startColumn + offset) })
          : String(value).trim(),
      colIndex: bounds.startColumn + offset,
      numeric: isNumericPivotColumn(raw.slice(1).map((line) => line[offset]!)),
    }))
    return { fields, grid }
  }
  /** Insert/Data → PivotTable: source fields of the selection, frozen for the dialog. */
  async pivotCreateContext() {
    if (!(await this.flushEditor())) throw new Error('Finish editing before creating a PivotTable.')
    const target = this.ribbonTarget,
      state = this.state
    if (!target || !state) throw new Error(t('appSelectCellFirst'))
    const bounds = parseRange(target.address)
    const frozen = { stateId: state.stateId, revision: this.revision }
    const context = { sheetId: target.sheetId, sourceRange: target.address, frozen }
    if (bounds.endRow - bounds.startRow < 1) return { ...context, fields: [] as PivotField[] }
    const { fields } = await this.readPivotSource(target.sheetId, target.address, frozen)
    return { ...context, fields }
  }
  /** The PivotTable under the selection (session or file), for the Fields pane. */
  private pivotAtSelection() {
    const selected = this.selection,
      state = this.state
    if (!selected || !state) return null
    const sheet = state.sheets.find((item) => item.id === selected.sheetId)
    if (!sheet) return null
    const { startRow: row, startColumn: column } = selected.range
    const inside = (area: GridRange, pageRows: number) =>
      row >= area.startRow - pageRows &&
      row <= area.endRow &&
      column >= area.startColumn &&
      column <= area.endColumn
    for (const add of (sheet.features?.pivotAdditions ?? []) as WorkbookPivotAdd[])
      if (inside(add.location, pivotPageRows(add.pageFieldIndices?.length ?? 0)))
        return { sheet, key: `session:${add.name}`, session: add }
    const relayouts = (sheet.features?.pivotRelayouts ?? {}) as Record<string, WorkbookPivotAdd>
    for (const pivot of filePivotsOfSheet(sheet)) {
      const relayout = relayouts[pivot.path]
      const area = relayout
        ? relayout.location
        : pivot.definition
          ? parseRange(pivot.definition.outputRef)
          : null
      const pageRows = relayout
        ? pivotPageRows(relayout.pageFieldIndices?.length ?? 0)
        : pivot.definition
          ? pivotPageRows(pivot.definition.pageFields.length)
          : 0
      if (area && inside(area, pageRows)) return { sheet, key: pivot.path, file: pivot, relayout }
    }
    return null
  }
  /** Identity of the pivot under the cursor; the Fields pane follows it like the desktop. */
  get pivotSelectionKey() {
    return this.pivotAtSelection()?.key ?? null
  }
  /** Edit PivotTable: the Fields pane seed for the pivot under the cursor. */
  async pivotEditSeed() {
    if (!(await this.flushEditor()))
      throw new Error('Finish editing before editing the PivotTable.')
    const state = this.state
    const hit = this.pivotAtSelection()
    if (!state || !hit) throw new Error(t('appPutCursorInPivot'))
    const frozen = { stateId: state.stateId, revision: this.revision }
    const add = hit.session ?? hit.relayout
    let model: PivotLayoutModel, sourceSheetId: string, sourceRange: string, fieldNames: string[]
    if (add) {
      model = modelOfPivotAdd(add)
      sourceSheetId = add.sourceSheetId
      sourceRange = `${formatAddress(add.sourceArea.startRow, add.sourceArea.startColumn)}:${formatAddress(add.sourceArea.endRow, add.sourceArea.endColumn)}`
      fieldNames = [...add.fieldNames]
    } else {
      const definition = hit.file?.definition
      if (!definition) throw new Error(t('appPivotDefNotLoadedSave'))
      const parsed = modelOfDefinition(definition)
      if (typeof parsed === 'string') throw new Error(parsed)
      const source = pivotSourceSheet(state, definition.sourceSheet)
      if (!source)
        throw new Error(t('appPivotSourceSheetNotFound', { name: definition.sourceSheet }))
      model = parsed.model
      sourceSheetId = source.id
      sourceRange = definition.sourceRef
      fieldNames = definition.fields.slice(0, parsed.sourceFieldCount).map((field) => field.name)
    }
    const read = await this.readPivotSource(sourceSheetId, sourceRange, frozen)
    const fields: PivotField[] = fieldNames.map((label, index) => ({
      label,
      colIndex: parseRange(sourceRange).startColumn + index,
      numeric: read.fields[index]?.numeric ?? false,
    }))
    return {
      key: hit.key,
      sheetId: hit.sheet.id,
      pivot: hit.session ? ({ session: hit.session.name } as const) : ({ file: hit.key } as const),
      fields,
      sourceRange,
      targetCell: sourceRange.split(':')[0]!,
      model,
      members: (fieldIndex: number) => {
        const members: string[] = []
        for (const line of read.grid.slice(1)) {
          const label = String(line[fieldIndex] ?? '')
          if (!members.includes(label)) members.push(label)
        }
        return members
      },
    }
  }
  /** Insert → Slicer / Timeline: the file PivotTable under the cursor, as the desktop requires. */
  filePivotAtSelection() {
    const hit = this.pivotAtSelection()
    if (!hit) throw new Error(t('appCursorNotInPivot'))
    if (hit.session || !hit.file?.definition) throw new Error(t('appPivotDefNotLoadedSave'))
    return { sheetId: hit.sheet.id, pivotPath: hit.key, definition: hit.file.definition }
  }
  /** The current (session-filtered) definition of a file PivotTable. */
  filePivotDefinition(sheetId: string, pivotPath: string) {
    const sheet = this.state?.sheets.find((item) => item.id === sheetId)
    return sheet && filePivotsOfSheet(sheet).find((pivot) => pivot.path === pivotPath)?.definition
  }
  /** Slicer / timeline selection (null = all members); returns a message or null. */
  async filterPivot(
    target: { readonly sheetId: string; readonly pivotPath: string; readonly field: number },
    selected: readonly number[] | null,
  ) {
    if (!(await this.flushEditor())) return 'Finish editing before filtering the PivotTable.'
    try {
      await this.enqueue(
        [
          {
            op: 'filter_pivot',
            sheetId: target.sheetId,
            pivot: target.pivotPath,
            field: target.field,
            selected: selected === null ? null : [...selected],
            subtotalLabel: t('appPivotSubtotal'),
          },
        ],
        undefined,
        { failureInDialog: true },
      )
    } catch (error) {
      return pivotErrorMessage(error)
    }
    return null
  }
  /** Fields pane apply: a new layout for the seeded pivot; returns a pane message or null. */
  async editPivot(
    seed: Awaited<ReturnType<WorkbookRuntime['pivotEditSeed']>>,
    model: PivotLayoutModel,
  ) {
    const parts = pivotConfigToOpParts(
      modelToConfig(model, seed.sourceRange, seed.targetCell),
      seed.fields,
    )
    if (typeof parts === 'string') return parts
    try {
      await this.enqueue(
        [
          {
            op: 'edit_pivot',
            sheetId: seed.sheetId,
            pivot: seed.pivot,
            ...parts,
            subtotalLabel: t('appPivotSubtotal'),
          },
        ],
        undefined,
        { failureInDialog: true },
      )
    } catch (error) {
      return pivotErrorMessage(error)
    }
    return null
  }
  /** Create PivotTable with Excel's default placement; returns a dialog message or null. */
  async createPivot(
    request: PivotCreateRequest,
    context: Awaited<ReturnType<WorkbookRuntime['pivotCreateContext']>>,
  ) {
    const model = defaultPivotModel(context.fields)
    if (validatePivotModel(model) !== null) return t('dlgPivotErrNeedValue')
    const parts = pivotConfigToOpParts(
      modelToConfig(model, request.sourceRange, request.targetCell),
      context.fields,
    )
    if (typeof parts === 'string') return parts
    try {
      await this.enqueue(
        [
          {
            op: 'add_pivot',
            sheetId: context.sheetId,
            sourceRange: request.sourceRange,
            targetCell: request.targetCell || 'A1',
            ...parts,
            subtotalLabel: t('appPivotSubtotal'),
          },
        ],
        context.frozen,
        { failureInDialog: true },
      )
    } catch (error) {
      return pivotErrorMessage(error)
    }
    return null
  }
  private pivotRefreshWork: Promise<number> | undefined
  /**
   * Data → Refresh (active sheet) / Refresh All (workbook): recomputes file
   * PivotTables from fresh source values, as the desktop does. Pivots created
   * this session are written on save and refresh after that.
   */
  refreshPivots(scope: 'sheet' | 'all') {
    if (this.pivotRefreshWork)
      return Promise.reject(new Error('PivotTables are already being refreshed.'))
    const work = (async () => {
      if (!(await this.flushCellEditor())) throw new Error('Finish editing before refreshing.')
      const state = this.state,
        target = this.ribbonTarget
      if (!state) throw new Error(t('appOpenXlsxFirst'))
      if (scope === 'sheet' && !target) throw new Error(t('appNoActiveSheet'))
      const sheets = state.sheets.filter((sheet) =>
        scope === 'sheet' ? sheet.id === target!.sheetId : true,
      )
      const withFilePivots = sheets.filter((sheet) => filePivotsOfSheet(sheet).length > 0)
      if (!withFilePivots.length) {
        const pending = sheets.some(
          (sheet) => ((sheet.features?.pivotAdditions ?? []) as unknown[]).length > 0,
        )
        throw new Error(
          t(
            pending
              ? 'appPivotDefNotLoadedSave'
              : scope === 'sheet'
                ? 'appCurrentSheetNoPivot'
                : 'appWorkbookNoPivot',
          ),
        )
      }
      const count = withFilePivots.reduce((sum, sheet) => sum + filePivotsOfSheet(sheet).length, 0)
      try {
        await this.enqueue(
          withFilePivots.map((sheet) => ({
            op: 'refresh_pivot' as const,
            sheetId: sheet.id,
            subtotalLabel: t('appPivotSubtotal'),
          })),
          { stateId: state.stateId, revision: this.revision },
          { failureInDialog: true },
        )
      } catch (error) {
        throw new Error(pivotErrorMessage(error), { cause: error })
      }
      return count
    })()
    this.pivotRefreshWork = work
    const clear = () => {
      if (this.pivotRefreshWork === work) this.pivotRefreshWork = undefined
    }
    void work.then(clear, clear)
    return work
  }
  async subtotalContext() {
    if (!(await this.flushEditor())) throw new Error('Finish editing before creating subtotals.')
    const target = this.ribbonTarget,
      state = this.state
    if (!target || !state) throw new Error(t('appSelectSourceRangeFirst'))
    const bounds = parseRange(target.address)
    if (bounds.endRow <= bounds.startRow) throw new Error(t('appSelectRangeWithHeader'))
    if (rangeCellCount(bounds) > 2000)
      throw new Error('Subtotal supports at most 2000 source cells.')
    const frozen = { stateId: state.stateId, revision: this.revision }
    const data = (await this.host.read(
      target.sheetId,
      { ...bounds, endRow: bounds.startRow },
      frozen.stateId,
    )) as { cells: ViewCell[] }
    if (this.closed || this.state?.stateId !== frozen.stateId || this.revision !== frozen.revision)
      throw new Error('The document changed while reading Subtotal headers. Try again.')
    const cells = new Map(data.cells.map((cell) => [cell.address, cell]))
    return {
      ...frozen,
      target: { sheetId: target.sheetId, address: target.address },
      fields: Array.from(
        { length: Math.min(200, bounds.endColumn - bounds.startColumn + 1) },
        (_, offset) => {
          const colIndex = bounds.startColumn + offset
          const cell = cells.get(formatAddress(bounds.startRow, colIndex))
          if (
            cell?.formula &&
            (cell.calculationState !== 'ready' || !Object.hasOwn(cell, 'cachedValue'))
          )
            throw new Error('Wait for Subtotal headers to finish calculating, then retry.')
          const value = cell?.formula ? cell.cachedValue : cell?.rawValue
          return {
            colIndex,
            label:
              value === null || value === undefined || value === ''
                ? t('appColumnLabel', { col: columnLabel(colIndex) })
                : String(value),
            numeric: false,
          }
        },
      ),
    }
  }
  private subtotalWork: Promise<void> | undefined
  subtotal(
    config: SubtotalConfig,
    context: Awaited<ReturnType<WorkbookRuntime['subtotalContext']>>,
  ) {
    if (this.subtotalWork) return Promise.reject(new Error('Subtotal is already being applied.'))
    const work = (async () => {
      if (!(await this.flushCellEditor()))
        throw new Error('Finish editing before creating subtotals.')
      if (
        this.closed ||
        this.state?.stateId !== context.stateId ||
        this.revision !== context.revision
      )
        throw new Error('The document changed while Subtotal was open. Reopen the dialog.')
      if (
        !context.fields.some((field) => field.colIndex === config.groupCol) ||
        !context.fields.some((field) => field.colIndex === config.valueCol)
      )
        throw new Error('Invalid Subtotal column.')
      await this.enqueue(
        [
          {
            op: 'subtotal',
            sheetId: context.target.sheetId,
            range: context.target.address,
            groupColumn: columnLabel(config.groupCol),
            valueColumn: columnLabel(config.valueCol),
            agg: config.agg,
          },
        ],
        context,
        { failureInDialog: true },
      )
    })()
    this.subtotalWork = work
    const clear = () => {
      if (this.subtotalWork === work) this.subtotalWork = undefined
    }
    void work.then(clear, clear)
    return work
  }
  private outlineGutter: OutlineGutterHandle | undefined
  private outlineProjection:
    | { data: ReturnType<typeof outlineReadResultSchema.parse>; model: OutlineGutterState }
    | undefined
  private outlineLoading: string | undefined
  private async loadOutlineProjection() {
    const state = this.state
    const sheetId = this.runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId()
    if (this.closed || !state || !sheetId || !this.host.outline || !this.outlineGutter) return
    if (
      this.outlineProjection?.data.stateId === state.stateId &&
      this.outlineProjection.data.sheetId === sheetId
    )
      return
    const key = `${state.stateId}:${sheetId}`
    if (this.outlineLoading === key) return
    this.outlineLoading = key
    try {
      const data = outlineReadResultSchema.parse(await this.host.outline(sheetId, state.stateId))
      if (
        this.closed ||
        this.state?.stateId !== state.stateId ||
        this.runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId() !== sheetId
      )
        return
      if (data.stateId !== state.stateId || data.sheetId !== sheetId)
        throw new Error('invalid_outline_projection')
      this.outlineProjection = {
        data,
        model: {
          rows: new Map(
            data.rows.map((entry) => [entry.index, { ...entry, fileHidden: entry.hidden }]),
          ),
          cols: new Map(
            data.cols.map((entry) => [entry.index, { ...entry, fileHidden: entry.hidden }]),
          ),
          summaryBelow: data.summaryBelow,
          summaryRight: data.summaryRight,
          version: this.revision,
        },
      }
      this.outlineGutter.refresh()
    } catch (error) {
      if (
        !this.closed &&
        this.state?.stateId === state.stateId &&
        this.runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId() === sheetId
      )
        throw error
    } finally {
      if (this.outlineLoading === key) this.outlineLoading = undefined
    }
  }
  get outlineSummary() {
    const features = this.selectedFeatures
    const original = features.originalMetadata as
      { outlineSummaryBelow?: boolean; outlineSummaryRight?: boolean } | undefined
    let below = original?.outlineSummaryBelow !== false,
      right = original?.outlineSummaryRight !== false
    for (const edit of (features.outlineEdits ?? []) as {
      op: string
      summaryBelow?: boolean
      summaryRight?: boolean
    }[])
      if (edit.op === 'set_outline_settings') {
        below = edit.summaryBelow ?? below
        right = edit.summaryRight ?? right
      }
    return { below, right }
  }
  private outlineWork: Promise<void> | undefined
  outlineCommand(
    action: OutlineAction,
    axis: 'rows' | 'cols' = 'rows',
    context?: OutlineGutterTarget,
  ) {
    if (this.outlineWork)
      return Promise.reject(new Error('An outline operation is already running.'))
    const work = this.runOutlineCommand(action, axis, context)
    this.outlineWork = work
    const clear = () => {
      if (this.outlineWork === work) this.outlineWork = undefined
    }
    void work.then(clear, clear)
    return work
  }
  private async runOutlineCommand(
    action: OutlineAction,
    axis: 'rows' | 'cols',
    context?: OutlineGutterTarget,
  ) {
    if (!(await this.flushCellEditor()))
      throw new Error('Finish editing before changing outline groups.')
    const target = this.ribbonTarget,
      state = this.state
    if (!target || !state) throw new Error(t('appSelectCellFirst'))
    if (context && (context.stateId !== state.stateId || context.sheetId !== target.sheetId))
      throw new Error('The outline changed. Try again.')
    if (!this.host.outline) throw new Error('Complete outline data is unavailable in this host.')
    const frozen = { stateId: state.stateId, revision: this.revision }
    const outline = outlineReadResultSchema.parse(
      await this.host.outline(target.sheetId, frozen.stateId),
    )
    if (
      this.closed ||
      this.state?.stateId !== frozen.stateId ||
      this.revision !== frozen.revision ||
      outline.stateId !== frozen.stateId ||
      outline.sheetId !== target.sheetId
    )
      throw new Error('The document changed while reading outline groups. Try again.')
    const operations: WorkbookOperation[] = []
    const bounds = parseRange(target.address)
    const entriesFor = (axis: 'rows' | 'cols') =>
      new Map((axis === 'rows' ? outline.rows : outline.cols).map((entry) => [entry.index, entry]))
    const addVisibility = (axis: 'rows' | 'cols', start: number, end: number, hidden: boolean) => {
      const limit = axis === 'rows' ? 10_000 : 1000
      for (let index = start; index <= end; index += limit) {
        const count = Math.min(limit, end - index + 1)
        operations.push(
          axis === 'rows'
            ? { op: 'set_rows_hidden', sheetId: target.sheetId, row: index + 1, count, hidden }
            : {
                op: 'set_cols_hidden',
                sheetId: target.sheetId,
                column: formatAddress(0, index).replace(/1$/, ''),
                count,
                hidden,
              },
        )
      }
    }
    if (action === 'clear') {
      const whole = bounds.startRow === bounds.endRow && bounds.startColumn === bounds.endColumn
      for (const axis of ['rows', 'cols'] as const) {
        const entries = entriesFor(axis)
        const groups = outlineGroups(
          entries,
          axis === 'rows' ? outline.summaryBelow : outline.summaryRight,
          (index) => entries.get(index)?.hidden === true,
        )
        const from = whole ? 0 : axis === 'rows' ? bounds.startRow : bounds.startColumn
        const to = whole ? 1048575 : axis === 'rows' ? bounds.endRow : bounds.endColumn
        // Ignore source-only hidden lines: clearing outline must preserve ordinary manual hiding.
        const indices = [...entries.keys()].filter(
          (index) =>
            index >= from &&
            index <= to &&
            (entries.get(index)!.level > 0 || entries.get(index)!.collapsed),
        )
        indices.sort((a, b) => a - b)
        const spans: { start: number; end: number; show: boolean }[] = []
        for (const index of indices) {
          const show = groups.some(
            (group) => group.collapsed && group.start <= index && group.end >= index,
          )
          const last = spans[spans.length - 1]
          if (last && last.end + 1 === index && last.show === show) last.end = index
          else spans.push({ start: index, end: index, show })
        }
        for (const span of spans) {
          if (span.show) addVisibility(axis, span.start, span.end, false)
          for (let index = span.start; index <= span.end; index += 10_000)
            operations.push({
              op: 'set_outline',
              sheetId: target.sheetId,
              axis,
              start: index + 1,
              count: Math.min(10_000, span.end - index + 1),
              level: 0,
              collapsed: false,
            })
        }
      }
      if (!operations.length) throw new Error(t('appNothingToUngroup'))
    } else if (action === 'hide-detail' || action === 'show-detail' || action === 'level') {
      const start = axis === 'rows' ? bounds.startRow : bounds.startColumn
      const end = axis === 'rows' ? bounds.endRow : bounds.endColumn
      const entries = entriesFor(axis)
      const groups = outlineGroups(
        entries,
        axis === 'rows' ? outline.summaryBelow : outline.summaryRight,
        (index) => entries.get(index)?.hidden === true,
      )
      const group = context?.group
        ? groups.find(
            (candidate) =>
              candidate.start === context.group!.start &&
              candidate.end === context.group!.end &&
              candidate.level === context.group!.level &&
              candidate.summary === context.group!.summary,
          )
        : groups
            .filter(
              (candidate) =>
                (candidate.start <= end && candidate.end >= start) ||
                (candidate.summary >= start && candidate.summary <= end),
            )
            .sort((a, b) => b.level - a.level)[0]
      if (action !== 'level' && !group) throw new Error(t('appOutlineSelectFirst'))
      const level = context?.level
      if (action === 'level' && (!Number.isInteger(level) || level! < 1 || level! > 8))
        throw new Error('invalid_outline_level')
      const plan =
        action === 'level'
          ? levelButtonPlan(entries, groups, level!)
          : groupTogglePlan(groups, group!, action === 'hide-detail')
      for (const span of plan.show) addVisibility(axis, span.start, span.end, false)
      for (const span of plan.hide) addVisibility(axis, span.start, span.end, true)
      for (const summary of plan.summaries) {
        if (summary.index >= (axis === 'rows' ? 1048576 : 16384)) continue
        operations.push({
          op: 'set_outline',
          sheetId: target.sheetId,
          axis,
          start: summary.index + 1,
          count: 1,
          level: entries.get(summary.index)?.level ?? 0,
          collapsed: summary.collapsed,
        })
      }
    } else if (action === 'summary-below' || action === 'summary-right') {
      operations.push({
        op: 'set_outline_settings',
        sheetId: target.sheetId,
        summaryBelow: action === 'summary-below' ? !outline.summaryBelow : outline.summaryBelow,
        summaryRight: action === 'summary-right' ? !outline.summaryRight : outline.summaryRight,
      })
    } else {
      const start = axis === 'rows' ? bounds.startRow : bounds.startColumn
      const end = axis === 'rows' ? bounds.endRow : bounds.endColumn
      const entries = new Map(
        (axis === 'rows' ? outline.rows : outline.cols).map((entry) => [entry.index, entry]),
      )
      const spans = groupLevelOps(entries, start, end, action === 'group' ? 1 : -1)
      for (const span of spans)
        for (let index = span.start; index <= span.end; index += 10_000)
          operations.push({
            op: 'set_outline',
            sheetId: target.sheetId,
            axis,
            start: index + 1,
            count: Math.min(10_000, span.end - index + 1),
            level: span.level,
          })
      if (!operations.length)
        throw new Error(t(action === 'group' ? 'appOutlineMaxLevel' : 'appNothingToUngroup'))
      if (operations.length > 2000)
        throw new Error('The selection contains too many separate outline levels.')
    }
    if (operations.length > 2000)
      throw new Error('The outline action requires too many separate updates.')
    await this.enqueue(operations, frozen, { failureInDialog: true })
  }
  get goalSeekContext() {
    const selection = this.selection
    if (!selection || !this.state) return null
    return {
      stateId: this.state.stateId,
      revision: this.revision,
      sheetId: selection.sheetId,
      initialSetCell: formatAddress(selection.range.startRow, selection.range.startColumn),
    }
  }
  private goalSeekWork: Promise<GoalSeekResult> | undefined
  goalSeek(
    setCell: string,
    toValue: number,
    byCell: string,
    context: NonNullable<WorkbookRuntime['goalSeekContext']>,
  ) {
    if (this.goalSeekWork) return Promise.reject(new Error('Goal Seek is already running.'))
    const work = this.runGoalSeek(setCell, toValue, byCell, context)
    this.goalSeekWork = work
    const clear = () => {
      if (this.goalSeekWork === work) this.goalSeekWork = undefined
    }
    void work.then(clear, clear)
    return work
  }
  private async runGoalSeek(
    setCell: string,
    toValue: number,
    byCell: string,
    context: NonNullable<WorkbookRuntime['goalSeekContext']>,
  ): Promise<GoalSeekResult> {
    if (!(await this.flushCellEditor())) throw new Error('Finish editing before Goal Seek.')
    if (
      this.closed ||
      this.state?.stateId !== context.stateId ||
      this.revision !== context.revision
    )
      throw new Error('The document changed while Goal Seek was open. Reopen the dialog.')
    if (!this.host.goalSeek) throw new Error('Goal Seek is unavailable in this host.')
    const result = goalSeekResultSchema.parse(
      await this.host.goalSeek({
        stateId: context.stateId,
        revision: context.revision,
        sheetId: context.sheetId,
        setCell: setCell.toUpperCase(),
        byCell: byCell.toUpperCase(),
        toValue,
      }),
    )
    context.stateId = result.stateId
    context.revision = result.revision
    return {
      found: result.found,
      solution: result.solution,
      reached: result.reached ?? Number.NaN,
      iterations: result.iterations,
    }
  }
  async customSortContext() {
    if (!(await this.flushEditor())) throw new Error('Finish editing before sorting.')
    const target = this.ribbonTarget,
      state = this.state
    if (!target || !state) throw new Error(t('appSelectCellFirst'))
    const bounds = parseRange(target.address)
    if (rangeCellCount(bounds) > 2000)
      throw new Error('Sort supports at most 2,000 selected cells.')
    const frozen = { stateId: state.stateId, revision: this.revision }
    const result = (await this.host.read(
      target.sheetId,
      {
        ...bounds,
        endRow: bounds.startRow,
      },
      frozen.stateId,
    )) as { cells: ViewCell[] }
    if (this.closed || this.state?.stateId !== frozen.stateId || this.revision !== frozen.revision)
      throw new Error('The document changed while reading column labels.')
    const cells = new Map(result.cells.map((cell) => [cell.address, cell]))
    const columns = Array.from(
      { length: bounds.endColumn - bounds.startColumn + 1 },
      (_, offset) => {
        const colIndex = bounds.startColumn + offset
        const cell = cells.get(formatAddress(bounds.startRow, colIndex))
        const label = cell?.displayValue ?? String(cell?.rawValue ?? '')
        return {
          colIndex,
          label: label ? `${columnLabel(colIndex)} — ${label}` : columnLabel(colIndex),
        }
      },
    )
    return { ...frozen, target: { sheetId: target.sheetId, address: target.address }, columns }
  }
  async customSort(
    command: string,
    context: Awaited<ReturnType<WorkbookRuntime['customSortContext']>>,
  ) {
    if (!(await this.flushCellEditor())) throw new Error('Finish editing before sorting.')
    const match = /^sort-custom:([01]):([0-9]+[ad](?:,[0-9]+[ad]){0,2})$/.exec(command)
    if (!match) throw new Error('Invalid custom sort settings.')
    const levels = match[2]!.split(',').map((token) => ({
      byColumn: columnLabel(Number(token.slice(0, -1))),
      order: token.endsWith('a') ? ('asc' as const) : ('desc' as const),
    }))
    await this.enqueue(
      [
        {
          op: 'sort_range',
          sheetId: context.target.sheetId,
          range: context.target.address,
          byColumn: levels[0]!.byColumn,
          order: levels[0]!.order,
          hasHeader: match[1] === '1',
          levels,
        },
      ],
      context,
      { failureInDialog: true },
    )
  }
  private consolidateWork: Promise<{ rows: number; columns: number }> | undefined
  consolidate(
    config: ConsolidateConfig,
    context?: {
      target: { sheetId: string; address: string }
      stateId: string
      revision: number
    },
  ) {
    if (this.consolidateWork) return Promise.reject(new Error('Consolidate is already running.'))
    const work = this.runConsolidate(config, context)
    this.consolidateWork = work
    const clear = () => {
      if (this.consolidateWork === work) this.consolidateWork = undefined
    }
    void work.then(clear, clear)
    return work
  }
  private async runConsolidate(
    config: ConsolidateConfig,
    context?: {
      target: { sheetId: string; address: string }
      stateId: string
      revision: number
    },
  ) {
    if (!(await this.flushCellEditor()))
      throw new Error('Finish editing before consolidating ranges.')
    if (context && (this.state?.stateId !== context.stateId || this.revision !== context.revision))
      throw new Error('The document changed while Consolidate was open. Reopen the dialog.')
    const state = this.state,
      target = context?.target ?? this.ribbonTarget
    if (!state || !target) throw new Error(t('appSelectTargetCellFirst'))
    if (config.references.length === 0 || config.references.length > 50)
      throw new Error('Consolidate supports between 1 and 50 source ranges.')
    const frozen = context
      ? { stateId: context.stateId, revision: context.revision }
      : { stateId: state.stateId, revision: this.revision }
    const current = () =>
      !this.closed && this.state?.stateId === frozen.stateId && this.revision === frozen.revision
    const targetBounds = parseRange(target.address)
    const outputOrigin = { row: targetBounds.startRow, column: targetBounds.startColumn }
    const areas: ConsolidateArea[] = config.references.map((reference) => {
      const parsed = parseConsolidateReference(reference)
      if (!parsed) throw new Error(`Invalid Consolidate reference: ${reference}`)
      const name = parsed.sheetName ?? state.sheets.find((item) => item.id === target.sheetId)?.name
      const source = state.sheets.find((item) => item.name.toLowerCase() === name?.toLowerCase())
      if (!source) throw new Error(`Worksheet "${name ?? ''}" was not found.`)
      if (parsed.range.endRow >= MAX_GRID_ROWS || parsed.range.endColumn >= MAX_GRID_COLUMNS)
        throw new Error(`Consolidate source is outside the worksheet grid: ${reference}`)
      return {
        sheetName: source.name,
        startRow: parsed.range.startRow,
        startColumn: parsed.range.startColumn,
        rows: parsed.range.endRow - parsed.range.startRow + 1,
        columns: parsed.range.endColumn - parsed.range.startColumn + 1,
      }
    })

    const maxColumns = Math.max(...areas.map((area) => area.columns))
    if (!config.leftLabels && Math.max(...areas.map((area) => area.rows)) * maxColumns > 2000)
      throw new Error('Consolidate output exceeds the 2,000-cell transaction limit.')
    let matrix: ReturnType<typeof buildPositionMatrix>
    if (config.leftLabels) {
      const labelCount = areas.reduce((total, area) => total + area.rows, 0)
      if (labelCount > 200_000)
        throw new Error('Consolidate label matching is limited to 200,000 source rows.')
      const labelsByArea: string[][] = []
      for (const area of areas) {
        const sourceSheet = state.sheets.find((item) => item.name === area.sheetName)!
        const labels: string[] = []
        for (let startRow = area.startRow; startRow < area.startRow + area.rows; startRow += 2000) {
          const endRow = Math.min(area.startRow + area.rows - 1, startRow + 1999)
          const result = (await this.host.read(
            sourceSheet.id,
            {
              startRow,
              endRow,
              startColumn: area.startColumn,
              endColumn: area.startColumn,
            },
            frozen.stateId,
          )) as { cells: ViewCell[] }
          if (!current()) throw new Error('The document changed while reading Consolidate sources.')
          const byAddress = new Map(result.cells.map((cell) => [cell.address, cell]))
          for (let row = startRow; row <= endRow; row++) {
            const cell = byAddress.get(formatAddress(row, area.startColumn))
            if (!cell) {
              labels.push('')
              continue
            }
            if (cell.kind === 'error' || cell.calculationState === 'error')
              throw new Error('Consolidate cannot group error-valued labels.')
            if (cell.formula && cell.calculationState !== 'ready')
              throw new Error('Wait for source formulas to calculate before consolidating.')
            if (cell.formula && !Object.hasOwn(cell, 'cachedValue'))
              throw new Error(
                'A Consolidate label formula has no current result. Recalculate and retry.',
              )
            labels.push(String((cell.formula ? cell.cachedValue : cell.rawValue) ?? ''))
          }
        }
        labelsByArea.push(labels)
      }
      const uniqueLabels = new Set(labelsByArea.flat().filter((label) => label !== ''))
      if (uniqueLabels.size * maxColumns > 2000)
        throw new Error('Consolidate output exceeds the 2,000-cell transaction limit.')
      const built = buildLabelMatrix(config.fn, areas, labelsByArea, outputOrigin)
      if (typeof built === 'string') throw new Error(built)
      matrix = built
    } else {
      matrix = buildPositionMatrix(config.fn, areas)
    }
    const rows = matrix.length,
      columns = matrix[0]?.length ?? 0
    if (!rows || !columns) throw new Error('Consolidate sources contain no output cells.')
    if (rows * columns > 2000)
      throw new Error('Consolidate output exceeds the 2,000-cell transaction limit.')
    if (outputOrigin.row + rows > MAX_GRID_ROWS || outputOrigin.column + columns > MAX_GRID_COLUMNS)
      throw new Error('Consolidate output exceeds the worksheet grid.')
    const targetSheet = state.sheets.find((item) => item.id === target.sheetId)
    if (!targetSheet) throw new Error('The destination worksheet no longer exists.')
    if (
      targetOverlapsSource(areas, targetSheet.name, {
        ...outputOrigin,
        rows,
        columns,
      })
    )
      throw new Error('The Consolidate destination overlaps a source range.')
    const destination = {
      startRow: outputOrigin.row,
      endRow: outputOrigin.row + rows - 1,
      startColumn: outputOrigin.column,
      endColumn: outputOrigin.column + columns - 1,
    }
    const targetPage = (await this.host.read(
      target.sheetId,
      {
        startRow: destination.startRow,
        endRow: destination.startRow,
        startColumn: destination.startColumn,
        endColumn: destination.startColumn,
      },
      frozen.stateId,
    )) as { merges?: GridRange[] }
    if (!current()) throw new Error('The document changed while checking the destination range.')
    if (
      targetPage.merges?.some(
        (merge) =>
          merge.startRow <= destination.endRow &&
          merge.endRow >= destination.startRow &&
          merge.startColumn <= destination.endColumn &&
          merge.endColumn >= destination.startColumn,
      )
    )
      throw new Error('The Consolidate destination cannot contain merged cells.')
    const operations: WorkbookOperation[] = matrix.flatMap((line, row) =>
      line.map((cell, column): WorkbookOperation => {
        const address = formatAddress(outputOrigin.row + row, outputOrigin.column + column)
        return cell.f === undefined
          ? { op: 'set_cell', sheetId: target.sheetId, address, value: cell.v ?? null }
          : { op: 'set_formula', sheetId: target.sheetId, address, formula: cell.f }
      }),
    )
    await this.enqueue(operations, frozen, { failureInDialog: true })
    return { rows, columns }
  }
  async setDataValidation(
    validation: DvRule | null,
    context?: {
      target: { sheetId: string; address: string }
      stateId: string
      revision: number
    },
  ) {
    if (!(await this.flushCellEditor()))
      throw new Error('Finish editing before changing data validation.')
    if (context && (this.state?.stateId !== context.stateId || this.revision !== context.revision))
      throw new Error('The document changed while the dialog was open. Reopen Data Validation.')
    const target = context?.target ?? this.ribbonTarget,
      state = this.state
    if (!target || !state) throw new Error(t('appSelectCellFirst'))
    const frozen = context
      ? { stateId: context.stateId, revision: context.revision }
      : { stateId: state.stateId, revision: this.revision }
    await this.enqueue(
      [
        {
          op: 'set_data_validation',
          sheetId: target.sheetId,
          range: target.address,
          validation,
        },
      ],
      frozen,
    )
    const workbook = this.runtime.univerAPI.getActiveWorkbook(),
      sheet = workbook?.getSheetBySheetId(target.sheetId)
    if (!workbook || !sheet) throw new Error('The worksheet view is unavailable.')
    let built: ReturnType<
      ReturnType<typeof this.runtime.univerAPI.newDataValidation>['build']
    > | null = null
    if (validation) {
      const builder = this.runtime.univerAPI.newDataValidation()
      const ruleOptions = validation.options ?? {}
      const options = {
        allowBlank: ruleOptions.allowBlank ?? true,
        showInputMessage: ruleOptions.showInputMessage ?? false,
        showErrorMessage: ruleOptions.showErrorMessage ?? true,
        ...(ruleOptions.showDropdown === undefined
          ? {}
          : { showDropDown: ruleOptions.showDropdown }),
        ...(ruleOptions.promptTitle === undefined ? {} : { promptTitle: ruleOptions.promptTitle }),
        ...(ruleOptions.prompt === undefined ? {} : { prompt: ruleOptions.prompt }),
        ...(ruleOptions.errorTitle === undefined ? {} : { errorTitle: ruleOptions.errorTitle }),
        ...(ruleOptions.error === undefined ? {} : { error: ruleOptions.error }),
        ...(ruleOptions.errorStyle === undefined
          ? {}
          : {
              errorStyle: {
                information: DataValidationErrorStyle.INFO,
                stop: DataValidationErrorStyle.STOP,
                warning: DataValidationErrorStyle.WARNING,
              }[ruleOptions.errorStyle],
            }),
      }
      let listSource: { sheet: NonNullable<typeof sheet>; range: string } | undefined
      if (validation.kind === 'listRef') {
        const qualified = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(validation.range)
        if (qualified) {
          const sourceSheetName = (qualified[1] ?? qualified[2]!).replaceAll("''", "'")
          const sourceSheet = workbook.getSheetByName(sourceSheetName)
          if (!sourceSheet)
            throw new Error(`The data-validation source sheet "${sourceSheetName}" is missing.`)
          listSource = { sheet: sourceSheet, range: qualified[3]! }
        } else listSource = { sheet, range: validation.range }
        listSource.range = listSource.range.replaceAll('$', '')
      }
      if (validation.kind === 'criterion' || validation.kind === 'any') {
        built = builder.requireFormulaSatisfied('=TRUE').setOptions(options).build()
        Object.assign(
          (built as unknown as { rule: Record<string, unknown> }).rule,
          dvRuleToWireRule(validation),
        )
      } else {
        const configured =
          validation.kind === 'list'
            ? builder.requireValueInList([...validation.values], false, true)
            : validation.kind === 'listRef'
              ? builder.requireValueInRange(
                  listSource!.sheet.getRange(listSource!.range),
                  false,
                  true,
                )
              : validation.kind === 'numberBetween'
                ? builder.requireNumberBetween(validation.min, validation.max)
                : validation.kind === 'dateBetween'
                  ? builder.requireDateBetween(new Date(validation.start), new Date(validation.end))
                  : validation.kind === 'checkbox'
                    ? builder.requireCheckbox()
                    : builder.requireFormulaSatisfied(validation.formula)
        built = configured.setOptions(options).build()
      }
    }
    this.internal++
    try {
      sheet.getRange(target.address).setDataValidation(built)
      this.runtime.univer.__getInjector().get(IUndoRedoService).clearUndoRedo(workbook.getId())
    } finally {
      this.internal--
    }
  }
  dataValidationDraft(sheetId: string, address: string): DvRule | undefined {
    const rule = this.runtime.univerAPI
      .getActiveWorkbook()
      ?.getSheetBySheetId(sheetId)
      ?.getRange(address)
      .getDataValidation()?.rule
    if (!rule) return undefined
    const formula1 = rule.formula1?.trim()
    const candidate = (value: unknown) => {
      const parsed = dvRuleSchema.safeParse(value)
      return parsed.success ? parsed.data : undefined
    }
    const errorStyle =
      rule.errorStyle === DataValidationErrorStyle.INFO
        ? 'information'
        : rule.errorStyle === DataValidationErrorStyle.WARNING
          ? 'warning'
          : rule.errorStyle === DataValidationErrorStyle.STOP
            ? 'stop'
            : undefined
    const options = {
      ...(rule.allowBlank === undefined ? {} : { allowBlank: rule.allowBlank }),
      ...(rule.showInputMessage === undefined ? {} : { showInputMessage: rule.showInputMessage }),
      ...(rule.promptTitle === undefined ? {} : { promptTitle: rule.promptTitle }),
      ...(rule.prompt === undefined ? {} : { prompt: rule.prompt }),
      ...(rule.showErrorMessage === undefined ? {} : { showErrorMessage: rule.showErrorMessage }),
      ...(errorStyle === undefined ? {} : { errorStyle }),
      ...(rule.errorTitle === undefined ? {} : { errorTitle: rule.errorTitle }),
      ...(rule.error === undefined ? {} : { error: rule.error }),
      ...(rule.type !== 'list' || rule.showDropDown === undefined
        ? {}
        : { showDropdown: rule.showDropDown }),
    }
    const withOptions = (value: Record<string, unknown>) => ({
      ...value,
      ...(Object.keys(options).length ? { options } : {}),
    })
    if (rule.type === 'any') return candidate(withOptions({ kind: 'any' }))
    if (rule.type === 'checkbox') return candidate(withOptions({ kind: 'checkbox' }))
    if (rule.type === 'list' && formula1 !== undefined) {
      if (formula1.startsWith('=')) {
        const range = formula1.slice(1)
        if (
          /^[A-Z]{1,3}[1-9][0-9]{0,6}(?::[A-Z]{1,3}[1-9][0-9]{0,6})?$/.test(range) ||
          /^(?:'(?:[^']|'')+'|[A-Za-z_][\w. ]*)!\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}(?::\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6})?$/.test(
            range,
          )
        )
          return candidate(withOptions({ kind: 'listRef', range }))
      }
      if (formula1.startsWith('"') && formula1.endsWith('"')) {
        const values = formula1
          .slice(1, -1)
          .split(',')
          .map((item) => item.trim().replaceAll('""', '"'))
        if (values.every((value) => value.length > 0))
          return candidate(withOptions({ kind: 'list', values }))
      }
    }
    if (
      rule.type === 'decimal' &&
      rule.operator === 'between' &&
      rule.formula1 !== undefined &&
      rule.formula2 !== undefined
    ) {
      const min = Number(rule.formula1),
        max = Number(rule.formula2)
      if (Number.isFinite(min) && Number.isFinite(max))
        return candidate(withOptions({ kind: 'numberBetween', min, max }))
    }
    if (rule.type === 'custom' && formula1 !== undefined)
      return candidate(
        withOptions({
          kind: 'formula',
          formula: formula1.startsWith('=') ? formula1 : `=${formula1}`,
        }),
      )
    if (
      ['whole', 'decimal', 'date', 'time', 'textLength'].includes(rule.type) &&
      rule.operator &&
      formula1 !== undefined
    )
      return candidate(
        withOptions({
          kind: 'criterion',
          type: rule.type,
          operator: rule.operator,
          formula1,
          ...(rule.formula2 === undefined ? {} : { formula2: rule.formula2 }),
        }),
      )
    return undefined
  }
  private dvListSourceContext(state: CandidateState, targetSheetId: string): DvListSourceContext {
    return {
      targetSheetId,
      definedNames: sessionDefinedNames(state).names,
      sheets: state.sheets.map((sheet) => {
        const metadata = sheet.features?.originalMetadata as
          { tables?: DvListSourceContext['sheets'][number]['tables'] } | undefined
        return { id: sheet.id, name: sheet.name, tables: metadata?.tables ?? [] }
      }),
    }
  }
  get calculationManual() {
    return isManualCalculation(this.runtime)
  }
  setCalculationManual(manual: boolean) {
    setManualCalculation(this.runtime, manual)
    this.changed()
    if (!manual) this.recalculate()
    // The mode is part of the workbook: it enters history and the saved calcPr.
    // A rejected edit is surfaced by the pending-error state `enqueue` records.
    this.enqueue([{ op: 'set_calculation_mode', mode: manual ? 'manual' : 'auto' }]).catch(() => {})
  }
  enqueue(
    operations: readonly WorkbookOperation[],
    frozen?: { stateId: string; revision: number },
    options?: { failureInDialog: boolean },
  ) {
    // Preserve character provenance for all browser resize paths, including
    // header drags and legacy numeric inputs, using this workbook's font metrics.
    const mdw = measureNormalFontMdw(
      (this.state?.features?.originalMetadata ?? {}) as NormalFontDescriptor,
    )
    operations = operations.map((op) =>
      op.op === 'set_col_width' && op.widthCharacters === undefined
        ? {
            ...op,
            widthCharacters: Math.max(0, Math.round((op.widthPx / mdw) * 256) / 256),
          }
        : op,
    )
    const result = this.pending.then(() => {
      if (frozen && (this.state?.stateId !== frozen.stateId || this.revision !== frozen.revision))
        throw new Error('The document changed before applying this operation. Try again.')
      return this.submit(operations)
    })
    this.pending = result.then(
      () => {
        this.pendingError = false
      },
      () => {
        // A rejected dialog transaction keeps its editable draft and surfaces
        // the rejection itself. It must permit retry without acknowledging or
        // hiding a failed cell-editor command.
        this.pendingError = !options?.failureInDialog
      },
    )
    return result
  }
  /** Several bounded stages, published by the host as one revision and history entry. */
  enqueueStaged(
    stages: readonly (readonly WorkbookOperation[])[],
    frozen: { stateId: string; revision: number },
  ) {
    const result = this.pending.then(() => {
      if (this.state?.stateId !== frozen.stateId || this.revision !== frozen.revision)
        throw new Error('The document changed before applying this operation. Try again.')
      if (!this.state || this.closed) throw new Error('view_unavailable')
      if (!this.host.submitStaged) throw new Error('Large edits are unavailable in this host.')
      return this.host.submitStaged(
        { transactionId: crypto.randomUUID(), expectedRevision: this.revision },
        stages,
      )
    })
    this.pending = result.then(
      () => {
        this.pendingError = false
      },
      () => {
        this.pendingError = false
      },
    )
    return result
  }
  private mergeWorkbooksWork: Promise<unknown> | undefined
  /** Data → Merge Workbooks; the host owns the picker, source reads and the staged commit. */
  mergeWorkbooks() {
    if (this.mergeWorkbooksWork)
      return Promise.reject(new Error('Workbooks are already being merged.'))
    const work = (async () => {
      const metadata = this.state?.features?.originalMetadata as
        { workbookProtection?: { lockStructure?: boolean } } | undefined
      if (metadata?.workbookProtection?.lockStructure) throw new Error(t('appMergeWorkbooksLocked'))
      if (!this.host.mergeWorkbooks) throw new Error('Merge Workbooks is unavailable in this host.')
      if (!(await this.flushCellEditor()))
        throw new Error('Finish editing before merging workbooks.')
      await this.pending
      return z
        .union([
          z.object({ canceled: z.literal(true) }),
          z.object({ sheets: z.number(), files: z.number(), sheetNames: z.array(z.string()) }),
        ])
        .parse(await this.host.mergeWorkbooks())
    })()
    this.mergeWorkbooksWork = work
    const clear = () => {
      if (this.mergeWorkbooksWork === work) this.mergeWorkbooksWork = undefined
    }
    void work.then(clear, clear)
    return work
  }
  requestHistory(direction: 'undo' | 'redo') {
    return requestHostHistory(this.host, direction, () => this.flushEditor())
  }
  async submit(operations: readonly WorkbookOperation[]) {
    if (!this.state || this.closed) throw new Error('view_unavailable')
    return this.host.submit({
      transactionId: crypto.randomUUID(),
      expectedRevision: this.revision,
      operations,
    })
  }
  private activeEditorWork(): Promise<unknown>[] {
    return [
      this.clipboardWork,
      this.functionWork,
      this.nameWork,
      this.auditWork,
      this.pageWork,
      this.sortFilterWork,
      this.filterPopupWork,
      this.advancedFilterWork,
      this.removeDuplicatesWork,
      this.textToColumnsWork,
      this.csvImportWork,
      this.mergeWorkbooksWork,
      this.pivotRefreshWork,
      this.consolidateWork,
      this.goalSeekWork,
      this.outlineWork,
      this.subtotalWork,
      this.symbolWork,
      this.chartWork,
      this.visualWork,
    ].filter((work): work is Promise<unknown> => work !== undefined)
  }
  async flushEditor(options?: { commitChartDraft?: boolean }) {
    // Retain in-flight results before flushing the cell editor: that flush may
    // await a recoverable host rejection whose work slot has already cleared.
    let work = this.activeEditorWork()
    const active = this.container.ownerDocument.activeElement
    if (
      options?.commitChartDraft !== false &&
      active instanceof HTMLElement &&
      ((active instanceof HTMLInputElement && active.closest('.chart-format-pane')) ||
        active.closest('.shape-text-editor')) &&
      this.container.closest('.workbook-editor')?.contains(active)
    )
      active.blur()
    while (!this.closed && (isVisualDragActive() || isChartApplyActive() || isSheetShapeDragging()))
      await new Promise<void>((resolve) => setTimeout(resolve, 16))
    if (this.closed || !(await this.flushCellEditor())) return false
    // An open note popup's text commits before Save and other flushing commands.
    if (this.dirtyNotes.size)
      try {
        await this.flushNotes()
      } catch {
        return false
      }
    for (;;) {
      try {
        await Promise.all(work)
      } catch {
        return false
      }
      if (this.closed || !(await this.flushCellEditor())) return false
      work = this.activeEditorWork()
      if (!work.length) return true
    }
  }
  private async flushCellEditor() {
    if (this.composing) return false
    const book = this.runtime.univerAPI.getActiveWorkbook()
    if (book?.isCellEditing())
      try {
        await book.endEditingAsync(true)
      } catch {
        /* Value command was forwarded to the host. */
      }
    await this.pending
    return (
      !this.pendingError &&
      !this.composing &&
      !this.runtime.univerAPI.getActiveWorkbook()?.isCellEditing()
    )
  }
  dispose() {
    this.shapeDrawDispose?.()
    if (this.closed) return
    this.closeFilterPopup()
    this.closed = true
    this.findReplace.dispose()
    this.clearTraceArrows()
    this.clearChartVisuals()
    clearVisualSelection()
    this.themeObserver.disconnect()
    this.systemTheme.removeEventListener('change', this.updateTheme)
    if (this.loadTimer) clearTimeout(this.loadTimer)
    this.container.removeEventListener('compositionstart', this.onCompositionStart)
    this.container.removeEventListener('compositionend', this.onCompositionEnd)
    for (const disposable of [...this.disposables, ...this.projectionDisposables])
      disposable.dispose()
    // Univer owns a nested React root; release it after a parent React cleanup.
    queueMicrotask(() => this.runtime.univer.dispose())
  }
}

function tableErrorMessage(error: unknown): string {
  const message = error instanceof Error ? error.message : String(error)
  const match = /table:(\w+)(?::(\{[^}]*\}))?/.exec(message)
  const key = match
    ? (
        {
          needsRows: 'appTableNeedsRows',
          tooWide: 'appTableTooWide',
          overlap: 'appTableOverlapsSession',
          nameUsed: 'appTableNameUsed',
          addedSheet: 'appTablesNeedFile',
        } as const
      )[match[1] as 'needsRows']
    : undefined
  if (!key) return message
  let params: Record<string, string>
  try {
    params = match?.[2] ? (JSON.parse(match[2]) as Record<string, string>) : {}
  } catch {
    params = {}
  }
  return t(key, { name: '', ...params })
}

/** A cell's projected number format, for series and painter seeds. */
function sheetCellFormat(
  workbook: {
    getSheetBySheetId(
      id: string,
    ): { getRange(row: number, column: number): { getCellStyleData(): unknown } } | null
  },
  sheetId: string,
  row: number,
  column: number,
): string | undefined {
  const style = workbook.getSheetBySheetId(sheetId)?.getRange(row, column).getCellStyleData()
  const pattern = (style as { n?: { pattern?: string } } | null)?.n?.pattern
  return pattern && pattern !== 'General' ? pattern : undefined
}

/** Univer's conditional-format model mutations, captured as declarative session edits. */
const CF_MUTATIONS = new Set([
  'sheet.mutation.add-conditional-rule',
  'sheet.mutation.set-conditional-rule',
  'sheet.mutation.delete-conditional-rule',
  'sheet.mutation.move-conditional-rule',
])

/** Univer's row/column visibility commands (Ctrl+9 / Ctrl+0 and their Shift variants). */
const HIDE_COMMANDS = new Set([
  'sheet.command.set-rows-hidden',
  'sheet.command.set-specific-rows-visible',
  'sheet.command.set-col-hidden',
  'sheet.command.set-col-visible-on-cols',
])
