import { platformShortcuts } from '@genoffice/i18n'
import {
  Dropdown,
  aiPanelInitiallyOpen,
  rememberAiPanelOpen,
  useAiSidebarEnabled,
  useRibbonCollapse,
  type AiTokenUsageView,
} from '@genoffice/ui'
import type { IFunctionInfo } from '@univerjs/engine-formula'
import { useEffect, useMemo, useRef, useState } from 'react'
import {
  SheetsRibbon as Ribbon,
  ribbonTabs,
  useEscapeClose,
  type PageLayoutEcho,
  type RibbonTab,
  type SelectedChartRibbon,
  type SelectedShapeRibbon,
} from './SheetsRibbon'

import { RedoIcon, SaveAsIcon, SaveIcon, UndoIcon } from './ribbon-icons'
import { sheetContextMenuAction } from './sheet-context-menu'

import type { SheetProtectionAllow } from '@genoffice/xlsx-gateway/gateway/xlsx-protection'
import { AllowEditRangesDialog } from './AllowEditRangesDialog'
import { registerPasteSpecialOpener, setContextMenuCommandSink } from './cell-context-menu'
import {
  isFormulaBarKeyTarget,
  isGridKeyTarget,
  shouldInterceptClearSelection,
} from './clear-selection-keyboard'
import { FormatCellsDialog } from './FormatCellsDialog'
import { isModalOpen, resolveGlobalShortcut } from './global-shortcuts'
import { GoToDialog } from './GoToDialog'
import { useI18n, type StringKey } from './i18n/locale'
import { NameManagerDialog, type DefinedNameAction, type DefinedNameRow } from './NameManagerDialog'
import {
  PASTE_SPECIAL_OPEN_SHELL_COMMAND,
  type PasteSpecialOptions,
  type PasteSpecialSource,
} from './paste-special'
import { PasteSpecialDialog } from './PasteSpecialDialog'
import { ProtectSheetDialog } from './ProtectSheetDialog'
import { numberFormatLabel, type SelectionFormat } from './selection-format'
import type { SheetTabActions } from './sheet-tab-actions'
import { SheetTabContextMenu } from './SheetTabContextMenu'

import type { ChangePlan } from '@genoffice/xlsx-gateway/domain/workbook.types'
import type { AttachmentMeta } from '../shared/desktop-api'
import { AiChatPanel, type AiChatMessage } from './ai/AiChatPanel'
import { AiSelectionAsk } from './ai/AiSelectionAsk'
import type { SelectionAskAnchor } from './ai/selection-ask'
import type { ConsolidateConfig } from './consolidate'
import { ConsolidateDialog } from './ConsolidateDialog'
import type { FillSeriesOptions, FillSeriesType } from './fill-series'
import type { FillSeriesContext } from './fill-series-apply'
import { FillSeriesDialog } from './FillSeriesDialog'
import {
  registerFxButtonOpener,
  type FunctionArgumentsHost,
  type FunctionTarget,
} from './function-arguments-runtime'
import type { GoalSeekResult } from './goal-seek'
import { GoalSeekDialog } from './GoalSeekDialog'
import { HeaderFooterDialog, type HeaderFooterResult } from './HeaderFooterDialog'
import { InsertFunctionFlow } from './InsertFunctionFlow'
import { useModalDialog } from './modal-dialog'
import type { PrintPreviewHost } from './page-layout-actions'
import type { PivotEditSeed } from './pivot-actions'
import {
  modelToConfig,
  validatePivotModel,
  type PivotLayoutModel,
  type PivotModelError,
} from './pivot-field-model'
import {
  PivotDialog,
  type OoXmlPivotConfig,
  type PivotCreateRequest,
  type PivotField,
} from './PivotDialog'
import { PivotFieldPane } from './PivotFieldPane'
import { PrintDialog } from './PrintDialog'
import { STATUS_BAR_FUNCS, type StatusBarFunc } from './status-bar-stats'
import { SubtotalDialog, type SubtotalConfig } from './SubtotalDialog'
import type { SelectedTableRibbon } from './table-design-actions'
import { SHEET_ZOOM_MAX, SHEET_ZOOM_MIN, clampZoomPercent } from './zoom-range'
import { ZoomDialog } from './ZoomDialog'
const CONTEXTUAL_TABS: readonly RibbonTab[] = [
  'Chart Design',
  'Table Design',
  'Shape Format',
  'Picture Format',
]

// Internal values keep English ids (=== matching, state values unchanged);
// translated only for display
const TAB_LABEL: Record<RibbonTab, StringKey> = {
  Home: 'appTabHome',
  Insert: 'appTabInsert',
  'Page Layout': 'appTabPageLayout',
  Formulas: 'appTabFormulas',
  Data: 'appTabData',
  Review: 'appTabReview',
  View: 'appTabView',
  'Chart Design': 'appTabChartDesign',
  'Shape Format': 'appTabShapeFormat',
  'Table Design': 'appTabTableDesign',
  'Picture Format': 'appTabPictureFormat',
}

type ChartTextTarget = 'title' | 'axis-category' | 'axis-value'

const CHART_TEXT_LABELS: Record<ChartTextTarget, { heading: StringKey; command: string }> = {
  title: { heading: 'appChartTitleEl', command: 'chart-title' },
  'axis-category': { heading: 'appCategoryAxisTitle', command: 'chart-axis-cat' },
  'axis-value': { heading: 'appValueAxisTitle', command: 'chart-axis-val' },
}

/// mode=tab: embedded in the shell's tab strip, which owns the traffic
/// lights / caption buttons — the ribbon must not reserve space for them.
const IN_TAB = new URLSearchParams(window.location.search).get('mode') === 'tab'
const IS_MAC = navigator.platform.toLowerCase().includes('mac')

/// Univer's formula bar keeps its expanded state in React component state
/// with no command behind it; the fold arrow is the only way in.
function toggleFormulaBarExpand(): void {
  document
    .querySelector<HTMLElement>('[data-u-comp="formula-bar"] .univer-w-5.univer-cursor-pointer')
    ?.click()
}

interface ExcelShellProps {
  readonly openingWorkbook: boolean
  readonly prompt: string
  readonly preview: ChangePlan | null
  readonly selectionFormat: SelectionFormat | null
  readonly formatPainterActive: boolean
  /// True when the workbook has any cell content (the one-click AI action
  /// buttons are greyed out on an empty sheet).
  readonly sheetHasContent: boolean
  /// true while the real LLM agent is running (composer disabled meanwhile).
  readonly aiBusy: boolean
  /** tokens billed across the session, shown beside the composer */
  readonly sessionUsage?: AiTokenUsageView
  readonly chat: readonly AiChatMessage[]
  readonly historicChat?: readonly AiChatMessage[]
  /// Chat attachments (chips + 📎 button + drag-and-drop), same structure as the
  /// docs/slides AI panels.
  readonly attachments: readonly AttachmentMeta[]
  readonly attachedSkills: readonly string[]
  readonly onToggleSkill: (name: string) => void
  readonly attachNotice: string | null
  readonly onPickAttachments: () => void
  readonly onAddAttachmentPaths: (paths: readonly string[]) => void
  readonly onAddPastedImage: (data: ArrayBuffer, ext: string) => void
  readonly onRemoveAttachment: (path: string) => void
  readonly onPromptChange: (prompt: string) => void
  /** Send the composer text, or the given instruction when provided (Retry also
   *  resends that message's original attachments and passes the failed bubble's
   *  chat index so the send replaces it in place) */
  readonly onSend: (
    instruction?: string,
    attachments?: readonly AttachmentMeta[],
    retryIndex?: number,
  ) => void
  readonly onStop: () => void
  readonly onNewChat: () => void
  readonly onUndo: (steps?: number) => void
  /// A1 notation of the multi-cell selection the AI composer offers as this
  /// run's scope, or null when the resting single-cell selection carries none.
  readonly aiScopeRange: string | null
  /// Header names when that scope covers whole columns: they label the chip in
  /// place of the range, because a column is a name to the user, not a letter.
  readonly aiScopeColumns: readonly string[] | null
  /// The range above belongs to a run in flight and can no longer be dropped.
  readonly aiScopeLocked: boolean
  /// Drag endpoint and viewport bounds used to place the localized trigger.
  readonly aiSelectionAskAnchor: SelectionAskAnchor | null
  readonly onAiSelectionAskDismiss: () => void
  readonly onAiScopeDismiss: () => void
  /// Citation link in an AI answer: jumps the grid to the cited cell/range.
  readonly onAiCitation: (href: string) => void
  readonly onCommand: (command: string) => void
  /// True while Univer's in-cell editor is open (Backspace must delete
  /// characters, not clear the selection).
  readonly onIsCellEditing: () => boolean
  /// Left side of the status bar (ready / streaming / AI progress messages).
  readonly statusMessage: string
  readonly emptyCsvNotice: boolean
  readonly onOpenWorkbook: () => void
  readonly onDismissEmptyCsvNotice: () => void
  /// Zoom of the active sheet in percent, echoed by the status-bar slider.
  readonly zoomPercent: number
  /// Aggregates ticked in the status bar's right-click menu (Excel parity).
  readonly statusBarFuncs: readonly StatusBarFunc[]
  readonly onToggleStatusBarFunc: (func: StatusBarFunc) => void
  /// True when the edit journal has unsaved changes (enables the QAT Save).
  readonly canSave: boolean
  readonly onSave: () => void
  /// Save As remains available for a clean workbook, but requires a real
  /// file-backed session (the in-memory demo workbook has nowhere to copy).
  readonly canSaveAs: boolean
  readonly onSaveAs: () => void
  /// QAT redo (workbook history, same path as the app menu's ⇧⌘Z); undo
  /// shares the AI panel's onUndo above.
  readonly onRedo: () => void
  /// Undo/redo stack occupancy: the QAT buttons grey out when there is nothing to apply.
  readonly canUndo: boolean
  readonly canRedo: boolean
  /// AutoSave toggle in the tab row (docs/slides parity).
  readonly autoSave: boolean
  readonly onAutoSaveChange: (on: boolean) => void
  /// Non-null while a floating chart is selected in the grid.
  readonly selectedChart: SelectedChartRibbon | null
  /// Non-null while an editable shape is selected (Shape Format tab).
  readonly selectedShape: SelectedShapeRibbon | null
  /// Kind of the selected floating picture/shape (Shape/Picture Format tab).
  readonly selectedVisualKind: 'image' | 'shape' | null
  /// Non-null while the selection sits inside a table (Table Design tab).
  readonly selectedTable: SelectedTableRibbon | null
  /// Active sheet's outlinePr echo for the Data > Group settings toggles.
  readonly outlineSummary: { below: boolean; right: boolean }
  /// Bumps after each successful Create Table; the Table Design tab activates.
  readonly tableCreatedSeq: number
  /// Default range for the Ctrl+T dialog (selection widened to its data block).
  readonly onGetCreateTableRange: () => string
  /// Column choices of the active selection, read when the Sort dialog opens.
  readonly onGetSortColumns: () => { label: string; colIndex: number }[]
  /// Effective protection of the active sheet (null = unknown / demo).
  readonly onGetSheetProtection: () => { protected: boolean; hasPassword: boolean } | null
  /// Protect Sheet dialog result; resolve to an error message or null.
  readonly onProtectSheet: (request: {
    password: string | null
    allow: SheetProtectionAllow
  }) => Promise<string | null>
  readonly onUnprotectSheet: (password: string) => Promise<string | null>
  readonly sheetTabActions: SheetTabActions | null
  /// Effective workbook structure lock (null = no file open).
  readonly onGetWorkbookProtection: () => boolean | null
  readonly formulaBarVisible: boolean
  /// Cross-highlight ("reading mode") of the active row/column, echoed by the View checkbox.
  readonly crossHighlightVisible: boolean
  /// Allow-edit ranges of the active sheet, read when the dialog opens.
  readonly onGetProtectedRanges: () => {
    ranges: readonly { name: string; sqref: string; hasPassword: boolean }[]
    error: string | null
  }
  readonly onApplyProtectedRanges: (
    ranges: readonly { name: string; sqref: string }[],
  ) => string | null
  /// Name Manager data + actions (actions return an error message or null).
  readonly onGetDefinedNames: () => {
    names: DefinedNameRow[]
    sheets: { id: string; name: string }[]
    activeSheetId: string | null
  }
  readonly onDefinedNameAction: (action: DefinedNameAction) => string | null
  /// Subtotals use the selection; pivots pass their resolved source range.
  readonly onGetPivotFields: (sourceRange?: string) => PivotField[]
  readonly onGetSourceRange: () => string
  readonly onCreatePivot: (request: PivotCreateRequest) => string | null
  /// PivotTable Fields pane: seed of the pivot under the cursor (null = App
  /// already showed why), the apply callback, and the member list of one
  /// source field for the report-filter item picker.
  readonly onGetPivotEditSeed: () => PivotEditSeed | null
  readonly onEditPivot: (config: OoXmlPivotConfig) => string | null
  readonly onGetPivotFieldMembers: (fieldIndex: number) => string[]
  /// Identity of the pivot under the selection; the pane follows it like Excel.
  readonly pivotSelectionKey: string | null
  readonly onRefreshPivot: () => string | null
  readonly onIsSelectionInPivot: () => boolean
  readonly onGetActiveCell: () => string
  /// Value of the selection's top-left cell, read when Format Cells opens
  /// (number-format preview).
  readonly onGetAnchorValue: () => number | string | null
  /// A1 label of the active cell, echoed live by the Name Box.
  readonly activeCellA1: string
  /// Name Box / Go To jump; returns an error message, or null on success.
  readonly onGoToReference: (ref: string) => string | null
  readonly onListDefinedNames: () => readonly { name: string; ref: string }[]
  readonly onReadPasteSpecialSource: () => Promise<PasteSpecialSource>
  /// Resolves to an error message, or null on success.
  readonly onPasteSpecial: (
    source: PasteSpecialSource,
    options: PasteSpecialOptions,
  ) => Promise<string | null>
  /// Function descriptions from the running formula engine (Insert Function).
  readonly onListFunctions: () => readonly IFunctionInfo[]
  readonly functionArguments: FunctionArgumentsHost
  readonly onCreateSubtotal: (config: SubtotalConfig) => string | null
  readonly onCreateConsolidate: (config: ConsolidateConfig) => string | null
  /// Prefill for the Consolidate reference input (current multi-cell selection).
  readonly onGetConsolidateDefault: () => string
  /// Header & Footer dialog OK; returns an error message, or null on success.
  readonly onApplyHeaderFooter: (result: HeaderFooterResult) => string | null
  /// Print with Preview, open while non-null.
  readonly printHost: PrintPreviewHost | null
  readonly onClosePrint: () => void
  /// Session page-layout settings of the active sheet, echoed by the Page
  /// Layout tab's controls (untouched fields show the app default).
  readonly pageLayout: PageLayoutEcho
  /// Manual-recalc mode echo for the Calculation Options menu.
  readonly calcManual: boolean
  /// Goal Seek solve; rejects with a user-facing Error message.
  readonly onGoalSeek: (setCell: string, toValue: number, byCell: string) => Promise<GoalSeekResult>
  /// Fill > Series… OK; returns an error message, or null on success.
  readonly onFillSeries: (options: FillSeriesOptions) => string | null
  /// Series dialog defaults (selection shape, workbook date system).
  readonly onGetFillSeriesContext: () => FillSeriesContext
}

interface PivotPaneState {
  readonly seed: PivotEditSeed
  readonly model: PivotLayoutModel
  readonly deferred: boolean
  readonly dirty: boolean
  readonly error: string | null
}

const PIVOT_MODEL_ERRORS: Record<PivotModelError, StringKey> = {
  needRow: 'appNeedRowField',
  needValue: 'dlgPivotErrNeedValue',
  colNeedsSingleValue: 'dlgPivotErrColNeedsSingleValue',
  calcIncomplete: 'dlgPivotErrCalcFieldIncomplete',
  rangeStep: 'dlgPivotErrRangeStep',
  labelFilterText: 'dlgPivotErrLabelFilterText',
  oneValueFilter: 'dlgPivotErrOneValueFilter',
  filterConflict: 'dlgPivotErrFilterConflict',
  dupValueName: 'dlgPivotErrDupValueName',
}

export function ExcelShell({
  prompt,
  preview,
  selectionFormat,
  formatPainterActive,
  sheetHasContent,
  aiBusy,
  sessionUsage,
  chat,
  historicChat,
  attachments,
  attachedSkills,
  onToggleSkill,
  attachNotice,
  onPickAttachments,
  onAddAttachmentPaths,
  onAddPastedImage,
  onRemoveAttachment,
  onGetSortColumns,
  onGetSheetProtection,
  onProtectSheet,
  onUnprotectSheet,
  sheetTabActions,
  onGetWorkbookProtection,
  formulaBarVisible,
  crossHighlightVisible,
  onGetProtectedRanges,
  onApplyProtectedRanges,
  onGetDefinedNames,
  onDefinedNameAction,
  onGetPivotFields,
  onGetSourceRange,
  onCreatePivot,
  onGetPivotEditSeed,
  onEditPivot,
  onGetPivotFieldMembers,
  pivotSelectionKey,
  onRefreshPivot,
  onIsSelectionInPivot,
  onGetActiveCell,
  onGetAnchorValue,
  activeCellA1,
  onGoToReference,
  onListDefinedNames,
  onReadPasteSpecialSource,
  onPasteSpecial,
  functionArguments,
  onListFunctions,
  onCreateSubtotal,
  onCreateConsolidate,
  onGetConsolidateDefault,
  onApplyHeaderFooter,
  printHost,
  onClosePrint,
  onPromptChange,
  onSend,
  onStop,
  onNewChat,
  onUndo,
  aiScopeRange,
  aiScopeColumns,
  aiScopeLocked,
  aiSelectionAskAnchor,
  onAiSelectionAskDismiss,
  onAiScopeDismiss,
  onAiCitation,
  onCommand,
  onIsCellEditing,
  openingWorkbook,
  statusMessage,
  emptyCsvNotice,
  onOpenWorkbook,
  onDismissEmptyCsvNotice,
  zoomPercent,
  statusBarFuncs,
  onToggleStatusBarFunc,
  canSave,
  onSave,
  canSaveAs,
  onSaveAs,
  onRedo,
  canUndo,
  canRedo,
  autoSave,
  onAutoSaveChange,
  selectedChart,
  selectedShape,
  selectedTable,
  outlineSummary,
  selectedVisualKind,
  tableCreatedSeq,
  onGetCreateTableRange,
  pageLayout,
  calcManual,
  onGoalSeek,
  onFillSeries,
  onGetFillSeriesContext,
}: ExcelShellProps): React.JSX.Element {
  const { t } = useI18n()
  const [activeTab, setActiveTab] = useState<RibbonTab>('Home')
  const collapse = useRibbonCollapse('ai-sheets-ribbon-collapsed', {
    collapse: t('appRibbonCollapse'),
    expand: t('appRibbonExpand'),
    // ⌥⌘K / Ctrl+Alt+K: the icon-only middle density (docs parity, see #362)
    compact: t('appRibbonCompact'),
    expandFull: t('appRibbonExpand'),
  })
  // Persisted so a closed AI panel stays closed on next launch (docs/slides parity)
  const [isCopilotOpen, setIsCopilotOpen] = useState(() =>
    aiPanelInitiallyOpen('ai-sheets-show-ai'),
  )
  // the sidebar switch is a soft hide; it closes the panel when it turns off
  const aiSidebarEnabled = useAiSidebarEnabled(setIsCopilotOpen)
  useEffect(() => {
    rememberAiPanelOpen('ai-sheets-show-ai', isCopilotOpen)
  }, [isCopilotOpen])
  // the "show the AI sidebar" setting off hides the copilot rail (soft hide):
  // an explicit AI action (ribbon run, toggle) still opens the panel, and
  // closing it returns to hidden — no dead affordances, no invisible runs
  const [showFormatCells, setShowFormatCells] = useState(false)
  const [axisSizeTarget, setAxisSizeTarget] = useState<'row' | 'col' | null>(null)
  const [showLinkDialog, setShowLinkDialog] = useState(false)
  const [showSortDialog, setShowSortDialog] = useState(false)
  const [showDedupeDialog, setShowDedupeDialog] = useState(false)
  const [showCreateTable, setShowCreateTable] = useState(false)
  const [showResizeTable, setShowResizeTable] = useState(false)
  const [showNameManager, setShowNameManager] = useState(false)
  const [showPivotDialog, setShowPivotDialog] = useState(false)
  const [pivotPane, setPivotPane] = useState<PivotPaneState | null>(null)
  /// Pane the user closed for this pivot; reopens once the selection leaves it.
  const dismissedPivotRef = useRef<string | null>(null)
  const openPivotPane = (): void => {
    const seed = onGetPivotEditSeed()
    if (!seed) return
    dismissedPivotRef.current = null
    setPivotPane({ seed, model: seed.model, deferred: false, dirty: false, error: null })
  }
  const applyPivotModel = (pane: PivotPaneState, model: PivotLayoutModel): PivotPaneState => {
    const invalid = validatePivotModel(model)
    if (invalid !== null)
      return { ...pane, model, dirty: false, error: t(PIVOT_MODEL_ERRORS[invalid]) }
    const failure = onEditPivot(modelToConfig(model, pane.seed.sourceRange, pane.seed.targetCell))
    return { ...pane, model, dirty: false, error: failure }
  }
  useEffect(() => {
    if (pivotSelectionKey === null) {
      // Leaving the pivot hides the pane unless deferred edits would be lost,
      // and a dismissed pane may open again on the next visit.
      dismissedPivotRef.current = null
      setPivotPane((current) => (current && !(current.deferred && current.dirty) ? null : current))
      return
    }
    if (pivotPane?.seed.key === pivotSelectionKey) return
    // Another pivot: the pane always follows the cursor so edits never land
    // on the previous table; a missing seed leaves it closed.
    setPivotPane(null)
    if (dismissedPivotRef.current !== pivotSelectionKey) openPivotPane()
    // Only the pivot identity drives this; the callbacks are App-scoped closures.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [pivotSelectionKey])
  /** null = closed; otherwise the catalog category to open on and the cell the formula lands in */
  const [insertFunction, setInsertFunction] = useState<{
    category: string
    target: FunctionTarget
  } | null>(null)
  const insertFunctionOpen = insertFunction !== null
  const liveFunctions = useMemo(
    () => (insertFunctionOpen ? onListFunctions() : []),
    [insertFunctionOpen],
  )
  const insertFunctionOpenRef = useRef(insertFunctionOpen)
  insertFunctionOpenRef.current = insertFunctionOpen
  const openInsertFunction = (category: string): void => {
    // The modeless Function Arguments card keeps its cell: a grid pick moved
    // the active cell, so a second fx click must not re-target the session.
    if (insertFunctionOpenRef.current) return
    void functionArguments.readTarget().then((target) => {
      if (target && !insertFunctionOpenRef.current) setInsertFunction({ category, target })
    })
  }
  const [showSubtotalDialog, setShowSubtotalDialog] = useState(false)
  const [showGoalSeek, setShowGoalSeek] = useState(false)
  const [showFillSeries, setShowFillSeries] = useState(false)
  const [showConsolidateDialog, setShowConsolidateDialog] = useState(false)
  const [showGoTo, setShowGoTo] = useState(false)
  const [pasteSpecialSource, setPasteSpecialSource] = useState<PasteSpecialSource | null>(null)
  const [showZoomDialog, setShowZoomDialog] = useState(false)
  const [statsMenuAt, setStatsMenuAt] = useState<{ x: number; y: number } | null>(null)
  const openStatsMenu = (event: { clientX: number; clientY: number; preventDefault(): void }) => {
    event.preventDefault()
    setStatsMenuAt({ x: event.clientX, y: event.clientY })
  }
  useEffect(() => {
    // Univer owns the footer that hosts the sheet tabs and the statistics;
    // right-clicking anywhere in it outside the tab strip opens our menu.
    // Everything else the decision covers is sheet-context-menu.ts's: the
    // grid right-click must cancel the DOM event or Electron's native menu
    // stacks on top of Univer's own (genoffice#1816).
    const onContextMenu = (event: MouseEvent): void => {
      const action = sheetContextMenuAction(event.target)
      if (action.statsMenu) openStatsMenu(event)
      else if (action.suppressNativeMenu) event.preventDefault()
    }
    document.addEventListener('contextmenu', onContextMenu)
    return () => document.removeEventListener('contextmenu', onContextMenu)
  }, [])
  const [showHeaderFooter, setShowHeaderFooter] = useState(false)
  const [printRevision, setPrintRevision] = useState(0)
  const [showAllowEditRanges, setShowAllowEditRanges] = useState(false)
  const [protectSheetMode, setProtectSheetMode] = useState<'protect' | 'unprotect' | null>(null)
  /// Non-null while the Chart Design → Add Chart Element text prompt is open.
  const [chartTextTarget, setChartTextTarget] = useState<ChartTextTarget | null>(null)
  const onCommandRef = useRef(onCommand)
  const onIsCellEditingRef = useRef(onIsCellEditing)
  const openingWorkbookRef = useRef(openingWorkbook)
  onCommandRef.current = onCommand
  onIsCellEditingRef.current = onIsCellEditing
  openingWorkbookRef.current = openingWorkbook
  const dispatchCommand = (command: string): void => {
    if (command === 'format-cells') setShowFormatCells(true)
    else if (command === 'row-height-open') setAxisSizeTarget('row')
    else if (command === 'col-width-open') setAxisSizeTarget('col')
    else if (command === 'link-open') setShowLinkDialog(true)
    else if (command === 'sort-custom-open') setShowSortDialog(true)
    else if (command === 'remove-duplicates-open') setShowDedupeDialog(true)
    else if (command === 'table-create-open') setShowCreateTable(true)
    else if (command === 'table-resize-open') setShowResizeTable(true)
    else if (command === 'name-manager-open') setShowNameManager(true)
    else if (command === 'pivot-open') setShowPivotDialog(true)
    else if (command === 'pivot-edit') openPivotPane()
    else if (command === 'insert-function-open') openInsertFunction('All')
    else if (command.startsWith('insert-function-open:'))
      openInsertFunction(command.slice('insert-function-open:'.length))
    else if (command === 'goal-seek-open') setShowGoalSeek(true)
    else if (command === 'fill-series-open') setShowFillSeries(true)
    else if (command === 'subtotal-open') setShowSubtotalDialog(true)
    else if (command === 'consolidate-open') setShowConsolidateDialog(true)
    else if (command === 'goto-open') setShowGoTo(true)
    else if (command === PASTE_SPECIAL_OPEN_SHELL_COMMAND) {
      void onReadPasteSpecialSource().then((source) => setPasteSpecialSource(source))
    } else if (command === 'header-footer-open') setShowHeaderFooter(true)
    else if (command === 'allow-edit-ranges-open') setShowAllowEditRanges(true)
    else if (command === 'sheet-protect') {
      const protection = onGetSheetProtection()
      if (!protection) onCommand('sheet-protect')
      else if (!protection.protected) setProtectSheetMode('protect')
      else if (protection.hasPassword) setProtectSheetMode('unprotect')
      else onCommand('sheet-unprotect')
    } else if (command === 'ai-open-panel') setIsCopilotOpen(true)
    else if (command === 'ai-toggle-panel') setIsCopilotOpen((v) => !v)
    else if (command === 'chart-element-title') setChartTextTarget('title')
    else if (command === 'chart-element-axis-cat') setChartTextTarget('axis-category')
    else if (command === 'chart-element-axis-val') setChartTextTarget('axis-value')
    else if (command === 'formula-bar-toggle') toggleFormulaBarExpand()
    else onCommand(command)
  }
  const dispatchCommandRef = useRef(dispatchCommand)
  dispatchCommandRef.current = dispatchCommand
  useEffect(() => {
    setContextMenuCommandSink((command) => dispatchCommandRef.current(command))
    const unregisterFx = registerFxButtonOpener(() =>
      dispatchCommandRef.current('insert-function-open'),
    )
    const unregisterPasteSpecial = registerPasteSpecialOpener(() =>
      dispatchCommandRef.current(PASTE_SPECIAL_OPEN_SHELL_COMMAND),
    )
    return () => {
      setContextMenuCommandSink(null)
      unregisterPasteSpecial()
      unregisterFx()
    }
  }, [])
  useEffect(() => {
    // Table + gating live in global-shortcuts.ts; text fields (AI chat,
    // dialogs, Univer's find/replace and formula bar) never trigger these.
    const onKeyDown = (event: KeyboardEvent): void => {
      if (openingWorkbookRef.current) return
      const action = resolveGlobalShortcut(event, {
        isMac: IS_MAC,
        modalOpen: isModalOpen(),
        cellEditing: onIsCellEditingRef.current(),
        gridTarget: isGridKeyTarget(event.target),
        formulaBarTarget: isFormulaBarKeyTarget(event.target),
      })
      if (!action) return
      event.preventDefault()
      if (action.kind === 'dialog') {
        if (action.dialog === 'formatCells') setShowFormatCells(true)
        else setShowGoTo(true)
        return
      }
      dispatchCommandRef.current(action.command)
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [])
  // Univer's shortcut dispatcher captures keydown and binds Backspace to
  // "delete-and-start-editing" (active cell only). Register on window
  // capture *here* (child effect runs before App creates Univer) so we
  // win the race and clear the whole selection instead. Keep the listener
  // mounted once: re-binding after Univer starts would lose capture order.
  useEffect(() => {
    const onKeyDownCapture = (event: KeyboardEvent): void => {
      if (openingWorkbookRef.current) return
      if (!shouldInterceptClearSelection(event, onIsCellEditingRef.current())) return
      event.preventDefault()
      event.stopImmediatePropagation()
      onCommandRef.current('clear-contents')
    }
    window.addEventListener('keydown', onKeyDownCapture, true)
    return () => window.removeEventListener('keydown', onKeyDownCapture, true)
  }, [])
  // Deselecting while on the contextual tab lands back on Home.
  useEffect(() => {
    if (!selectedChart && activeTab === 'Chart Design') setActiveTab('Home')
  }, [selectedChart, activeTab])
  // Leaving a table while on its tab lands back on Home.
  useEffect(() => {
    if (!selectedTable && activeTab === 'Table Design') setActiveTab('Home')
  }, [selectedTable, activeTab])
  useEffect(() => {
    if (
      (activeTab === 'Shape Format' && selectedVisualKind !== 'shape') ||
      (activeTab === 'Picture Format' && selectedVisualKind !== 'image')
    ) {
      setActiveTab('Home')
    }
  }, [selectedVisualKind, activeTab])
  // A freshly created table activates its tab (Excel). The create may finish
  // asynchronously (streamed rows load first), so the signal is a counter,
  // consumed only once the selection echo has caught up with the new table.
  const seenTableCreatedSeq = useRef(tableCreatedSeq)
  useEffect(() => {
    if (tableCreatedSeq === seenTableCreatedSeq.current || !selectedTable) return
    seenTableCreatedSeq.current = tableCreatedSeq
    setActiveTab('Table Design')
  }, [tableCreatedSeq, selectedTable])
  // Univer's formula-bar Name Box (the defined-name selector) is the only
  // cell-reference box; the Go To ▾ arrow rides inside the bar next to it.
  // The bar is Univer-owned DOM that can remount with the workbench, so the
  // button is (re)inserted on mutation rather than rendered by React.
  const gotoTip = t('appGoToButtonTitle')
  const gotoTipRef = useRef(gotoTip)
  gotoTipRef.current = gotoTip
  useEffect(() => {
    const button = document.querySelector<HTMLButtonElement>('.goto-in-bar')
    if (button) button.dataset['tip'] = gotoTip
  }, [gotoTip])
  useEffect(() => {
    const ensure = (): void => {
      // The Name Box sits in a fixed-width block wrapper; the flex row is the
      // bar itself, so the button must ride as the wrapper's sibling.
      const wrapper = document.querySelector('[data-u-comp="defined-name"]')?.parentElement
      const bar = wrapper?.parentElement
      if (!wrapper || !bar) return
      let button = bar.querySelector<HTMLButtonElement>('.goto-in-bar')
      if (!button) {
        button = document.createElement('button')
        button.type = 'button'
        button.className = 'goto-in-bar'
        button.textContent = '▾'
        button.setAttribute('aria-label', 'Go To')
        button.addEventListener('click', () => setShowGoTo(true))
        wrapper.after(button)
      }
      if (button.dataset['tip'] !== gotoTipRef.current) button.dataset['tip'] = gotoTipRef.current
    }
    ensure()
    const container = document.getElementById('univer-container')
    if (!container) return undefined
    const observer = new MutationObserver(ensure)
    observer.observe(container, { childList: true, subtree: true })
    return () => {
      observer.disconnect()
      document.querySelector('.goto-in-bar')?.remove()
    }
  }, [])
  const visibleTabs: readonly RibbonTab[] = [
    ...ribbonTabs,
    ...(selectedChart ? (['Chart Design'] as const) : []),
    ...(selectedTable ? (['Table Design'] as const) : []),
    ...(selectedVisualKind === 'shape' ? (['Shape Format'] as const) : []),
    ...(selectedVisualKind === 'image' ? (['Picture Format'] as const) : []),
  ]
  const saveAsTitle = `${t('appSaveAs')} (${platformShortcuts('⇧⌘S')})`

  return (
    <main
      className={`app-shell ${isCopilotOpen ? '' : 'copilot-collapsed'}${
        !isCopilotOpen && !aiSidebarEnabled ? ' copilot-off' : ''
      }`}
      inert={openingWorkbook}
      aria-busy={openingWorkbook}
    >
      <header className={`excel-header ${collapse.rootClass}`} ref={collapse.rootRef}>
        <nav
          className={`ribbon-tabs ${IN_TAB ? '' : IS_MAC ? 'ribbon-tabs-mac' : 'ribbon-tabs-win'}`}
          aria-label={t('appScopeWorkbook')}
          onDoubleClick={collapse.onTabsDoubleClick}
        >
          <button
            type="button"
            className="qa-btn"
            data-tip={t('appSaveTitle')}
            aria-label={t('appSaveTitle')}
            disabled={!canSave}
            onClick={onSave}
          >
            <SaveIcon />
          </button>
          <button
            type="button"
            className="qa-btn"
            data-tip={saveAsTitle}
            aria-label={saveAsTitle}
            disabled={!canSaveAs}
            onClick={onSaveAs}
          >
            <SaveAsIcon />
          </button>
          <button
            type="button"
            className="qa-btn"
            data-tip={t('appUndo')}
            aria-label={t('appUndo')}
            disabled={!canUndo}
            onClick={() => onUndo()}
          >
            <UndoIcon />
          </button>
          <button
            type="button"
            className="qa-btn"
            data-tip={t('appRedo')}
            aria-label={t('appRedo')}
            disabled={!canRedo}
            onClick={onRedo}
          >
            <RedoIcon />
          </button>
          <label
            className={`autosave-toggle ${autoSave ? 'on' : ''}`}
            data-tip={t('appAutoSaveTip')}
          >
            <span className="autosave-knob" />
            <span className="autosave-text">{t('appAutoSave')}</span>
            <input
              type="checkbox"
              checked={autoSave}
              onChange={(e) => onAutoSaveChange(e.target.checked)}
            />
          </label>
          <span className="qa-sep" aria-hidden="true" />
          {visibleTabs.map((tab) => (
            <button
              className={`${collapse.tabClass(tab === activeTab)} ${CONTEXTUAL_TABS.includes(tab) ? 'contextual' : ''}`}
              data-tip={collapse.tabTip(tab === activeTab)}
              key={tab}
              onClick={() => {
                collapse.onTabPress(tab === activeTab)
                setActiveTab(tab)
              }}
            >
              {t(TAB_LABEL[tab])}
            </button>
          ))}
          <span className="ribbon-tabs-spacer" />
          <span className="workbook-status" role="status" aria-live="polite">
            {statusMessage}
          </span>
        </nav>

        <Ribbon
          activeTab={activeTab}
          selectionFormat={selectionFormat}
          formatPainterActive={formatPainterActive}
          sheetHasContent={sheetHasContent}
          sheetProtected={onGetSheetProtection()?.protected ?? null}
          workbookProtected={onGetWorkbookProtection()}
          formulaBarVisible={formulaBarVisible}
          crossHighlightVisible={crossHighlightVisible}
          pageLayout={pageLayout}
          selectedChart={selectedChart}
          selectedShape={selectedShape}
          selectedTable={selectedTable}
          outlineSummary={outlineSummary}
          onListNames={() => {
            // Names scoped to another sheet resolve to #NAME? here; only
            // workbook-scoped and active-sheet names are usable in a formula.
            const data = onGetDefinedNames()
            return data.names
              .filter(
                (entry) => entry.scopeSheetId === null || entry.scopeSheetId === data.activeSheetId,
              )
              .map((entry) => entry.name)
          }}
          calcManual={calcManual}
          onRefreshPivot={onRefreshPivot}
          onIsSelectionInPivot={onIsSelectionInPivot}
          onCommand={dispatchCommand}
          onAiRun={(nextPrompt) => {
            setIsCopilotOpen(true)
            onSend(nextPrompt)
          }}
          aiOpen={isCopilotOpen}
          onAiToggle={() => setIsCopilotOpen((open) => !open)}
        />
      </header>

      {/* AI panel docks on the left, full height under the ribbon (unified with docs) */}
      <div className="sheet-body">
        <AiChatPanel
          isOpen={isCopilotOpen}
          hideRail={!aiSidebarEnabled}
          hasContent={sheetHasContent}
          chat={chat}
          {...(historicChat !== undefined ? { historicChat } : {})}
          attachments={attachments}
          attachedSkills={attachedSkills}
          onToggleSkill={onToggleSkill}
          attachNotice={attachNotice}
          onPickAttachments={onPickAttachments}
          onAddAttachmentPaths={onAddAttachmentPaths}
          onAddPastedImage={onAddPastedImage}
          onRemoveAttachment={onRemoveAttachment}
          prompt={prompt}
          preview={preview}
          aiBusy={aiBusy}
          sessionUsage={sessionUsage}
          onPromptChange={onPromptChange}
          onSend={onSend}
          onSendQueued={(text, atts) => {
            if (aiBusy) return false
            onSend(text, atts)
            return true
          }}
          onStop={onStop}
          onNewChat={onNewChat}
          onUndo={onUndo}
          scopeRange={aiScopeRange}
          scopeColumns={aiScopeColumns}
          scopeLocked={aiScopeLocked}
          onScopeDismiss={onAiScopeDismiss}
          onCitation={onAiCitation}
          onExpand={() => setIsCopilotOpen(true)}
          onCollapse={() => setIsCopilotOpen(false)}
        />
        <div className="sheet-main">
          <section className="workbook-area">
            <div id="univer-container" className="spreadsheet" />
            {emptyCsvNotice && (
              <div className="csv-empty-notice" role="status">
                <span>{t('appCsvEmpty')}</span>
                <button type="button" onClick={onOpenWorkbook}>
                  {t('appOpenWorkbookTitle')}
                </button>
                <button
                  type="button"
                  className="csv-empty-notice-close"
                  aria-label={t('appClose')}
                  onClick={onDismissEmptyCsvNotice}
                >
                  ×
                </button>
              </div>
            )}
          </section>
          {aiSelectionAskAnchor && aiScopeRange && !aiBusy && (
            <AiSelectionAsk
              anchor={aiSelectionAskAnchor}
              range={aiScopeRange}
              onDismiss={onAiSelectionAskDismiss}
              onSend={(instruction) => {
                setIsCopilotOpen(true)
                onSend(instruction)
              }}
            />
          )}

          {/* Status bar spans the sheet column only — the AI dock keeps the full window height (unified with docs/slides). */}
          <footer className="status-bar" onContextMenu={openStatsMenu}>
            <div className="status-left">
              <span className="status-msg">{statusMessage}</span>
            </div>
            <div className="status-right">
              {/* One-of-N view switcher, so it is a radiogroup rather than a
                  second set of buttons named like the View tab's. */}
              <div
                className="view-switch"
                role="radiogroup"
                aria-label={t('appGroupWorkbookViews')}
              >
                <button
                  className="view-btn"
                  role="radio"
                  data-tip={t('appNormalViewTip')}
                  aria-label={t('appNormalViewTip')}
                  aria-checked={pageLayout.pageBreakPreview !== true}
                  onClick={() => {
                    if (pageLayout.pageBreakPreview === true) onCommand('toggle-page-break-preview')
                  }}
                >
                  ▦
                </button>
                <button
                  className="view-btn"
                  role="radio"
                  data-tip={t('appPageBreakPreviewTip')}
                  aria-label={t('appPageBreakPreviewTip')}
                  aria-checked={pageLayout.pageBreakPreview === true}
                  onClick={() => {
                    if (pageLayout.pageBreakPreview !== true) onCommand('toggle-page-break-preview')
                  }}
                >
                  ┆
                </button>
              </div>
              <button
                className="zoom-btn"
                data-tip={t('appZoomOut')}
                aria-label={t('appZoomOut')}
                onClick={() => onCommand('zoom-out')}
              >
                −
              </button>
              <input
                className="zoom-slider"
                type="range"
                min={SHEET_ZOOM_MIN}
                max={SHEET_ZOOM_MAX}
                value={clampZoomPercent(zoomPercent)}
                onChange={(event) => onCommand(`zoom:${event.target.value}`)}
              />
              <button
                className="zoom-btn"
                data-tip={t('appZoomIn')}
                aria-label={t('appZoomIn')}
                onClick={() => onCommand('zoom-in')}
              >
                +
              </button>
              <button
                className="zoom-value"
                data-tip={t('appZoomLevel')}
                aria-label={t('appZoomLevel')}
                onClick={() => setShowZoomDialog(true)}
              >
                {zoomPercent}%
              </button>
            </div>
          </footer>
          {statsMenuAt && (
            <StatusBarStatsMenu
              at={statsMenuAt}
              enabled={statusBarFuncs}
              onToggle={onToggleStatusBarFunc}
              onClose={() => setStatsMenuAt(null)}
            />
          )}
        </div>
      </div>
      {showFormatCells && (
        <FormatCellsDialog
          selectionFormat={selectionFormat}
          anchorValue={onGetAnchorValue()}
          onCommand={onCommand}
          onClose={() => setShowFormatCells(false)}
        />
      )}
      {axisSizeTarget && (
        <AxisSizeDialog
          axis={axisSizeTarget}
          onCommand={onCommand}
          onClose={() => setAxisSizeTarget(null)}
        />
      )}
      {showLinkDialog && (
        <LinkDialog
          currentTarget={selectionFormat?.link ?? null}
          onCommand={onCommand}
          onClose={() => setShowLinkDialog(false)}
        />
      )}
      {chartTextTarget && selectedChart && (
        <ChartTextDialog
          target={chartTextTarget}
          initial={
            chartTextTarget === 'title'
              ? selectedChart.title
              : ((chartTextTarget === 'axis-category'
                  ? selectedChart.axisTitles?.category
                  : selectedChart.axisTitles?.value) ?? '')
          }
          onCommand={onCommand}
          onClose={() => setChartTextTarget(null)}
        />
      )}
      {showSortDialog && (
        <SortDialog
          columns={onGetSortColumns()}
          onCommand={onCommand}
          onClose={() => setShowSortDialog(false)}
        />
      )}
      {showDedupeDialog && (
        <RemoveDuplicatesDialog onCommand={onCommand} onClose={() => setShowDedupeDialog(false)} />
      )}
      {showCreateTable && (
        <TableRangeDialog
          title={t('appCreateTable')}
          note={t('appCreateTableWhere')}
          initialRange={onGetCreateTableRange()}
          headersOption
          onSubmit={(range, headers) => dispatchCommand(`table-create:${range}:${headers ? 1 : 0}`)}
          onClose={() => setShowCreateTable(false)}
        />
      )}
      {showResizeTable && selectedTable && (
        <TableRangeDialog
          title={t('appResizeTable')}
          note={t('appResizeTableNote')}
          initialRange={selectedTable.rangeA1}
          onSubmit={(range) => onCommand(`table-resize:${range}`)}
          onClose={() => setShowResizeTable(false)}
        />
      )}
      {showNameManager &&
        (() => {
          const data = onGetDefinedNames()
          return (
            <NameManagerDialog
              names={data.names}
              sheets={data.sheets}
              onAction={onDefinedNameAction}
              onClose={() => setShowNameManager(false)}
            />
          )
        })()}
      {showPivotDialog &&
        (() => {
          const sourceRange = onGetSourceRange()
          return (
            <PivotDialog
              hasFields={onGetPivotFields(sourceRange).length > 0}
              sourceRange={sourceRange}
              onCreate={(request) => {
                const failure = onCreatePivot(request)
                if (failure === null) openPivotPane()
                return failure
              }}
              onClose={() => setShowPivotDialog(false)}
            />
          )
        })()}
      {pivotPane && (
        <PivotFieldPane
          fields={pivotPane.seed.fields}
          model={pivotPane.model}
          deferred={pivotPane.deferred}
          dirty={pivotPane.dirty}
          error={pivotPane.error}
          onModelChange={(model) =>
            setPivotPane(
              pivotPane.deferred
                ? { ...pivotPane, model, dirty: true, error: null }
                : applyPivotModel(pivotPane, model),
            )
          }
          onDeferredChange={(deferred) =>
            setPivotPane(
              !deferred && pivotPane.dirty
                ? { ...applyPivotModel(pivotPane, pivotPane.model), deferred }
                : { ...pivotPane, deferred },
            )
          }
          onUpdate={() => setPivotPane(applyPivotModel(pivotPane, pivotPane.model))}
          onGetFieldMembers={onGetPivotFieldMembers}
          onClose={() => {
            dismissedPivotRef.current = pivotPane.seed.key
            setPivotPane(null)
          }}
        />
      )}
      {showFillSeries && (
        <FillSeriesDialog
          context={onGetFillSeriesContext()}
          defaultType={fillSeriesDefaultType(selectionFormat?.numberFormat)}
          onApply={onFillSeries}
          onClose={() => setShowFillSeries(false)}
        />
      )}
      {showGoalSeek && (
        <GoalSeekDialog
          initialSetCell={onGetActiveCell()}
          onSolve={onGoalSeek}
          onClose={() => setShowGoalSeek(false)}
        />
      )}
      {insertFunction !== null && (
        <InsertFunctionFlow
          functions={liveFunctions}
          target={insertFunction.target}
          initialCategory={insertFunction.category}
          host={functionArguments}
          onClose={() => setInsertFunction(null)}
        />
      )}
      {showSubtotalDialog && (
        <SubtotalDialog
          fields={onGetPivotFields()}
          onCreate={onCreateSubtotal}
          onClose={() => setShowSubtotalDialog(false)}
        />
      )}
      {showConsolidateDialog && (
        <ConsolidateDialog
          defaultReference={onGetConsolidateDefault()}
          targetLabel={onGetActiveCell()}
          onCreate={onCreateConsolidate}
          onClose={() => setShowConsolidateDialog(false)}
        />
      )}
      {showZoomDialog && (
        <ZoomDialog
          zoomPercent={zoomPercent}
          onCommand={onCommand}
          onClose={() => setShowZoomDialog(false)}
        />
      )}
      {pasteSpecialSource && (
        <PasteSpecialDialog
          source={pasteSpecialSource}
          onApply={(options) => onPasteSpecial(pasteSpecialSource, options)}
          onClose={() => setPasteSpecialSource(null)}
        />
      )}
      {showGoTo && (
        <GoToDialog
          names={onListDefinedNames()}
          onGo={onGoToReference}
          onClose={() => setShowGoTo(false)}
        />
      )}
      {protectSheetMode && (
        <ProtectSheetDialog
          mode={protectSheetMode}
          onProtect={onProtectSheet}
          onUnprotect={onUnprotectSheet}
          onClose={() => setProtectSheetMode(null)}
        />
      )}
      {showAllowEditRanges &&
        (() => {
          const snapshot = onGetProtectedRanges()
          return (
            <AllowEditRangesDialog
              ranges={snapshot.error === null ? snapshot.ranges : []}
              defaultRef={activeCellA1}
              onApply={(ranges) => snapshot.error ?? onApplyProtectedRanges(ranges)}
              onClose={() => setShowAllowEditRanges(false)}
            />
          )
        })()}
      {printHost && (
        <PrintDialog
          host={printHost}
          onClose={onClosePrint}
          onPageSetup={() => setShowHeaderFooter(true)}
          revision={printRevision}
        />
      )}
      {showHeaderFooter && (
        <HeaderFooterDialog
          initialHeader={pageLayout.header ?? null}
          initialFooter={pageLayout.footer ?? null}
          onApply={onApplyHeaderFooter}
          onClose={() => {
            setShowHeaderFooter(false)
            setPrintRevision((value) => value + 1)
          }}
        />
      )}
      <SheetTabContextMenu
        actions={sheetTabActions}
        getSheetProtected={() => onGetSheetProtection()?.protected ?? null}
        getWorkbookLocked={onGetWorkbookProtection}
        onProtectSheet={() => dispatchCommand('sheet-protect')}
      />
    </main>
  )
}

const NO_SORT_LEVEL = -1

function SortDialog({
  columns,
  onCommand,
  onClose,
}: {
  readonly columns: { label: string; colIndex: number }[]
  readonly onCommand: (command: string) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [hasHeader, setHasHeader] = useState(true)
  const first = columns[0]?.colIndex ?? NO_SORT_LEVEL
  const [levels, setLevels] = useState<{ colIndex: number; order: 'a' | 'd' }[]>([
    { colIndex: first, order: 'a' },
    { colIndex: NO_SORT_LEVEL, order: 'a' },
    { colIndex: NO_SORT_LEVEL, order: 'a' },
  ])
  const apply = (): void => {
    const rules = levels
      .filter((level) => level.colIndex !== NO_SORT_LEVEL)
      .map((level) => `${level.colIndex}${level.order}`)
    if (rules.length > 0) onCommand(`sort-custom:${hasHeader ? 1 : 0}:${rules.join(',')}`)
    onClose()
  }
  const setLevel = (
    index: number,
    patch: Partial<{ colIndex: number; order: 'a' | 'd' }>,
  ): void => {
    setLevels((previous) =>
      previous.map((level, at) => (at === index ? { ...level, ...patch } : level)),
    )
  }
  const modal = useModalDialog(onClose)
  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog sort-dialog"
        role="dialog"
        {...modal}
        aria-label={t('appCustomSort')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('appSort')}</header>
        <div className="dialog-body">
          <label className="sort-header-check">
            <input
              type="checkbox"
              checked={hasHeader}
              onChange={(event) => setHasHeader(event.target.checked)}
            />
            {t('appMyDataHasHeaders')}
          </label>
          {levels.map((level, index) => (
            <div className="sort-level" key={index}>
              <span>{t(index === 0 ? 'appSortBy' : 'appThenBy')}</span>
              <Dropdown
                ariaLabel={t(index === 0 ? 'appSortBy' : 'appThenBy')}
                value={String(level.colIndex)}
                options={[
                  ...(index > 0 ? [{ value: String(NO_SORT_LEVEL), label: t('appSortNone') }] : []),
                  ...columns.map((column) => ({
                    value: String(column.colIndex),
                    label: column.label,
                  })),
                ]}
                onPick={(v) => setLevel(index, { colIndex: Number(v) })}
              />
              <Dropdown
                value={level.order}
                options={[
                  { value: 'a', label: t('appSortAsc') },
                  { value: 'd', label: t('appSortDesc') },
                ]}
                onPick={(v) => setLevel(index, { order: v })}
              />
            </div>
          ))}
        </div>
        <footer className="dialog-actions">
          <button onClick={onClose}>{t('appCancel')}</button>
          <button
            className="primary"
            disabled={levels.every((level) => level.colIndex === NO_SORT_LEVEL)}
            onClick={apply}
          >
            {t('appOk')}
          </button>
        </footer>
      </div>
    </div>
  )
}

function RemoveDuplicatesDialog({
  onCommand,
  onClose,
}: {
  readonly onCommand: (command: string) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [hasHeader, setHasHeader] = useState(true)
  const modal = useModalDialog(onClose)
  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog sort-dialog"
        role="dialog"
        {...modal}
        aria-label={t('appRemoveDuplicates')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t('appRemoveDuplicates')}</header>
        <div className="dialog-body">
          <p className="dialog-note">{t('appRemoveDuplicatesNote')}</p>
          <label className="sort-header-check">
            <input
              type="checkbox"
              checked={hasHeader}
              onChange={(event) => setHasHeader(event.target.checked)}
            />
            {t('appMyDataHasHeaders')}
          </label>
        </div>
        <footer className="dialog-actions">
          <button onClick={onClose}>{t('appCancel')}</button>
          <button
            className="primary"
            onClick={() => {
              onCommand(`remove-duplicates:${hasHeader ? 1 : 0}`)
              onClose()
            }}
          >
            {t('appOk')}
          </button>
        </footer>
      </div>
    </div>
  )
}

/// Create Table (Ctrl+T) / Resize Table: one range box, optionally the
/// "My table has headers" check.
function TableRangeDialog({
  title,
  note,
  initialRange,
  headersOption = false,
  onSubmit,
  onClose,
}: {
  readonly title: string
  readonly note: string
  readonly initialRange: string
  readonly headersOption?: boolean
  readonly onSubmit: (range: string, headers: boolean) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [range, setRange] = useState(initialRange)
  const [headers, setHeaders] = useState(true)
  const modal = useModalDialog(onClose)
  const submit = (): void => {
    if (range.trim().length === 0) return
    onSubmit(range.trim().toUpperCase(), headers)
    onClose()
  }
  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog sort-dialog"
        role="dialog"
        {...modal}
        aria-label={title}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{title}</header>
        <div className="dialog-body">
          <p className="dialog-note">{note}</p>
          <input
            className="table-range-input"
            autoFocus
            value={range}
            spellCheck={false}
            onChange={(event) => setRange(event.target.value)}
            onKeyDown={(event) => {
              if (event.key === 'Enter') submit()
            }}
          />
          {headersOption && (
            <label className="sort-header-check">
              <input
                type="checkbox"
                checked={headers}
                onChange={(event) => setHeaders(event.target.checked)}
              />
              {t('appMyTableHasHeaders')}
            </label>
          )}
        </div>
        <footer className="dialog-actions">
          <button onClick={onClose}>{t('appCancel')}</button>
          <button className="primary" onClick={submit}>
            {t('appOk')}
          </button>
        </footer>
      </div>
    </div>
  )
}

/// Chart Design → Add Chart Element text prompt (title / axis titles).
/// An empty axis title removes that axis title; an empty chart title blanks it.
function ChartTextDialog({
  target,
  initial,
  onCommand,
  onClose,
}: {
  readonly target: ChartTextTarget
  readonly initial: string
  readonly onCommand: (command: string) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [value, setValue] = useState(initial)
  const { heading, command } = CHART_TEXT_LABELS[target]
  const apply = (): void => {
    onCommand(`${command}:${value.trim()}`)
    onClose()
  }
  const modal = useModalDialog(onClose)
  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog link-dialog"
        role="dialog"
        {...modal}
        aria-label={t(heading)}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t(heading)}</header>
        <div className="dialog-body">
          <label>
            {t(target === 'title' ? 'appTitleText' : 'appAxisTitleText')}
            <input
              autoFocus
              type="text"
              value={value}
              maxLength={255}
              onChange={(event) => setValue(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === 'Enter') apply()
              }}
            />
          </label>
        </div>
        <footer className="dialog-actions">
          <button onClick={onClose}>{t('appCancel')}</button>
          <button className="primary" onClick={apply}>
            {t('appOk')}
          </button>
        </footer>
      </div>
    </div>
  )
}

/// Numeric row-height / column-width entry (Excel's Format → Row Height /
/// Column Width). Applies to the rows/columns of the current selection.
function AxisSizeDialog({
  axis,
  onCommand,
  onClose,
}: {
  readonly axis: 'row' | 'col'
  readonly onCommand: (command: string) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [value, setValue] = useState('')
  const max = axis === 'row' ? 409.5 : 255
  const parsed = Number(value.trim().replace(',', '.'))
  const valid = value.trim() !== '' && Number.isFinite(parsed) && parsed >= 0 && parsed <= max
  const apply = (): void => {
    if (!valid) return
    onCommand(`${axis === 'row' ? 'row-height' : 'col-width'}:${parsed}`)
    onClose()
  }
  const modal = useModalDialog(onClose)
  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog link-dialog"
        role="dialog"
        {...modal}
        aria-label={t(axis === 'row' ? 'appRowHeight' : 'appColWidth')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t(axis === 'row' ? 'appRowHeight' : 'appColWidth')}</header>
        <div className="dialog-body">
          <label>
            {t(axis === 'row' ? 'appRowHeightLabel' : 'appColWidthLabel')}
            <input
              autoFocus
              type="text"
              inputMode="decimal"
              value={value}
              onChange={(event) => setValue(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === 'Enter') apply()
              }}
            />
          </label>
        </div>
        <footer className="dialog-actions">
          <button onClick={onClose}>{t('appCancel')}</button>
          <button className="primary" disabled={!valid} onClick={apply}>
            {t('appOk')}
          </button>
        </footer>
      </div>
    </div>
  )
}

function LinkDialog({
  currentTarget,
  onCommand,
  onClose,
}: {
  readonly currentTarget: string | null
  readonly onCommand: (command: string) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [value, setValue] = useState(currentTarget ?? '')
  const apply = (): void => {
    if (value.trim()) onCommand(`link-set:${encodeURIComponent(value)}`)
    onClose()
  }
  const modal = useModalDialog(onClose)
  return (
    <div className="dialog-backdrop" onClick={onClose}>
      <div
        className="format-cells-dialog link-dialog"
        role="dialog"
        {...modal}
        aria-label={t(currentTarget ? 'appEditLinkTitle' : 'appInsertLinkTitle')}
        onClick={(event) => event.stopPropagation()}
      >
        <header>{t(currentTarget ? 'appEditLinkTitle' : 'appInsertLinkTitle')}</header>
        <div className="dialog-body">
          <label>
            {t('appLinkAddressLabel')}
            <input
              autoFocus
              type="text"
              value={value}
              placeholder={t('appLinkPlaceholder')}
              onChange={(event) => setValue(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === 'Enter') apply()
              }}
            />
          </label>
        </div>
        <footer className="dialog-actions">
          {currentTarget && (
            <button
              onClick={() => {
                onCommand('link-remove')
                onClose()
              }}
            >
              {t('appRemoveLink')}
            </button>
          )}
          <button onClick={onClose}>{t('appCancel')}</button>
          <button className="primary" disabled={!value.trim()} onClick={apply}>
            {t('appOk')}
          </button>
        </footer>
      </div>
    </div>
  )
}

const STAT_LABEL: Record<StatusBarFunc, StringKey> = {
  AVERAGE: 'appStatAverage',
  COUNTA: 'appStatCount',
  COUNT: 'appStatNumericalCount',
  MIN: 'appStatMin',
  MAX: 'appStatMax',
  SUM: 'appStatSum',
}

/// Excel's status bar right-click menu: tick the aggregates the footer shows.
function StatusBarStatsMenu({
  at,
  enabled,
  onToggle,
  onClose,
}: {
  readonly at: { x: number; y: number }
  readonly enabled: readonly StatusBarFunc[]
  readonly onToggle: (func: StatusBarFunc) => void
  readonly onClose: () => void
}): React.JSX.Element {
  const { t } = useI18n()
  const ref = useRef<HTMLDivElement>(null)
  useEscapeClose(true, onClose)
  useEffect(() => {
    const onPress = (event: MouseEvent): void => {
      if (!ref.current?.contains(event.target as Node)) onClose()
    }
    window.addEventListener('mousedown', onPress)
    return () => window.removeEventListener('mousedown', onPress)
  }, [onClose])
  // Flip up when the cursor sits near the bottom: the status bar always does.
  const style: React.CSSProperties = {
    left: Math.min(at.x, window.innerWidth - 200),
    bottom: Math.max(8, window.innerHeight - at.y),
  }
  return (
    <div ref={ref} className="status-stats-menu" role="menu" style={style}>
      {STATUS_BAR_FUNCS.map((func) => (
        <button
          key={func}
          type="button"
          role="menuitemcheckbox"
          aria-checked={enabled.includes(func)}
          onClick={() => onToggle(func)}
        >
          <i className="check-box">{enabled.includes(func) ? '✓' : ''}</i>
          {t(STAT_LABEL[func])}
        </button>
      ))}
    </div>
  )
}

function fillSeriesDefaultType(numberFormat: string | undefined): FillSeriesType {
  return numberFormat && numberFormatLabel(numberFormat) === 'Date' ? 'date' : 'linear'
}
export type { PageLayoutEcho, SelectedChartRibbon, SelectedShapeRibbon } from './SheetsRibbon'
