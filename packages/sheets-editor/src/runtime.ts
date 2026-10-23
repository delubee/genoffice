import {
  CommandType,
  ICommandService,
  IUndoRedoService,
  UndoCommandId,
  RedoCommandId,
  IUniverInstanceService,
  LocaleType,
  type ICellData,
  type IWorkbookData,
} from '@univerjs/core'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import enUS from '@univerjs/preset-sheets-core/locales/en-US'
import zhCN from '@univerjs/preset-sheets-core/locales/zh-CN'
import zhTW from '@univerjs/preset-sheets-core/locales/zh-TW'
import '@univerjs/preset-sheets-core/lib/index.css'
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import {
  formatAddress,
  parseAddress,
  parseRange,
  rangeCellCount,
} from '@genoffice/xlsx-gateway/domain/cell-address'
import { workbookStateSchema } from '@genoffice/xlsx-gateway/shared/session-state'
import type { CandidateState } from './changeset'
import type { EditorHost, EditorProjection, GridRange } from './ports'
import { createUniver } from './create-univer'
import { manualStyleIntent, manualValueIntents } from './manual-intents'
import { gridCell, type ViewCell } from './grid-cell'
import { requestHostHistory } from './history-bridge'
import type { SheetVisual } from '@genoffice/xlsx-gateway/domain/chart-visual'

export class WorkbookRuntime implements EditorProjection {
  private runtime: ReturnType<typeof createUniver>
  private state: CandidateState | undefined
  private revision = 0
  private internal = 0
  private closed = false
  private composing = false
  private pending: Promise<unknown> = Promise.resolve()
  private pendingError = false
  private disposables: { dispose(): void }[] = []
  private loaded = new Set<string>()
  private viewports = new Map<string, { row: number; column: number }>()
  private projectionEpoch = 0
  private projectionDisposables: { dispose(): void }[] = []
  private loadTimer: ReturnType<typeof setTimeout> | undefined
  private themeObserver: MutationObserver
  private systemTheme = matchMedia('(prefers-color-scheme: dark)')
  constructor(
    private readonly container: HTMLElement,
    private readonly host: EditorHost,
    private readonly report: (error: unknown) => void,
    private readonly changed: () => void = () => {},
    locale = navigator.language,
  ) {
    this.runtime = createUniver({
      locale: /^zh[-_](TW|HK|MO|Hant)/i.test(locale)
        ? LocaleType.ZH_TW
        : /^zh/i.test(locale)
          ? LocaleType.ZH_CN
          : LocaleType.EN_US,
      locales: { [LocaleType.EN_US]: enUS, [LocaleType.ZH_CN]: zhCN, [LocaleType.ZH_TW]: zhTW },
      darkMode: this.darkTheme(),
      presets: [
        UniverSheetsCorePreset({
          container,
          header: true,
          toolbar: false,
          contextMenu: false,
          formulaBar: true,
          footer: { sheetBar: true, statisticBar: true, menus: false, zoomSlider: false },
        }),
      ],
    })
    this.themeObserver = new MutationObserver(this.updateTheme)
    this.themeObserver.observe(document.body, {
      attributes: true,
      attributeFilter: ['class', 'data-vscode-theme-kind'],
    })
    this.systemTheme.addEventListener('change', this.updateTheme)
    const commands = this.runtime.univer.__getInjector().get(ICommandService)
    this.disposables.push(
      commands.beforeCommandExecuted((command, options) => {
        if (this.internal) return
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
          const operations = manualValueIntents(targetSheetId, {
            ...params,
            range: params.range ?? sheet.getActiveRange()?.getRange(),
          })
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
    const classes = document.body.classList
    if (classes.contains('vscode-dark') || classes.contains('vscode-high-contrast')) return true
    if (classes.contains('vscode-light') || classes.contains('vscode-high-contrast-light'))
      return false
    return this.systemTheme.matches
  }
  private updateTheme = () => {
    this.runtime.univerAPI.toggleDarkMode(this.darkTheme())
  }
  private onCompositionEnd = () => {
    this.composing = false
  }
  async install(stateId: string, revision: number, data: unknown) {
    const state = workbookStateSchema.parse(data)
    if (state.stateId !== stateId || this.closed) throw new Error('invalid_projection')
    const active = this.runtime.univerAPI.getActiveWorkbook(),
      sheetId = active?.getActiveSheet()?.getSheetId(),
      selected = active?.getActiveSheet()?.getActiveRange()?.getRange()
    this.projectionEpoch++
    for (const disposable of this.projectionDisposables) disposable.dispose()
    this.projectionDisposables = []
    this.state = state
    this.revision = revision
    this.loaded.clear()
    this.viewports.clear()
    this.internal++
    try {
      if (active)
        this.runtime.univer.__getInjector().get(IUniverInstanceService).disposeUnit(active.getId())
      const sheets: IWorkbookData['sheets'] = Object.fromEntries(
        state.sheets.map((sheet) => {
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
              hidden: (sheet.features?.originalMetadata as { hidden?: boolean } | undefined)?.hidden
                ? 1
                : 0,
              rowCount: Math.min(1_048_576, Math.max(1000, sheet.rows + 100)),
              columnCount: Math.min(16_384, Math.max(20, sheet.columns + 5)),
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
          const sheet = workbook.getActiveSheet()
          if (sheet) this.host.selection(sheet.getSheetId(), ranges[0] ?? null)
        }),
      )
    } finally {
      this.internal--
    }
    await this.loadVisible()
    this.changed()
  }
  private scheduleLoad() {
    if (this.loadTimer) clearTimeout(this.loadTimer)
    this.loadTimer = setTimeout(() => {
      void this.loadVisible().catch(this.report)
    }, 80)
  }
  private async loadVisible() {
    const epoch = this.projectionEpoch
    const state = this.state,
      sheet = this.runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
    if (!state || !sheet || this.closed) return
    const scroll = this.viewports.get(sheet.getSheetId())
    const startRow = Math.max(0, scroll?.row ?? 0),
      startColumn = Math.max(0, scroll?.column ?? 0)
    const range = {
      startRow,
      endRow: Math.min(sheet.getMaxRows() - 1, startRow + 49),
      startColumn,
      endColumn: Math.min(sheet.getMaxColumns() - 1, startColumn + 39),
    }
    const key = `${state.stateId}:${sheet.getSheetId()}:${startRow}:${startColumn}`
    if (this.loaded.has(key)) return
    let result: { cells: ViewCell[] }
    const started = Date.now()
    for (;;) {
      try {
        result = (await this.host.read(sheet.getSheetId(), range, state.stateId)) as {
          cells: ViewCell[]
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
    const cells: Record<number, Record<number, ICellData>> = {}
    for (const cell of result.cells) {
      const point = parseAddress(cell.address)
      ;(cells[point.row] ??= {})[point.column] = gridCell(cell)
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
  /** Read the installed grid projection for presentation verification. */
  projectedValues(range: string) {
    if (rangeCellCount(parseRange(range)) > 2000) throw new Error('presentation_range_limit')
    return this.runtime.univerAPI
      .getActiveWorkbook()
      ?.getActiveSheet()
      ?.getRange(range)
      .getRawValues()
  }
  get charts(): readonly SheetVisual[] {
    const sheetId = this.runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId()
    return (this.state?.sheets.find((sheet) => sheet.id === sheetId)?.features?.charts ??
      []) as unknown as SheetVisual[]
  }
  enqueue(operations: readonly WorkbookOperation[]) {
    const result = this.pending.then(() => this.submit(operations))
    this.pending = result.then(
      () => {
        this.pendingError = false
      },
      () => {
        this.pendingError = true
      },
    )
    return result
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
  async flushEditor() {
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
    if (this.closed) return
    this.closed = true
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
