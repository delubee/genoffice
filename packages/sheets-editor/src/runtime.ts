import {
  CommandType,
  ICommandService,
  IUndoRedoService,
  UndoCommandId,
  RedoCommandId,
  IUniverInstanceService,
  LocaleType,
  ThemeService,
  RANGE_TYPE,
  type ICellData,
  type IWorkbookData,
} from '@univerjs/core'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import { AddWorksheetMergeMutation } from '@univerjs/sheets'
import {
  DeltaColumnWidthCommand,
  DeltaRowHeightCommand,
  SetColWidthCommand,
  SetRowHeightCommand,
  SheetsSelectionsService,
} from '@univerjs/sheets'
import { dimensionIntents } from './dimension-intents'
import { columnWidthToPixels } from '@genoffice/xlsx-gateway/domain/sheet-dimensions'
import { originalAxisIndex } from '@genoffice/xlsx-gateway/domain/session-coordinates'
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
import type { WorkbookAppearance } from './appearance'
import { toSelectionFormat } from '../../../apps/sheets/src/renderer/selection-format'
import { ribbonIntents, type RibbonTarget } from './ribbon-intents'

export class WorkbookRuntime implements EditorProjection {
  get ribbonTarget(): RibbonTarget | null {
    const selected = this.selection
    if (!selected) return null
    return {
      sheetId: selected.sheetId,
      address: selected.address,
      format: toSelectionFormat(selected.style, selected.style.n?.pattern ?? 'General'),
      doubleUnderline: selected.style.ul?.s === 1 && selected.style.ul.t === 10,
    }
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
      protected: sheet?.protected ?? null,
    }
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
    const target = frozen?.target ?? this.ribbonTarget
    if (!target) throw new Error('Select cells before using this ribbon command.')
    const operations = commands.flatMap((command) => ribbonIntents(command, target))
    if (operations.length) await this.enqueue(operations)
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
    this.baseTheme = this.runtime.univer.__getInjector().get(ThemeService).getCurrentTheme()
    this.setAppearance(appearance)
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
    const state = workbookStateSchema.parse(data)
    if (state.stateId !== stateId || this.closed) throw new Error('invalid_projection')
    this.installing++
    this.container.setAttribute('aria-busy', 'true')
    try {
      const active = this.runtime.univerAPI.getActiveWorkbook(),
        sheetId = active?.getActiveSheet()?.getSheetId(),
        selected = active?.getActiveSheet()?.getActiveRange()?.getRange()
      const viewport = sheetId ? this.viewports.get(sheetId) : undefined
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
              { showGridlines?: boolean; showHeadings?: boolean; zoomScale?: number } | undefined
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
                  ...(span.width === undefined ? {} : { w: columnWidthToPixels(span.width) }),
                  hd: span.hidden ? 1 : 0,
                })
            const columnData: NonNullable<IWorkbookData['sheets'][string]['columnData']> = {}
            const rowData: NonNullable<IWorkbookData['sheets'][string]['rowData']> =
              Object.fromEntries(
                Object.entries((sheet.features?.rowHeights ?? {}) as Record<string, number>).map(
                  ([row, points]) => [Number(row) - 1, { h: (points * 4) / 3, ia: 0 }],
                ),
              )
            const filter = sheet.features?.filterEdits as
              | { visibilityRange: { startRow: number; endRow: number }; hiddenRows: number[] }
              | undefined
            if (filter) {
              const hidden = new Set(filter.hiddenRows)
              for (
                let row = filter.visibilityRange.startRow + 1;
                row <= filter.visibilityRange.endRow;
                row++
              )
                rowData[row] = { ...rowData[row], hd: hidden.has(row) ? 1 : 0 }
            }
            for (
              let column = 0;
              column < Math.min(16384, Math.max(20, sheet.columns + 5));
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
                rowCount: Math.min(1_048_576, Math.max(1000, sheet.rows + 100)),
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
                  ? { defaultColumnWidth: columnWidthToPixels(metadata.defaultColumnWidth) }
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
      if (viewport && sheetId && state.sheets.some((sheet) => sheet.id === sheetId)) {
        this.viewports.set(sheetId, viewport)
        this.runtime.univerAPI
          .getActiveWorkbook()
          ?.getActiveSheet()
          ?.scrollToCell(viewport.row, viewport.column, 0)
      }
      await this.loadVisible()
      this.changed()
    } finally {
      this.installing--
      if (!this.installing) this.container.setAttribute('aria-busy', 'false')
    }
  }
  private scheduleLoad() {
    if (this.loadTimer) clearTimeout(this.loadTimer)
    this.loadTimer = setTimeout(() => {
      void this.loadVisible().catch(this.report)
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
      pageSetup?: unknown
      cells: ViewCell[]
      merges?: GridRange[]
      rows?: { row: number; height?: number; hidden?: boolean; customHeight?: boolean }[]
    }
    const started = Date.now()
    for (;;) {
      try {
        result = (await this.host.read(sheet.getSheetId(), range, state.stateId)) as {
          cells: ViewCell[]
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
    const cells: Record<number, Record<number, ICellData>> = {}
    for (const cell of result.cells) {
      const point = parseAddress(cell.address)
      ;(cells[point.row] ??= {})[point.column] = gridCell(cell)
    }
    this.internal++
    try {
      const canonicalSheet = state.sheets.find((item) => item.id === sheet.getSheetId())!
      const heights = (canonicalSheet.features?.rowHeights ?? {}) as Record<string, number>
      const filter = canonicalSheet.features?.filterEdits as
        { visibilityRange: { startRow: number; endRow: number }; hiddenRows: number[] } | undefined
      const filterHidden = new Set(filter?.hiddenRows)
      for (const row of result.rows ?? []) {
        const hidden =
          filter &&
          row.row > filter.visibilityRange.startRow &&
          row.row <= filter.visibilityRange.endRow
            ? filterHidden.has(row.row)
            : row.hidden === true
        if (hidden && sheet.getSheet().getRowVisible(row.row)) sheet.hideRows(row.row, 1)
        if (!hidden && !sheet.getSheet().getRowVisible(row.row)) sheet.showRows(row.row, 1)
        const height = heights[String(row.row + 1)] ?? row.height
        if (height && height > 0) {
          if (heights[String(row.row + 1)] !== undefined || row.customHeight)
            sheet.setRowHeightsForced(row.row, 1, (height * 4) / 3)
          else sheet.setRowHeights(row.row, 1, (height * 4) / 3)
        }
      }
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
  get charts(): readonly SheetVisual[] {
    const sheetId = this.runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId()
    return (this.state?.sheets.find((sheet) => sheet.id === sheetId)?.features?.charts ??
      []) as unknown as SheetVisual[]
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
    this.runtime.univerAPI.getFormula().executeCalculation()
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
