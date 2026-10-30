import { IRenderManagerService, SHEET_VIEWPORT_KEY } from '@univerjs/engine-render'
import { SheetSkeletonManagerService } from '@univerjs/preset-sheets-core'
import { readCurrentOutline } from '@genoffice/workbook-session/read-outline'
import type { WorkbookSource } from '@genoffice/workbook-session/ports'
import { FORMULA_ERROR_VALUE } from '@genoffice/xlsx-gateway/shared/session-audit'
import { formulaReadsCell } from '@genoffice/xlsx-gateway/domain/formula-audit'
import {
  isVisualDragActive,
  isChartApplyActive,
} from '../../../../apps/sheets/src/renderer/WorkbookVisuals'
import { isSheetShapeDragging } from '../../../../apps/sheets/src/renderer/shape-draw'
import { buildChartVisual } from '@genoffice/xlsx-gateway/domain/chart-visual'
import { parseAddress } from '@genoffice/xlsx-gateway/domain/cell-address'
import { createRoot } from 'react-dom/client'
import { useState } from 'react'
import type { WorkbookAppearance } from '../../src/appearance'
import { WorkbookEditor } from '../../src/WorkbookEditor'
import type { WorkbookRuntime } from '../../src/runtime'
import {
  workbookStateSchema,
  transactionInputSchema,
} from '@genoffice/xlsx-gateway/shared/session-state'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import type { EditorHost } from '../../src/ports'
import '../../../../apps/vscode-office/src/webview/theme.css'
import { ThemeService } from '@univerjs/core'
import * as uiModule from '@univerjs/ui'

/** In-memory authority for UI tests; native sidecar/save are tested separately. */
export async function mountManualToolbar(
  appearanceControls = false,
  merged = false,
  autofitSource = false,
  navigationSource = false,
  headerFooterSource = false,
  chartSource = false,
  originalChartSource: boolean | 'uncached' | 'multiple' = false,
  sourceVisuals = false,
  pageSettingsSource = false,
) {
  let state = workbookStateSchema.parse({
      schemaVersion: 1,
      stateId: 'toolbar-before',
      structuralLog: [],
      sheets: [
        {
          id: 'toolbar-sheet',
          name: 'Data',
          rows: 3,
          columns: 2,
          cells: {
            A1: { kind: 'string', rawValue: 'Region' },
            B1: { kind: 'string', rawValue: 'Sales' },
            A2: { kind: 'string', rawValue: 'North' },
            B2: { kind: 'number', rawValue: 120 },
            A3: { kind: 'string', rawValue: 'South' },
            B3: { kind: 'number', rawValue: -30 },
          },
          features: merged
            ? { merges: [{ startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 }] }
            : {},
        },
      ],
    }),
    revision = 0
  let projection: WorkbookRuntime
  const sourceCells: Record<
    string,
    Omit<import('../../src/grid-cell').ViewCell, 'address'>
  > = autofitSource
    ? {
        A80: {
          kind: 'string' as const,
          rawValue: 'AutoFit must read this original cell beyond the first visible page',
        },
        D2: {
          kind: 'string' as const,
          rawValue: 'Wrapped original source text outside the selected column needs several lines',
          style: { wrapText: true, fontSize: 14, fontFamily: 'Calibri' },
        },
      }
    : {}
  if (autofitSource) {
    state.sheets[0]!.rows = 80
    state.sheets[0]!.columns = 4
    state.sheets[0]!.features = { originalMetadata: { rowCount: 80, columnCount: 4 } }
  }
  if (chartSource) {
    state.sheets[0]!.rows = 81
    state.sheets[0]!.columns = 4
    state.sheets[0]!.features = { originalMetadata: { rowCount: 81, columnCount: 4 } }
    Object.assign(sourceCells, {
      D78: { kind: 'blank', rawValue: null, formula: '=B79', calculationState: 'pending' },
      A81: { kind: 'number', rawValue: 2022 },
      B81: { kind: 'number', rawValue: 18 },
      C81: { kind: 'number', rawValue: 14 },
      A78: { kind: 'string', rawValue: 'Year' },
      B78: { kind: 'string', rawValue: 'Revenue' },
      C78: { kind: 'string', rawValue: 'Cost' },
      A79: { kind: 'number', rawValue: 2020 },
      B79: { kind: 'number', rawValue: 12 },
      C79: { kind: 'number', rawValue: 6 },
      A80: { kind: 'number', rawValue: 2021 },
      B80: { kind: 'number', rawValue: 8 },
      C80: { kind: 'number', rawValue: 10 },
    })
  }
  if (navigationSource) {
    state.features = {
      originalMetadata: {
        definedNames: [
          { name: 'GoTarget', formula: '#REF!', sheetIndex: 0 },
          { name: 'GoTarget', formula: "='Other Sheet'!$A$1" },
          { name: 'MovingTarget', formula: 'OFFSET(Data!$A$1,1,1)' },
        ],
      },
    }
    state.sheets.push({
      id: 'toolbar-other',
      name: 'Other Sheet',
      rows: 2,
      columns: 2,
      cells: { A1: { kind: 'string', rawValue: 'other original' } },
      features: {},
    })
  }
  if (originalChartSource) {
    const visual = buildChartVisual({
      id: 'original-chart',
      sheetId: 'source-sheet',
      sheetName: 'Data',
      chartType: 'column',
      dataRange: 'A1:B3',
      values: [
        ['Region', 'Sales'],
        ['North', 120],
        ['South', -30],
      ],
    })
    visual.editAs = 'twoCell'
    if (originalChartSource === 'uncached') {
      for (const series of visual.chart.series) {
        series.values = []
        series.categories = []
        series.name = 'Series1'
        series.nameRef = "'Data'!$B$1"
      }
    }
    state.features = JSON.parse(
      JSON.stringify({
        originalMetadata: {
          visuals: [
            {
              ...visual,
              chartPath: 'xl/charts/chart1.xml',
              drawingPath: 'xl/drawings/drawing1.xml',
              drawingIndex: 0,
            },
          ],
        },
      }),
    )
    if (originalChartSource === 'multiple') {
      const visuals = (state.features!.originalMetadata as any).visuals
      const template = visuals[0]
      visuals.splice(
        0,
        1,
        ...[0, 1, 2].map((index) => ({
          ...structuredClone(template),
          id: `source-layer-${index}`,
          chartPath: `xl/charts/chart${index + 1}.xml`,
          drawingIndex: index,
          chart: { ...structuredClone(template.chart), title: `Layer ${index}` },
          anchor: {
            ...template.anchor,
            fromColumn: 3 + index * 3,
            toColumn: 5 + index * 3,
            fromRow: 0,
            toRow: 6,
          },
        })),
      )
    }
    state.sheets[0]!.originalSheetId = 'source-sheet'
    state.sheetMap['toolbar-sheet'] = { originalSheetId: 'source-sheet', originalName: 'Data' }
  }
  const mediaReads: string[] = []
  if (sourceVisuals) {
    const anchor = (fromColumn: number, fromRow: number, toColumn: number, toRow: number) => ({
      fromColumn,
      fromRow,
      toColumn,
      toRow,
      fromColumnOffset: 0,
      fromRowOffset: 0,
      toColumnOffset: 0,
      toRowOffset: 0,
    })
    const boxes = [anchor(3, 0, 5, 3), anchor(7, 2, 10, 6), anchor(4, 8, 6, 11)]
    state.features = {
      originalMetadata: {
        visuals: [
          ...boxes.map((box, i) => ({
            id: `shape-${i}`,
            kind: 'shape',
            sheetId: 'source-sheet',
            drawingPath: 'xl/drawings/drawing1.xml',
            drawingIndex: i,
            shapeType: 'rect',
            name: `Shape ${i}`,
            altText: `Shape ${i}`,
            fillColor: '#3399FF',
            lineColor: '#222222',
            text: `Shape ${i}`,
            anchor: box,
            editAs: 'twoCell',
          })),
          {
            id: 'picture-0',
            kind: 'image',
            sheetId: 'source-sheet',
            drawingPath: 'xl/drawings/drawing1.xml',
            drawingIndex: 3,
            altText: 'Picture 0',
            anchor: anchor(9, 8, 12, 12),
            editAs: 'twoCell',
          },
        ],
      },
    }
    state.sheets[0]!.originalSheetId = 'source-sheet'
    state.sheetMap['toolbar-sheet'] = { originalSheetId: 'source-sheet', originalName: 'Data' }
  }
  if (pageSettingsSource)
    state.sheets[0]!.features = {
      ...state.sheets[0]!.features,
      originalMetadata: {
        name: 'Data',
        printArea: "'Data'!$B$2:$E$10",
        printTitles: "'Data'!$1:$3",
      },
    }
  const reads: Parameters<EditorHost['read']>[1][] = []
  let readGate: Promise<void> | undefined
  state.features = {
    ...state.features,
    originalMetadata: JSON.parse(
      JSON.stringify({
        ...(state.features?.originalMetadata as Record<string, unknown> | undefined),
        nameBindingsComplete: true,
        reservedDefinedNames: [],
      }),
    ),
  }
  let submitGate: Promise<void> | undefined
  let releaseSubmit: (() => void) | undefined
  let submitError: string | undefined
  let submitCount = 0
  let releaseRead: (() => void) | undefined
  const host: EditorHost = {
    async outline(sheetId, stateId) {
      if (stateId !== state.stateId) throw new Error('stale_outline_fixture')
      const captured = state
      if (readGate) await readGate
      return readCurrentOutline({
        state: captured,
        source: {
          readRowOutline: async () => ({ indexingComplete: true, rows: [] }),
        } as unknown as WorkbookSource,
        sessionId: 'fixture',
        sheetId,
      })
    },
    async audit(request) {
      if (readGate) await readGate
      const owner = state.sheets.find((sheet) => sheet.id === request.sheetId)!
      const hits = state.sheets.flatMap((sheet, index) =>
        Object.entries({ ...(index === 0 ? sourceCells : {}), ...sheet.cells }).flatMap(
          ([address, cell]) => {
            const matches =
              request.kind === 'errors'
                ? typeof cell.rawValue === 'string' && FORMULA_ERROR_VALUE.test(cell.rawValue)
                : cell.formula &&
                  !(sheet.id === owner.id && address === request.address) &&
                  formulaReadsCell(
                    cell.formula,
                    sheet.id === owner.id,
                    owner.name,
                    parseAddress(request.address),
                  )
            return matches
              ? [
                  {
                    sheetId: sheet.id,
                    address,
                    ...(request.kind === 'errors' ? { value: String(cell.rawValue) } : {}),
                  },
                ]
              : []
          },
        ),
      )
      return { stateId: request.stateId, valuesPending: false, hits }
    },
    async media(visualId) {
      mediaReads.push(visualId)
      return {
        mediaType: 'image/png',
        base64:
          'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
      }
    },
    async read(sheetId, range) {
      reads.push({ ...range })
      if (readGate) await readGate
      const filter = state.sheets.find((sheet) => sheet.id === sheetId)!.features?.filterEdits as
        { filter: { range: unknown; columns: unknown[] } | null } | undefined
      return {
        autoFilter: filter?.filter?.range ?? null,
        autoFilterColumns: filter?.filter?.columns ?? [],
        pageBreaks: pageSettingsSource ? { rowBreaks: [2, 10], colBreaks: [1] } : undefined,
        pageSetup: pageSettingsSource
          ? {
              orientation: 'landscape',
              paperSize: 5,
              fitToWidth: 2,
              fitToHeight: 1,
              fitToPage: true,
              printGridlines: true,
              printHeadings: false,
            }
          : headerFooterSource
            ? { oddHeader: '&LOriginal left&CPage &P of &N', oddFooter: '&RConfidential' }
            : null,
        merges:
          state.sheets.find((sheet) => sheet.id === sheetId)!.features?.mergeEdits ??
          state.sheets.find((sheet) => sheet.id === sheetId)!.features?.merges ??
          [],
        cells: Object.entries({
          ...sourceCells,
          ...state.sheets.find((sheet) => sheet.id === sheetId)!.cells,
        })
          .filter(([address]) => {
            const point = parseAddress(address)
            return (
              point.row >= range.startRow &&
              point.row <= range.endRow &&
              point.column >= range.startColumn &&
              point.column <= range.endColumn
            )
          })
          .map(([address, cell]) => ({ address, ...cell })),
      }
    },
    async submit(raw) {
      submitCount++
      await submitGate
      if (submitError) {
        const error = submitError
        submitError = undefined
        throw new Error(error)
      }
      const intent = transactionInputSchema.parse({
        schemaVersion: 1,
        sessionId: 'toolbar',
        origin: 'manual',
        summary: 'UI edit',
        ...raw,
      })
      if (intent.expectedRevision !== revision) throw new Error('stale_revision')
      const material = structuredClone(state)
      if (chartSource) material.sheets[0]!.cells = { ...sourceCells, ...material.sheets[0]!.cells }
      const result = await computeCandidateState(state, intent, {
        state: material,
        completeSheetIds: state.sheets.map((sheet) => sheet.id),
        writableSheetIds: state.sheets.map((sheet) => sheet.id),
        viewWritableSheetIds: state.sheets
          .filter((sheet) => !sheet.readOnly)
          .map((sheet) => sheet.id),
        allocation: {
          stateId: crypto.randomUUID(),
          sheetIds: Object.fromEntries(
            intent.operations.flatMap((op, index) =>
              op.op === 'add_sheet' ? [[String(index), crypto.randomUUID()]] : [],
            ),
          ),
        },
      })
      state = result.nextState
      revision++
      await projection.install(state.stateId, revision, state)
      return {
        sessionId: 'toolbar',
        transactionId: intent.transactionId,
        status: 'committed',
        revision,
        stateId: state.stateId,
        previousRevision: revision - 1,
        affectedTargets: result.affectedTargets,
        summary: intent.summary,
        calculationState: result.calculationState,
      }
    },
    selection() {},
    async history() {
      throw new Error('History is outside this fixture')
    },
  }
  const root = createRoot(document.body)
  if (appearanceControls) document.documentElement.dataset.workbookAppearance = 'excel'
  await new Promise<void>((resolve, reject) => {
    const onReady = (runtime: WorkbookRuntime) => {
      projection = runtime
      void runtime.install(state.stateId, revision, state).then(resolve, reject)
    }
    function Fixture() {
      const [appearance, setAppearance] = useState<WorkbookAppearance>('excel')
      return (
        <WorkbookEditor
          host={host}
          locale="en"
          onReady={onReady}
          appearance={appearanceControls ? appearance : 'vscode'}
          onAppearanceChange={
            appearanceControls
              ? (value) => {
                  document.documentElement.dataset.workbookAppearance = value
                  setAppearance(value)
                }
              : undefined
          }
        />
      )
    }
    root.render(<Fixture />)
  })
  const facade = (projection! as unknown as { runtime: { univerAPI: any } }).runtime.univerAPI
  return {
    diagnostics() {
      const runtime = projection as unknown as {
        pendingError: boolean
        composing: boolean
        installing: number
        visualWork?: unknown
        chartWork?: unknown
        symbolWork?: unknown
        filterPopupWork?: unknown
        filterPopupContext?: unknown
        filterPopupRejected?: boolean
      }
      return {
        pendingError: runtime.pendingError,
        composing: runtime.composing,
        installing: runtime.installing,
        visualWork: !!runtime.visualWork,
        chartWork: !!runtime.chartWork,
        symbolWork: !!runtime.symbolWork,
        filterPopupWork: !!runtime.filterPopupWork,
        filterPopupContext: runtime.filterPopupContext,
        filterPopupRejected: runtime.filterPopupRejected,
        editing: facade.getActiveWorkbook()?.isCellEditing(),
        drag: isVisualDragActive(),
        chartApply: isChartApplyActive(),
        shapeDrag: isSheetShapeDragging(),
        active: document.activeElement?.outerHTML.slice(0, 200),
      }
    },
    async seedSpecialSource() {
      state = {
        ...state,
        stateId: crypto.randomUUID(),
        sheets: state.sheets.map((sheet) => ({
          ...sheet,
          rows: 100,
          columns: 12,
          features: { ...sheet.features, colWidths: { A: 80, B: 120 } },
          cells: {
            ...sheet.cells,
            A80: {
              kind: 'number' as const,
              rawValue: 2,
              numberFormat: '0.00',
              style: {
                bold: true,
                numberFormat: '0.00',
                fillColor: '#FFDD99',
                borderRight: { style: 'thin' as const },
              },
            },
            B80: { kind: 'blank' as const, rawValue: null },
            A81: { kind: 'number' as const, rawValue: 4 },
            B81: { kind: 'number' as const, rawValue: 0 },
            D5: {
              kind: 'number' as const,
              rawValue: 10,
              style: {
                italic: true,
                fillColor: '#CCEEFF',
                borderLeft: { style: 'double' as const, color: '#123456' },
              },
            },
            E5: { kind: 'number' as const, rawValue: 20 },
            D6: { kind: 'number' as const, rawValue: 30 },
            E6: { kind: 'number' as const, rawValue: 40 },
          },
        })),
      }
      revision++
      await projection.install(state.stateId, revision, state)
    },
    formatSpecialSource() {
      return projection.enqueue([
        {
          op: 'format_range',
          sheetId: 'toolbar-sheet',
          range: 'A80',
          format: { bold: false, fontColor: '#445566' },
        },
      ])
    },
    setSpecialFormula() {
      return projection.enqueue([
        { op: 'set_formula', sheetId: 'toolbar-sheet', address: 'B81', formula: '=A80+$A$1' },
      ])
    },
    resizeSpecialSource() {
      return projection.enqueue([
        { op: 'set_col_width', sheetId: 'toolbar-sheet', column: 'A', count: 1, widthPx: 240 },
      ])
    },
    async seedCutSource() {
      state = {
        ...state,
        stateId: crypto.randomUUID(),
        sheets: state.sheets.map((sheet) => ({
          ...sheet,
          rows: 100,
          columns: 10,
          cells: {
            ...sheet.cells,
            A80: { kind: 'number' as const, rawValue: 12, style: { bold: true } },
            B80: {
              kind: 'blank' as const,
              rawValue: null,
              formula: '=A80+$A$1',
              calculationState: 'pending' as const,
            },
            J90: {
              kind: 'blank' as const,
              rawValue: null,
              formula: '=$A$80',
              calculationState: 'pending' as const,
            },
          },
        })),
      }
      revision++
      await projection.install(state.stateId, revision, state)
    },
    async seedClipboardSource() {
      sourceCells.A80 = {
        kind: 'string',
        rawValue: '=literal',
        style: {
          bold: true,
          fillColor: '#FFEECC',
          borderLeft: { style: 'mediumDashDotDot', color: '#123456' },
        },
      }
      sourceCells.B80 = {
        kind: 'string',
        rawValue: '=literalRegion',
        formula: '=A80&$A$1',
        calculationState: 'ready',
        cachedValue: '=literalRegion',
        style: { italic: true },
      }
      sourceCells.A81 = {
        kind: 'date',
        rawValue: 45000,
        dateSystem: '1900',
        numberFormat: 'yyyy-mm-dd',
      }
      sourceCells.B81 = { kind: 'error', rawValue: '#DIV/0!' }
      state = {
        ...state,
        stateId: crypto.randomUUID(),
        dateSystem: '1900',
        sheets: state.sheets.map((s) => ({ ...s, rows: Math.max(s.rows, 81) })),
      }
      revision++
      await projection.install(state.stateId, revision, state)
    },
    changeCopiedSource() {
      return projection.enqueue([
        { op: 'set_cell', sheetId: 'toolbar-sheet', address: 'A80', value: 'New source' },
      ])
    },
    beginMergedEdit() {
      const book = facade.getActiveWorkbook()
      const sheet = book.getActiveSheet()
      sheet.setActiveRange(sheet.getRange('A1:B1'))
      return book.startEditing()
    },
    reveal: (range: string) => projection.goToReference(range, { focus: false }),
    facade: () => facade,
    /** Leaf commands of the cell context menu, for menu parity audits. */
    contextMenuCommands() {
      const { IMenuManagerService, ContextMenuPosition } = uiModule
      const injector = (projection as unknown as { runtime: { univer: any } }).runtime.univer.__getInjector()
      const out: { key: string; id: string; title?: string; params?: unknown }[] = []
      const walk = (items: any[]) => {
        for (const item of items ?? []) {
          if (item.item?.id) out.push({ key: item.key, id: item.item.commandId ?? item.item.id, title: item.item.title, params: item.item.params })
          walk(item.children)
        }
      }
      for (const position of [ContextMenuPosition.MAIN_AREA, ContextMenuPosition.ROW_HEADER, ContextMenuPosition.COL_HEADER])
        walk(injector.get(IMenuManagerService).getMenuByPositionKey(position))
      return out
    },
    select(range: string) {
      const sheet = facade.getActiveWorkbook().getActiveSheet()
      sheet.setActiveRange(sheet.getRange(range))
    },
    selectSheet(name: string) {
      const sheet = state.sheets.find((sheet) => sheet.name === name)!
      facade.getActiveWorkbook().setActiveSheet(sheet.id)
    },
    value(address: string) {
      return facade.getActiveWorkbook().getActiveSheet().getRange(address).getValue()
    },
    styles: (range: string) => projection.projectedStyles(range),
    outlinePoint(axis: 'rows' | 'cols', index: number) {
      const runtime = (projection as any).runtime
      const workbook = facade.getActiveWorkbook()
      const render = runtime.univer
        .__getInjector()
        .get(IRenderManagerService)
        .getRenderById(workbook.getId())
      const skeleton = render.with(SheetSkeletonManagerService).getCurrentSkeleton()
      const viewport = render.scene.getViewport(SHEET_VIEWPORT_KEY.VIEW_MAIN)
      const scale = render.scene.getAncestorScale()
      const accumulation =
        axis === 'rows' ? skeleton.rowHeightAccumulation : skeleton.columnWidthAccumulation
      const mid = ((index > 0 ? accumulation[index - 1] : 0) + accumulation[index]) / 2
      const grid = render.engine.getCanvasElement().getBoundingClientRect()
      return axis === 'rows'
        ? {
            x: grid.left,
            y: grid.top + viewport.top + (mid - viewport.viewportScrollY) * scale.scaleY,
          }
        : {
            x: grid.left + viewport.left + (mid - viewport.viewportScrollX) * scale.scaleX,
            y: grid.top,
          }
    },
    axisVisibility(axis: 'rows' | 'cols', indices: number[]) {
      const sheet = facade.getActiveWorkbook().getActiveSheet().getSheet()
      return indices.map((index) =>
        axis === 'rows' ? sheet.getRowVisible(index) : sheet.getColVisible(index),
      )
    },
    holdSubmissions() {
      submitGate = new Promise<void>((resolve) => {
        releaseSubmit = resolve
      })
    },
    releaseSubmissions() {
      releaseSubmit?.()
      submitGate = undefined
      releaseSubmit = undefined
    },
    rejectNextSubmission(error: string) {
      submitError = error
    },
    submissions: () => submitCount,
    holdReads() {
      readGate = new Promise<void>((resolve) => {
        releaseRead = resolve
      })
    },
    releaseReads() {
      releaseRead?.()
      readGate = undefined
      releaseRead = undefined
    },
    selection: () => projection.selectionContext,
    reads: () => structuredClone(reads),
    sizes: () => projection.projectedDimensions('A1:C3'),
    flush: (options?: { commitChartDraft?: boolean }) => projection.flushEditor(options),
    cellRect: (address: string) => {
      const cell = facade.getActiveWorkbook().getActiveSheet().getRange(address).getCellRect()
      return { left: cell.left, right: cell.right, top: cell.top, bottom: cell.bottom }
    },
    view: () => projection.projectedView(),
    merges: () => projection.projectedMerges(),
    rowVisible: (row: number) =>
      facade.getActiveWorkbook().getActiveSheet().getSheet().getRowVisible(row),
    formula() {
      try {
        facade.getActiveWorkbook().getActiveSheet().getRange('C2').setFormula('=B2*2')
      } catch (error) {
        if (!String(error).includes('manual_intent_forwarded')) throw error
      }
    },
    seedClearTargets() {
      return projection.enqueue([
        { op: 'set_cell', sheetId: 'toolbar-sheet', address: 'B2', value: 120 },
        {
          op: 'format_range',
          sheetId: 'toolbar-sheet',
          range: 'B2',
          format: { bold: true, indent: 3, fillColor: '#FFFF00' },
        },
        { op: 'set_note', sheetId: 'toolbar-sheet', address: 'A2', text: 'keep note' },
        { op: 'set_note', sheetId: 'toolbar-sheet', address: 'B2', text: 'clear note' },
        {
          op: 'set_hyperlink',
          sheetId: 'toolbar-sheet',
          address: 'A2',
          target: 'https://example.com/keep',
        },
        {
          op: 'set_hyperlink',
          sheetId: 'toolbar-sheet',
          address: 'B2',
          target: 'https://example.com/clear',
        },
      ])
    },
    seedMixedIndent() {
      return projection.enqueue([
        { op: 'format_range', sheetId: 'toolbar-sheet', range: 'A2', format: { indent: 1 } },
        { op: 'format_range', sheetId: 'toolbar-sheet', range: 'B2', format: { indent: 3 } },
      ])
    },
    async overlapVisualLayers() {
      await projection.enqueue(
        projection.charts.map((v) => ({
          op: 'set_visual_anchor',
          sheetId: v.sheetId,
          visualId: v.id,
          anchor: {
            ...v.anchor,
            fromColumn: 3,
            toColumn: 5,
            fromRow: 0,
            toRow: 6,
            fromColumnOffset: 0,
            fromRowOffset: 0,
            toColumnOffset: 0,
            toRowOffset: 0,
          },
        })),
      )
    },
    async seedVisualLayers() {
      await projection.enqueue(
        [0, 1, 2].map((i) => ({
          op: 'add_chart',
          sheetId: 'toolbar-sheet',
          dataRange: 'A1:B3',
          chartType: 'column',
          title: `Layer ${i}`,
        })),
      )
      const visuals = state.sheets[0]!.features!.charts as any[]
      await projection.enqueue(
        visuals.map((v, i) => ({
          op: 'set_visual_anchor',
          sheetId: 'toolbar-sheet',
          visualId: v.id,
          anchor: { ...v.anchor, fromColumn: 3 + i * 3, toColumn: 5 + i * 3, fromRow: 0, toRow: 6 },
        })),
      )
    },
    mediaReads: () => [...mediaReads],
    putFormula(address: string, formula: string) {
      return projection.enqueue([
        { op: 'set_formula', sheetId: state.sheets[0]!.id, address, formula },
      ])
    },
    putHeaders(labels: string[]) {
      return projection.enqueue([
        {
          op: 'paste_cells',
          sheetId: state.sheets[0]!.id,
          start: 'A1',
          cells: [labels.map((label) => ({ kind: 'string' as const, rawValue: label }))],
        },
      ])
    },
    dateLabel() {
      return projection.enqueue([
        { op: 'set_cell', sheetId: state.sheets[0]!.id, address: 'A1', value: 44927 },
        {
          op: 'format_range',
          sheetId: state.sheets[0]!.id,
          range: 'A1:A1',
          format: { numberFormat: 'mmm yyyy' },
        },
      ])
    },
    setName(name: string, ref: string, scopeSheetId?: string) {
      return projection.enqueue([
        { op: 'add_defined_name', name, ref, ...(scopeSheetId ? { scopeSheetId } : {}) },
      ])
    },
    deleteName(name: string, scopeSheetId?: string) {
      return projection.enqueue([
        { op: 'delete_defined_name', name, ...(scopeSheetId ? { scopeSheetId } : {}) },
      ])
    },
    async protectViewSheet() {
      state.stateId = crypto.randomUUID()
      state.sheets[0]!.protected = true
      await projection.install(state.stateId, revision, state)
    },
    async seedDedupeSource() {
      const sheet = state.sheets[0]!
      sheet.rows = 5
      sheet.cells = {
        A1: { kind: 'string', rawValue: 'North' },
        B1: {
          kind: 'number',
          rawValue: 120,
          formula: '=60*2',
          cachedValue: 120,
          calculationState: 'ready',
        },
        A2: { kind: 'string', rawValue: 'South' },
        B2: { kind: 'boolean', rawValue: false },
        A3: { kind: 'string', rawValue: 'north', style: { fillColor: '#CCEEFF' } },
        B3: { kind: 'number', rawValue: 120 },
        A4: { kind: 'string', rawValue: 'south' },
        B4: { kind: 'boolean', rawValue: false },
        A5: { kind: 'string', rawValue: 'South' },
        B5: { kind: 'string', rawValue: 'false' },
      }
      await projection.install(state.stateId, revision, state)
    },
    async seedAuditSource(showFormulas = false) {
      Object.assign(sourceCells, {
        C100001: { kind: 'number', rawValue: 121, formula: '=B2+1', calculationState: 'ready' },
        D100001: { kind: 'error', rawValue: '#DIV/0!', formula: '=1/0', calculationState: 'ready' },
        E100001: { kind: 'string', rawValue: '#not-an-error' },
      })
      state.stateId = crypto.randomUUID()
      state.sheets[0]!.rows = 100001
      state.sheets[0]!.columns = 5
      state.sheets[0]!.features = {
        ...state.sheets[0]!.features,
        originalMetadata: { showFormulas },
      }
      state.sheets.push({
        id: 'audit-other',
        name: 'Other',
        rows: 2,
        columns: 2,
        cells: {
          A1: { kind: 'number', rawValue: 120, formula: '=Data!B2', calculationState: 'ready' },
          B2: { kind: 'error', rawValue: '#REF!' },
        },
        features: {},
      })
      await projection.install(state.stateId, revision, state)
    },
    displayedCell(address: string) {
      const point = parseAddress(address)
      return facade.getActiveWorkbook().getActiveSheet().getSheet().getCell(point.row, point.column)
    },
    snapshot: () => structuredClone(state),
    revision: () => revision,
    darkTheme: () =>
      (
        projection as unknown as {
          runtime: {
            univer: { __getInjector(): { get(token: typeof ThemeService): ThemeService } }
          }
        }
      ).runtime.univer
        .__getInjector()
        .get(ThemeService).darkMode,
    editChartSource(value: number) {
      return projection.enqueue([
        { op: 'set_cell', sheetId: 'toolbar-sheet', address: 'B79', value },
      ])
    },
    async seedFilterColors() {
      state = structuredClone(state)
      state.stateId = crypto.randomUUID()
      state.sheets[0]!.rows = 5
      state.sheets[0]!.cells = {
        A1: { kind: 'string', rawValue: 'Color' },
        A2: {
          kind: 'string',
          rawValue: 'Red',
          style: { fillColor: '#FF0000', fontColor: '#0000FF' },
        },
        A3: {
          kind: 'string',
          rawValue: 'Blue',
          style: { fillColor: '#0000FF', fontColor: '#FF0000' },
        },
        A4: { kind: 'string', rawValue: 'Plain' },
        A5: {
          kind: 'string',
          rawValue: 'White',
          style: { fillColor: '#FFFFFF', fontColor: '#0000FF' },
        },
      }
      await projection.install(state.stateId, revision, state)
    },
    editFill(address: string, color: string) {
      return projection.enqueue([
        {
          op: 'format_range',
          sheetId: 'toolbar-sheet',
          range: address,
          format: { fillColor: color },
        },
      ])
    },
    async seedSubtotal() {
      state = structuredClone(state)
      state.stateId = crypto.randomUUID()
      state.sheets[0]!.rows = 8
      state.sheets[0]!.columns = 3
      state.sheets[0]!.cells = {
        B3: { kind: 'string', rawValue: 'Region' },
        C3: { kind: 'string', rawValue: 'Amount' },
        B4: { kind: 'string', rawValue: 'North' },
        C4: { kind: 'number', rawValue: 12.5 },
        B5: { kind: 'string', rawValue: 'North' },
        C5: { kind: 'number', rawValue: 7 },
        B6: { kind: 'string', rawValue: 'South' },
        C6: { kind: 'number', rawValue: 3 },
        A8: { kind: 'number', rawValue: 42 },
      }
      await projection.install(state.stateId, revision, state)
      facade
        .getActiveWorkbook()
        .getActiveSheet()
        .setActiveRange(facade.getActiveWorkbook().getActiveSheet().getRange('B3:C6'))
    },
    async seedTextSplit(text = 'left,right') {
      state = structuredClone(state)
      state.stateId = crypto.randomUUID()
      state.sheets[0]!.rows = 6
      state.sheets[0]!.columns = 2
      state.sheets[0]!.cells = {
        A1: {
          kind: 'number',
          rawValue: 0,
          formula: '=D4',
          calculationState: 'ready',
          cachedValue: 0,
        },
        B4: { kind: 'string', rawValue: text },
        B6: { kind: 'string', rawValue: text },
      }
      await projection.install(state.stateId, revision, state)
      const sheet = facade.getActiveWorkbook().getActiveSheet()
      sheet.setActiveRange(sheet.getRange('B4:B6'))
    },
    splitText: (delimiter: 1 | 2 | 4 | 8) => projection.textToColumns(delimiter),
    async seedVisibleSort() {
      state = structuredClone(state)
      state.stateId = crypto.randomUUID()
      state.sheets[0]!.rows = 6
      state.sheets[0]!.cells = {
        A1: { kind: 'string', rawValue: 'Group' },
        B1: { kind: 'string', rawValue: 'Amount' },
        A2: { kind: 'string', rawValue: 'Keep' },
        B2: { kind: 'number', rawValue: 30, style: { bold: true } },
        A3: { kind: 'string', rawValue: 'Drop' },
        B3: { kind: 'number', rawValue: 999 },
        A4: { kind: 'string', rawValue: 'Keep' },
        B4: { kind: 'number', rawValue: 10, style: { fillColor: '#FF0000' } },
        A5: { kind: 'string', rawValue: 'Keep' },
        B5: { kind: 'number', rawValue: 5 },
        A6: { kind: 'string', rawValue: 'Keep' },
        B6: { kind: 'number', rawValue: 20, numberFormat: '0.00' },
      }
      await projection.install(state.stateId, revision, state)
    },
    hideRow(row: number) {
      return projection.enqueue([
        { op: 'set_rows_hidden', sheetId: 'toolbar-sheet', row, count: 1, hidden: true },
      ])
    },
    async seedFilterFormulaSource() {
      sourceCells.Z900 = { kind: 'number', rawValue: 42 }
      state = structuredClone(state)
      state.stateId = crypto.randomUUID()
      state.sheets[0]!.rows = 900
      state.sheets[0]!.cells.A2 = {
        kind: 'number',
        rawValue: 42,
        formula: '=Z900',
        cachedValue: 42,
        calculationState: 'ready',
      }
      await projection.install(state.stateId, revision, state)
    },
    createFilter(range = 'A1:B3') {
      return projection.enqueue([{ op: 'set_filter', sheetId: 'toolbar-sheet', range }])
    },
    editNote(value: string) {
      return projection.enqueue([
        { op: 'set_note', sheetId: 'toolbar-sheet', address: 'B2', text: value },
      ])
    },
    editValue(address: string, value: string | number | boolean | null) {
      return projection.enqueue([{ op: 'set_cell', sheetId: 'toolbar-sheet', address, value }])
    },
    edit() {
      return projection.enqueue([
        { op: 'set_cell', sheetId: 'toolbar-sheet', address: 'B2', value: 200 },
      ])
    },
    dispose() {
      releaseRead?.()
      releaseSubmit?.()
      root.unmount()
    },
  }
}
