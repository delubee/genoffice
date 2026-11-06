import {
  buildChartVisual,
  chartKindFromVisual,
  type SheetChartVisual as SheetVisual,
} from '@genoffice/xlsx-gateway/domain/chart-visual'
import { expect, it } from 'vitest'
import { initialState } from '../src/state'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'

function fixture() {
  let state = structuredClone(
    initialState({ sheets: [{ id: 'base', name: 'Data', rowCount: 3, columnCount: 2 }] }),
  )
  const sheetId = state.sheets[0]!.id
  state.sheets[0]!.cells = {
    A1: { kind: 'string', rawValue: 'Region' },
    B1: { kind: 'string', rawValue: 'Sales' },
    A2: { kind: 'string', rawValue: 'North' },
    B2: { kind: 'number', rawValue: 12.5, numberFormat: '0.00', style: { numberFormat: '0.00' } },
    A3: { kind: 'string', rawValue: 'South' },
    B3: { kind: 'number', rawValue: 8 },
  }
  return {
    sheetId,
    get state() {
      return state
    },
    async apply(operations: WorkbookOperation[], hiddenRows?: number[]) {
      const before = structuredClone(state)
      const result = await computeCandidateState(
        state,
        {
          schemaVersion: 1,
          sessionId: 'ribbon',
          transactionId: crypto.randomUUID(),
          expectedRevision: 0,
          origin: 'manual',
          summary: 'Ribbon action',
          operations,
        },
        {
          state: structuredClone(state),
          ...(hiddenRows ? { hiddenRows: { [sheetId]: hiddenRows } } : {}),
          completeSheetIds: state.sheets.map((sheet) => sheet.id),
          writableSheetIds: state.sheets.map((sheet) => sheet.id),
          allocation: { stateId: crypto.randomUUID(), sheetIds: {} },
        },
      )
      expect(result.inverse.state).toEqual(before)
      state = result.nextState
      return result
    },
  }
}
it('commits page, view, freeze and note settings atomically without cell rewrites', async () => {
  const f = fixture(),
    before = structuredClone(f.state.sheets[0]!.cells)
  await f.apply([
    {
      op: 'set_page_setup',
      sheetId: f.sheetId,
      orientation: 'landscape',
      paperSize: 8,
      margins: 'narrow',
      showGridlines: false,
      zoomScale: 125,
    },
    { op: 'set_freeze', sheetId: f.sheetId, rows: 1, columns: 0 },
    { op: 'set_note', sheetId: f.sheetId, address: 'B2', text: 'Check & approve' },
  ])
  expect(f.state.sheets[0]!.features).toMatchObject({
    pageSetupEdits: {
      orientation: 'landscape',
      paperSize: 8,
      showGridlines: false,
      zoomScale: 125,
    },
    freezeEdits: { frozenRows: 1, frozenColumns: 0 },
    noteEdits: { B2: 'Check & approve' },
  })
  expect(f.state.sheets[0]!.cells).toEqual(before)
  await f.apply([{ op: 'set_note', sheetId: f.sheetId, address: 'B2', text: null }])
  expect(f.state.sheets[0]!.features?.noteEdits).toEqual({ B2: null })
})
it('filters by current formatted values, sorts only visible rows and clears row visibility', async () => {
  const f = fixture()
  await f.apply([
    { op: 'set_filter', sheetId: f.sheetId, range: 'A1:B3' },
    { op: 'set_filter_criteria', sheetId: f.sheetId, column: 'B', values: ['12.50'] },
  ])
  expect(f.state.sheets[0]!.features?.filterEdits).toMatchObject({ hiddenRows: [2] })
  const before = structuredClone(f.state.sheets[0]!.cells)
  await f.apply([
    {
      op: 'sort_range',
      sheetId: f.sheetId,
      range: 'A1:B3',
      byColumn: 'B',
      order: 'asc',
      hasHeader: true,
    },
  ])
  expect(f.state.sheets[0]!.cells).toEqual(before)
  expect(f.state.sheets[0]!.features?.filterEdits).toMatchObject({ hiddenRows: [2] })
  await f.apply([{ op: 'clear_filter', sheetId: f.sheetId }])
  expect(f.state.sheets[0]!.features?.filterEdits).toMatchObject({ filter: null, hiddenRows: [] })
})
it('sorting moves typed dates and their formats instead of coercing them to ordinary numbers', async () => {
  const f = fixture()
  f.state.sheets[0]!.cells.A3 = {
    kind: 'date',
    rawValue: 60,
    dateSystem: '1900',
    style: { numberFormat: 'yyyy-mm-dd', bold: true },
    numberFormat: 'yyyy-mm-dd',
  }
  await f.apply([
    {
      op: 'sort_range',
      sheetId: f.sheetId,
      range: 'A1:B3',
      byColumn: 'B',
      order: 'asc',
      hasHeader: true,
    },
  ])
  expect(f.state.sheets[0]!.cells.A2).toMatchObject({
    kind: 'date',
    rawValue: 60,
    dateSystem: '1900',
    numberFormat: 'yyyy-mm-dd',
    style: { bold: true },
  })
  expect(f.state.sheets[0]!.cells.B2!.rawValue).toBe(8)
})
it('keeps image notes and formula sorts unchanged when unsupported', async () => {
  const f = fixture()
  f.state.sheets[0]!.features = { originalMetadata: { comments: [{ image: { id: 'picture' } }] } }
  await expect(
    f.apply([{ op: 'set_note', sheetId: f.sheetId, address: 'B2', text: 'new' }]),
  ).rejects.toThrow('unsupported_note_image')
  f.state.sheets[0]!.cells.B2 = {
    kind: 'number',
    rawValue: 2,
    formula: '=1+1',
    calculationState: 'ready',
  }
  await expect(
    f.apply([
      {
        op: 'sort_range',
        sheetId: f.sheetId,
        range: 'A1:B3',
        byColumn: 'B',
        order: 'asc',
        hasHeader: true,
      },
    ]),
  ).rejects.toThrow()
})

it('uses the workbook date system when matching displayed filter values', async () => {
  const f = fixture()
  f.state.dateSystem = '1904'
  f.state.sheets[0]!.cells.B2 = {
    kind: 'date',
    rawValue: 1,
    dateSystem: '1904',
    numberFormat: 'yyyy-mm-dd',
    style: { numberFormat: 'yyyy-mm-dd' },
  }
  await f.apply([
    { op: 'set_filter', sheetId: f.sheetId, range: 'B1:B3' },
    { op: 'set_filter_criteria', sheetId: f.sheetId, column: 'B', values: ['1904-01-02'] },
  ])
  expect(f.state.sheets[0]!.features?.filterEdits).toMatchObject({ hiddenRows: [2] })
})

it('sorts Han text by an explicit pinyin collation independently of the host UI language', async () => {
  const f = fixture()
  f.state.sheets[0]!.cells.A2 = { kind: 'string', rawValue: '中' }
  f.state.sheets[0]!.cells.A3 = { kind: 'string', rawValue: '阿' }
  await f.apply([
    {
      op: 'sort_range',
      sheetId: f.sheetId,
      range: 'A1:B3',
      byColumn: 'A',
      order: 'asc',
      hasHeader: true,
    },
  ])
  expect(f.state.sheets[0]!.cells.A2!.rawValue).toBe('阿')
  expect(f.state.sheets[0]!.cells.A3!.rawValue).toBe('中')
  expect(f.state.sheets[0]!.cells.B2!.rawValue).toBe(8)
})

it.each(['column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter', 'radar', 'combo'] as const)(
  'creates %s charts as one authoritative feature edit using year/X categories',
  async (chartType) => {
    const f = fixture()
    f.state.sheets[0]!.columns = 3
    f.state.sheets[0]!.rows = 4
    Object.assign(f.state.sheets[0]!.cells, {
      A4: { kind: 'number', rawValue: 2022 },
      B4: { kind: 'number', rawValue: 18 },
      C4: { kind: 'number', rawValue: 14 },
      A1: { kind: 'string', rawValue: 'Year' },
      A2: { kind: 'number', rawValue: 2020 },
      A3: { kind: 'number', rawValue: 2021 },
      C1: { kind: 'string', rawValue: 'Cost' },
      C2: { kind: 'number', rawValue: 4 },
      C3: { kind: 'number', rawValue: 6 },
    })
    const before = structuredClone(f.state.sheets[0]!.cells)
    await f.apply([{ op: 'add_chart', sheetId: f.sheetId, chartType, dataRange: 'A1:C4' }])
    const visual = (f.state.sheets[0]!.features!.charts as unknown as SheetVisual[])[0]!
    expect(chartKindFromVisual(visual.chart)).toBe(chartType)
    expect(visual.chart.series.map((series) => series.values)).toEqual([
      [12.5, 8, 18],
      [4, 6, 14],
    ])
    expect(visual.chart.series[0]!.categories).toEqual(['2020', '2021', '2022'])
    expect(f.state.sheets[0]!.cells).toEqual(before)
  },
)

it('moves/resizes original drawings and deletes original charts without mutating source metadata or cells', async () => {
  const f = fixture()
  const anchor = {
    fromRow: 0,
    fromColumn: 3,
    fromRowOffset: 0,
    fromColumnOffset: 0,
    toRow: 15,
    toColumn: 10,
    toRowOffset: 0,
    toColumnOffset: 0,
  }
  f.state.features = {
    originalMetadata: {
      visuals: [
        {
          id: 'original-chart',
          sheetId: 'base',
          kind: 'chart',
          drawingPath: 'xl/drawings/drawing1.xml',
          drawingIndex: 0,
          anchor,
        },
      ],
    },
  }
  const source = structuredClone(f.state.features)
  const cells = structuredClone(f.state.sheets[0]!.cells)
  const moved = {
    ...anchor,
    fromRow: 2,
    toRow: 19,
    fromColumnOffset: 95250,
    toColumnOffset: 190500,
  }
  await f.apply([
    { op: 'set_visual_anchor', sheetId: f.sheetId, visualId: 'original-chart', anchor: moved },
  ])
  expect(f.state.sheets[0]!.features?.visualEdits).toEqual({ 'original-chart': { anchor: moved } })
  expect(f.state.features).toEqual(source)
  expect(f.state.sheets[0]!.cells).toEqual(cells)
  await f.apply([{ op: 'delete_visual', visualId: 'original-chart' }])
  expect(f.state.sheets[0]!.features?.visualEdits).toEqual({ 'original-chart': { remove: true } })
  const before = structuredClone(f.state)
  await expect(
    f.apply([{ op: 'set_visual_anchor', sheetId: f.sheetId, visualId: 'original-chart', anchor }]),
  ).rejects.toThrow('visual_removed')
  expect(f.state).toEqual(before)
})
it('keeps original drawing permissions and ownership authoritative for movement and deletion', async () => {
  const f = fixture()
  const anchor = {
    fromRow: 0,
    fromColumn: 3,
    fromRowOffset: 0,
    fromColumnOffset: 0,
    toRow: 15,
    toColumn: 10,
    toRowOffset: 0,
    toColumnOffset: 0,
  }
  f.state.features = {
    originalMetadata: {
      visuals: [
        {
          id: 'original-chart',
          sheetId: 'base',
          kind: 'chart',
          drawingPath: 'xl/drawings/drawing1.xml',
          drawingIndex: 0,
          anchor,
        },
      ],
    },
  }
  f.state.sheets[0]!.protected = true
  const before = structuredClone(f.state)
  await expect(
    f.apply([{ op: 'set_visual_anchor', sheetId: f.sheetId, visualId: 'original-chart', anchor }]),
  ).rejects.toThrow('protected_target')
  await expect(f.apply([{ op: 'delete_visual', visualId: 'original-chart' }])).rejects.toThrow(
    'protected_target',
  )
  expect(f.state).toEqual(before)
})
it('moves a session chart without adding cell changes and rejects a bad anchor batch atomically', async () => {
  const f = fixture()
  await f.apply([{ op: 'add_chart', sheetId: f.sheetId, chartType: 'column', dataRange: 'A1:B3' }])
  const visual = (f.state.sheets[0]!.features!.charts as unknown as SheetVisual[])[0]!
  const anchor = { ...visual.anchor, fromColumn: 4, toColumn: 12 }
  const before = structuredClone(f.state)
  await f.apply([{ op: 'set_visual_anchor', sheetId: f.sheetId, visualId: visual.id, anchor }])
  expect((f.state.sheets[0]!.features!.charts as unknown as SheetVisual[])[0]!.anchor).toEqual(
    anchor,
  )
  expect(f.state.sheets[0]!.cells).toEqual(before.sheets[0]!.cells)
  const accepted = structuredClone(f.state)
  await expect(
    f.apply([
      { op: 'set_cell', sheetId: f.sheetId, address: 'B2', value: 99 },
      {
        op: 'set_visual_anchor',
        sheetId: f.sheetId,
        visualId: visual.id,
        anchor: { ...anchor, toColumn: 0 },
      },
    ]),
  ).rejects.toThrow()
  expect(f.state).toEqual(accepted)
})

it('edits authored pie point colors atomically without replacing cells or source series', async () => {
  const f = fixture()
  await f.apply([{ op: 'add_chart', sheetId: f.sheetId, chartType: 'pie', dataRange: 'A1:B3' }])
  const visual = (f.state.sheets[0]!.features!.charts as unknown as SheetVisual[])[0]!
  const before = structuredClone(f.state)
  await f.apply([
    {
      op: 'edit_chart',
      chartPath: visual.id,
      title: 'Regional share',
      pointColors: { '0': { '1': '#123456' } },
    },
  ])
  const chart = (f.state.sheets[0]!.features!.charts as unknown as SheetVisual[])[0]!.chart
  expect(chart.title).toBe('Regional share')
  expect(chart.series[0]!.pointColors).toEqual([{ index: 1, color: '#123456' }])
  expect(chart.series[0]!.values).toEqual(visual.chart.series[0]!.values)
  expect(f.state.sheets[0]!.cells).toEqual(before.sheets[0]!.cells)
  const accepted = structuredClone(f.state)
  await expect(
    f.apply([{ op: 'edit_chart', chartPath: visual.id, pointColors: { '0': { '0': 'invalid' } } }]),
  ).rejects.toThrow()
  expect(f.state).toEqual(accepted)
  await f.apply([
    { op: 'edit_chart', chartPath: visual.id, pointColors: { '0': { '0': '#ABCDEF' } } },
  ])
  expect(
    (f.state.sheets[0]!.features!.charts as unknown as SheetVisual[])[0]!.chart.series[0]!
      .pointColors,
  ).toEqual([
    { index: 0, color: '#ABCDEF' },
    { index: 1, color: '#123456' },
  ])
})

function originalChartFixture() {
  const f = fixture()
  const visual = buildChartVisual({
    id: 'original-chart',
    sheetId: 'base',
    sheetName: 'Data',
    chartType: 'column',
    dataRange: 'A1:B3',
    values: [
      ['Region', 'Sales'],
      ['North', 12.5],
      ['South', 8],
    ],
  })
  f.state.features = JSON.parse(
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
  return f
}
it('journals original chart properties and merges colors without rewriting source metadata or cells', async () => {
  const f = originalChartFixture()
  const before = structuredClone(f.state)
  await f.apply([
    {
      op: 'edit_chart',
      chartPath: 'xl/charts/chart1.xml',
      title: 'Original title',
      seriesColors: { '0': '#123456' },
      pointColors: { '0': { '0': '#ABCDEF' } },
    },
  ])
  await f.apply([
    {
      op: 'edit_chart',
      chartPath: 'xl/charts/chart1.xml',
      chartType: 'line',
      pointColors: { '0': { '1': '#456789' } },
    },
  ])
  expect(f.state.sheets[0]!.features!.chartEdits).toEqual({
    'xl/charts/chart1.xml': {
      chartPath: 'xl/charts/chart1.xml',
      title: 'Original title',
      chartType: 'line',
      seriesColors: { '0': '#123456' },
      pointColors: { '0': { '0': '#ABCDEF', '1': '#456789' } },
    },
  })
  expect(f.state.features).toEqual(before.features)
  expect(f.state.sheets[0]!.cells).toEqual(before.sheets[0]!.cells)
  const accepted = structuredClone(f.state)
  await expect(
    f.apply([
      { op: 'set_cell', sheetId: f.sheetId, address: 'B2', value: 99 },
      { op: 'edit_chart', chartPath: 'xl/charts/chart1.xml', seriesColors: { '7': '#123456' } },
    ]),
  ).rejects.toThrow('unknown_chart_series')
  expect(f.state).toEqual(accepted)
  await expect(
    f.apply([{ op: 'edit_chart', chartPath: 'xl/charts/unknown.xml', title: 'Unknown' }]),
  ).rejects.toThrow('unknown_chart')
  expect(f.state).toEqual(accepted)
})
it('rejects protected, removed, ambiguous and unsupported original chart edits atomically', async () => {
  const f = originalChartFixture()
  f.state.sheets[0]!.protected = true
  const op = { op: 'edit_chart' as const, chartPath: 'xl/charts/chart1.xml', title: 'Rejected' }
  await expect(f.apply([op])).rejects.toThrow('protected_target')
  f.state.sheets[0]!.protected = false
  await expect(
    f.apply([{ ...op, chartType: 'pie', axisTitles: { value: 'Invalid' } }]),
  ).rejects.toThrow('unsupported_chart_axes')
  await expect(f.apply([{ ...op, pointColors: { '0': { '9': '#123456' } } }])).rejects.toThrow(
    'unknown_chart_point',
  )
  await f.apply([{ op: 'delete_visual', visualId: 'original-chart' }])
  await expect(f.apply([op])).rejects.toThrow('visual_removed')
  const ambiguous = originalChartFixture()
  const visuals = (ambiguous.state.features!.originalMetadata as any).visuals
  visuals.push({ ...visuals[0], id: 'duplicate' })
  await expect(ambiguous.apply([op])).rejects.toThrow('unknown_chart')
})

it('repoints original series across sheets with authoritative caches and keeps journaled ranges live', async () => {
  const f = originalChartFixture()
  f.state.sheets.push({
    id: 'other',
    name: "Q!1's",
    rows: 2,
    columns: 1,
    cells: { A1: { kind: 'number', rawValue: 42 }, A2: { kind: 'number', rawValue: 64 } },
    features: {},
  })
  f.state.sheetMap.other = { createdName: "Q!1's" }
  const before = structuredClone(f.state)
  await f.apply([
    {
      op: 'edit_chart',
      chartPath: 'xl/charts/chart1.xml',
      seriesData: [
        { index: 0, sheetId: 'other', valuesRange: 'A1:A2', name: 'Repointed' },
        { index: 0, sheetId: f.sheetId, categoriesRange: 'A2:A3' },
      ],
    },
  ])
  const series = () =>
    (f.state.sheets[0]!.features!.chartEdits as any)['xl/charts/chart1.xml'].series[0]
  expect(series()).toEqual({
    index: 0,
    name: 'Repointed',
    values: [42, 64],
    valuesRef: "'Q!1''s'!$A$1:$A$2",
    categories: ['North', 'South'],
    categoriesRef: "'Data'!$A$2:$A$3",
  })
  expect(f.state.features).toEqual(before.features)
  expect(f.state.sheets.map((sheet) => sheet.cells)).toEqual(
    before.sheets.map((sheet) => sheet.cells),
  )
  await f.apply([{ op: 'set_cell', sheetId: 'other', address: 'A1', value: 99 }])
  expect(series().values).toEqual([99, 64])
  await f.apply([
    { op: 'edit_chart', chartPath: 'xl/charts/chart1.xml', title: 'Keep range edits' },
  ])
  expect(series().values).toEqual([99, 64])
})
it('rejects unsupported original chart vectors and pending formula caches without publishing partial writes', async () => {
  const f = originalChartFixture()
  const accepted = structuredClone(f.state)
  for (const [range, message] of [
    ['A1:B2', 'chart_vector_required'],
    ['A1:A1001', 'chart_vector_limit'],
  ] as const) {
    await expect(
      f.apply([
        { op: 'set_cell', sheetId: f.sheetId, address: 'B2', value: 99 },
        {
          op: 'edit_chart',
          chartPath: 'xl/charts/chart1.xml',
          seriesData: [{ index: 0, valuesRange: range }],
        },
      ]),
    ).rejects.toThrow(message)
    expect(f.state).toEqual(accepted)
  }
  f.state.sheets[0]!.cells.B2 = {
    kind: 'blank',
    rawValue: null,
    formula: '=1',
    calculationState: 'pending',
  }
  const pending = structuredClone(f.state)
  await expect(
    f.apply([
      {
        op: 'edit_chart',
        chartPath: 'xl/charts/chart1.xml',
        seriesData: [{ index: 0, valuesRange: 'B2:B3' }],
      },
    ]),
  ).rejects.toThrow('chart_formula_pending')
  expect(f.state).toEqual(pending)
})
it('persists new references when repointing a session chart so old ranges cannot overwrite its new cache', async () => {
  const f = fixture()
  f.state.sheets[0]!.cells.B3 = { kind: 'string', rawValue: '8' }
  await f.apply([{ op: 'add_chart', sheetId: f.sheetId, chartType: 'column', dataRange: 'A1:B3' }])
  const id = (f.state.sheets[0]!.features!.charts as unknown as SheetVisual[])[0]!.id
  ;(f.state.sheets[0]!.features!.charts as unknown as SheetVisual[])[0]!.chart.series[0]!.nameRef =
    "'Data'!$B$1"
  await f.apply([
    {
      op: 'edit_chart',
      chartPath: id,
      seriesData: [
        { index: 0, name: 'Renamed series', valuesRange: 'B3:B3', categoriesRange: 'A3:A3' },
      ],
    },
  ])
  const series = (f.state.sheets[0]!.features!.charts as unknown as SheetVisual[])[0]!.chart
    .series[0]!
  expect(series.name).toBe('Renamed series')
  expect(series.nameRef).toBeUndefined()
  expect(series.values).toEqual([8])
  expect(series.valuesRef).toBe("'Data'!$B$3:$B$3")
  expect(series.categoriesRef).toBe("'Data'!$A$3:$A$3")
})

it('applies complete chart format controls to original and session charts and rejects invalid combined bounds', async () => {
  for (const original of [true, false]) {
    const f = original ? originalChartFixture() : fixture()
    if (!original)
      await f.apply([
        { op: 'add_chart', sheetId: f.sheetId, chartType: 'column', dataRange: 'A1:B3' },
      ])
    const path = original
      ? 'xl/charts/chart1.xml'
      : (f.state.sheets[0]!.features!.charts as unknown as SheetVisual[])[0]!.id
    const cells = structuredClone(f.state.sheets[0]!.cells)
    await f.apply([
      {
        op: 'edit_chart',
        chartPath: path,
        legend: 'bottom',
        dataLabels: 'value',
        dataLabelPosition: 'inside-end',
        dataLabelFormat: '0.00',
        gridlines: false,
        valueAxis: { min: 0, max: 50 },
        gapWidthPct: 60,
      },
    ])
    const chart = () =>
      original
        ? (f.state.sheets[0]!.features!.chartEdits as any)[path]
        : (f.state.sheets[0]!.features!.charts as unknown as SheetVisual[])[0]!.chart
    expect(chart()).toMatchObject({
      legend: 'bottom',
      dataLabels: 'value',
      dataLabelPosition: 'inside-end',
      dataLabelFormat: '0.00',
      gridlines: false,
      valueAxis: { min: 0, max: 50 },
      gapWidthPct: 60,
    })
    const accepted = structuredClone(f.state)
    await expect(
      f.apply([{ op: 'edit_chart', chartPath: path, valueAxis: { min: 60 } }]),
    ).rejects.toThrow('invalid_chart_axis_bounds')
    expect(f.state).toEqual(accepted)
    await expect(f.apply([{ op: 'edit_chart', chartPath: path, holeSizePct: 60 }])).rejects.toThrow(
      'unsupported_chart_hole',
    )
    expect(f.state).toEqual(accepted)
    await f.apply([{ op: 'edit_chart', chartPath: path, valueAxis: { min: null } }])
    expect(chart().valueAxis.max).toBe(50)
    expect(original ? chart().valueAxis.min : (chart().valueAxis.min ?? null)).toBeNull()
    expect(f.state.sheets[0]!.cells).toEqual(cells)
    await f.apply([{ op: 'edit_chart', chartPath: path, chartType: 'pie' }])
    expect(chart().valueAxis).toBeUndefined()
    expect(chart().gridlines).toBeUndefined()
    expect(chart().gapWidthPct).toBeUndefined()
  }
})
it('retains pie and doughnut explosion settings and validates individual point targets', async () => {
  const f = fixture()
  await f.apply([
    { op: 'add_chart', sheetId: f.sheetId, chartType: 'doughnut', dataRange: 'A1:B3' },
  ])
  const id = (f.state.sheets[0]!.features!.charts as unknown as SheetVisual[])[0]!.id
  await f.apply([
    {
      op: 'edit_chart',
      chartPath: id,
      holeSizePct: 65,
      explosionPct: 25,
      pointExplosions: { '1': 40 },
    },
  ])
  const chart = (f.state.sheets[0]!.features!.charts as unknown as SheetVisual[])[0]!.chart
  expect(chart.holeSizePct).toBe(65)
  expect(chart.series[0]!.explosionPct).toBe(25)
  expect(chart.series[0]!.pointExplosions).toEqual([{ index: 1, pct: 40 }])
  const accepted = structuredClone(f.state)
  await expect(
    f.apply([{ op: 'edit_chart', chartPath: id, pointExplosions: { '8': 40 } }]),
  ).rejects.toThrow('unknown_chart_point')
  await expect(f.apply([{ op: 'edit_chart', chartPath: id, gridlines: true }])).rejects.toThrow(
    'unsupported_chart_axes',
  )
  expect(f.state).toEqual(accepted)
})

it('formats points of an uncached original chart from its trusted reference extent', async () => {
  const f = originalChartFixture()
  const chart = (f.state.features!.originalMetadata as any).visuals[0].chart
  chart.series[0].values = []
  chart.series[0].categories = []
  const before = structuredClone(f.state.features)
  await f.apply([
    {
      op: 'edit_chart',
      chartPath: 'xl/charts/chart1.xml',
      pointColors: { '0': { '1': '#123456' } },
    },
  ])
  expect(f.state.features).toEqual(before)
  expect(
    (f.state.sheets[0]!.features!.chartEdits as any)['xl/charts/chart1.xml'].pointColors,
  ).toEqual({ '0': { '1': '#123456' } })
  await expect(
    f.apply([
      {
        op: 'edit_chart',
        chartPath: 'xl/charts/chart1.xml',
        pointColors: { '0': { '2': '#123456' } },
      },
    ]),
  ).rejects.toThrow('unknown_chart_point')
})

for (const original of [true, false]) {
  it(`replaces ${original ? 'original' : 'session'} series using current references, preserves reordered colors and refreshes later edits`, async () => {
    const f = original ? originalChartFixture() : fixture()
    if (!original)
      await f.apply([
        { op: 'add_chart', sheetId: f.sheetId, chartType: 'column', dataRange: 'A1:B3' },
      ])
    const path = original
      ? 'xl/charts/chart1.xml'
      : (f.state.sheets[0]!.features!.charts as unknown as SheetVisual[])[0]!.id
    const chart = () =>
      original
        ? (f.state.sheets[0]!.features!.chartEdits as any)[path].seriesSet
        : (f.state.sheets[0]!.features!.charts as unknown as SheetVisual[])[0]!.chart.series
    await f.apply([
      {
        op: 'edit_chart',
        chartPath: path,
        seriesColors: { '0': '#123456' },
        pointColors: { '0': { '1': '#ABCDEF' } },
      },
    ])
    const metadata = structuredClone(f.state.features)
    await f.apply([
      { op: 'set_cell', sheetId: f.sheetId, address: 'B2', value: 42 },
      {
        op: 'edit_chart',
        chartPath: path,
        seriesSelection: {
          mode: 'rows',
          rows: [
            { sourceIndex: 0, name: 'Retained' },
            { name: 'Added', valuesRef: 'B2:B3', categoriesRef: 'A2:A3' },
          ],
        },
      },
    ])
    expect(chart()).toHaveLength(2)
    expect(chart()[0]).toMatchObject({ name: 'Retained', values: [42, 8], color: '#123456' })
    expect(chart()[0].pointColors).toBeUndefined()
    expect(chart()[1]).toMatchObject({
      name: 'Added',
      values: [42, 8],
      categories: ['North', 'South'],
    })
    if (original) {
      expect((f.state.sheets[0]!.features!.chartEdits as any)[path].pointColors).toBeUndefined()
      expect(f.state.features).toEqual(metadata)
    }
    await f.apply([
      {
        op: 'edit_chart',
        chartPath: path,
        seriesSelection: {
          mode: 'rows',
          rows: [
            { sourceIndex: 1, name: 'First' },
            { sourceIndex: 0, name: 'Second' },
          ],
        },
      },
    ])
    expect(chart()[1].color).toBe('#123456')
    await f.apply([
      {
        op: 'edit_chart',
        chartPath: path,
        seriesData: [{ index: 1, name: 'Inline after replace', valuesRange: 'B3' }],
      },
    ])
    await f.apply([{ op: 'set_cell', sheetId: f.sheetId, address: 'B3', value: 99 }])
    expect(chart()[0].values).toEqual([42, 99])
    expect(chart()[1]).toMatchObject({ name: 'Inline after replace', values: [99] })
    await f.apply([{ op: 'edit_chart', chartPath: path, seriesSelection: { mode: 'transpose' } }])
    expect(chart()[0]).toMatchObject({
      name: 'North',
      values: [42, 99],
      categories: ['First', 'Inline after replace'],
    })
    expect(chart()[0].valuesRef).toBeUndefined()
    await f.apply([{ op: 'set_cell', sheetId: f.sheetId, address: 'B2', value: 100 }])
    expect(chart()[0].values).toEqual([42, 99])
  })
}
it('rejects invalid series replacement atomically and never accepts browser caches', async () => {
  const f = originalChartFixture(),
    path = 'xl/charts/chart1.xml'
  for (const seriesSelection of [
    { mode: 'rows', rows: [{ sourceIndex: 10, name: 'Missing' }] },
    { mode: 'rows', rows: [{ name: 'Missing values' }] },
    { mode: 'rows', rows: [] },
    { mode: 'rows', rows: [{ name: 'Forged', values: [999], valuesRef: 'B2:B3' }] },
    { mode: 'rows', rows: [{ name: 'Matrix', valuesRef: 'A1:B3' }] },
  ]) {
    const before = structuredClone(f.state)
    await expect(
      f.apply([
        { op: 'set_cell', sheetId: f.sheetId, address: 'B2', value: 999 },
        { op: 'edit_chart', chartPath: path, seriesSelection } as any,
      ]),
    ).rejects.toThrow()
    expect(f.state).toEqual(before)
  }
  f.state.sheets[0]!.cells.B2 = {
    kind: 'blank',
    rawValue: null,
    formula: '=1+1',
    cachedValue: 2,
    calculationState: 'pending',
  }
  await expect(
    f.apply([
      {
        op: 'edit_chart',
        chartPath: path,
        seriesSelection: { mode: 'rows', rows: [{ name: 'Pending', valuesRef: 'B2:B3' }] },
      },
    ]),
  ).rejects.toThrow('chart_formula_pending')
})

for (const original of [true, false]) {
  it(`preserves ${original ? 'source' : 'session'} visual metadata through geometry changes and reorders layers atomically`, async () => {
    const f = original ? originalChartFixture() : fixture()
    if (original) {
      const metadata = (f.state.features!.originalMetadata as any).visuals
      const template = structuredClone(metadata[0])
      template.editAs = 'twoCell'
      metadata.splice(
        0,
        1,
        ...[0, 1, 2].map((i) => ({
          ...structuredClone(template),
          id: `source-${i}`,
          chartPath: `xl/charts/chart${i + 1}.xml`,
          drawingIndex: i,
        })),
      )
    } else
      await f.apply(
        [0, 1, 2].map((i) => ({
          op: 'add_chart',
          sheetId: f.sheetId,
          chartType: 'column',
          dataRange: 'A1:B3',
          title: `Chart ${i}`,
        })),
      )
    const originals = structuredClone(f.state.features)
    const ids = original
      ? (f.state.features!.originalMetadata as any).visuals.map((v: any) => v.id)
      : (f.state.sheets[0]!.features!.charts as unknown as SheetVisual[]).map((v) => v.id)
    const anchor = original
      ? (f.state.features!.originalMetadata as any).visuals[0].anchor
      : (f.state.sheets[0]!.features!.charts as unknown as SheetVisual[])[0]!.anchor
    const moved = { ...anchor, fromRow: anchor.fromRow + 1, toRow: anchor.toRow + 1 }
    await f.apply([
      { op: 'edit_visual', visualId: ids[0], altText: 'Sales & regions', editAs: 'absolute' },
      { op: 'set_visual_anchor', sheetId: f.sheetId, visualId: ids[0], anchor: moved },
    ])
    const props = () =>
      original
        ? (f.state.sheets[0]!.features!.visualEdits as any)[ids[0]]
        : (f.state.sheets[0]!.features!.charts as any[]).find((v) => v.id === ids[0])
    expect(props()).toMatchObject({ altText: 'Sales & regions', editAs: 'absolute', anchor: moved })
    await f.apply([{ op: 'edit_visual', visualId: ids[0], zIndex: 2 }])
    if (original) {
      const edits = f.state.sheets[0]!.features!.visualEdits as any
      expect(ids.map((id: string) => edits[id].zIndex)).toEqual([2, 0, 1])
      expect(f.state.features).toEqual(originals)
    } else
      expect((f.state.sheets[0]!.features!.charts as any[]).map((v) => v.id)).toEqual([
        ids[1],
        ids[2],
        ids[0],
      ])
    expect(props()).toMatchObject({ altText: 'Sales & regions', editAs: 'absolute', anchor: moved })
    await f.apply([
      { op: 'edit_visual', visualId: ids[0], altText: '', editAs: 'twoCell', zIndex: 0 },
    ])
    expect(props()).toMatchObject({ altText: '', editAs: 'twoCell', anchor: moved })
    const before = structuredClone(f.state)
    await expect(
      f.apply([
        { op: 'edit_visual', visualId: ids[0], altText: 'Must roll back' },
        { op: 'edit_visual', visualId: ids[1], zIndex: 100 },
      ]),
    ).rejects.toThrow('invalid_visual_z_index')
    expect(f.state).toEqual(before)
  })
}
it('rejects protected, removed, unknown and unsupported visual properties atomically', async () => {
  const f = originalChartFixture(),
    visual = (f.state.features!.originalMetadata as any).visuals[0]
  const before = structuredClone(f.state)
  await expect(
    f.apply([{ op: 'edit_visual', visualId: visual.id, editAs: 'oneCell' }]),
  ).rejects.toThrow('unsupported_visual_placement')
  await expect(f.apply([{ op: 'edit_visual', visualId: 'missing', altText: 'x' }])).rejects.toThrow(
    'unknown_visual',
  )
  await expect(
    f.apply([{ op: 'edit_visual', visualId: visual.id, altText: 'x'.repeat(2001) }]),
  ).rejects.toThrow()
  expect(f.state).toEqual(before)
  f.state.sheets[0]!.protected = true
  await expect(f.apply([{ op: 'edit_visual', visualId: visual.id, altText: 'x' }])).rejects.toThrow(
    'protected_target',
  )
  f.state.sheets[0]!.protected = false
  await f.apply([{ op: 'delete_visual', visualId: visual.id }])
  await expect(f.apply([{ op: 'edit_visual', visualId: visual.id, altText: 'x' }])).rejects.toThrow(
    'visual_removed',
  )
})

it('composes source shape paint, frame transforms, anchor and flip edits without altering chart data', async () => {
  const f = originalChartFixture(),
    visual = (f.state.features!.originalMetadata as any).visuals[0]
  visual.kind = 'shape'
  delete visual.chart
  delete visual.chartPath
  visual.editAs = 'twoCell'
  visual.shapeType = 'rect'
  visual.fillColor = '#ABCDEF'
  const before = structuredClone(f.state.features),
    anchor = structuredClone(visual.anchor)
  await f.apply([
    { op: 'set_visual_anchor', sheetId: f.sheetId, visualId: visual.id, anchor },
    {
      op: 'edit_visual',
      visualId: visual.id,
      fillColor: '#112233',
      lineColor: 'none',
      rotation: 90,
      flipH: true,
      frameSize: { width: 1000000, height: 500000 },
    },
  ])
  expect((f.state.sheets[0]!.features!.visualEdits as any)[visual.id]).toMatchObject({
    anchor,
    fillColor: '#112233',
    lineColor: 'none',
    rotation: 90,
    flipH: true,
    frameSize: { width: 1000000, height: 500000 },
  })
  expect(f.state.features).toEqual(before)
  const root = structuredClone(f.state)
  await expect(
    f.apply([{ op: 'edit_visual', visualId: visual.id, fillColor: 'bad' } as any]),
  ).rejects.toThrow()
  expect(f.state).toEqual(root)
})
it('rejects unsupported source paint/transforms and frame changes without an anchor atomically', async () => {
  const f = originalChartFixture(),
    visual = (f.state.features!.originalMetadata as any).visuals[0]
  for (const patch of [
    { fillColor: '#112233' },
    { rotation: 90 },
    { flipH: true },
    { frameSize: { width: 100, height: 100 } },
  ])
    await expect(
      f.apply([{ op: 'edit_visual', visualId: visual.id, ...patch } as any]),
    ).rejects.toThrow('unsupported_visual')
  visual.kind = 'image'
  delete visual.chart
  delete visual.chartPath
  await expect(
    f.apply([{ op: 'edit_visual', visualId: visual.id, frameSize: { width: 100, height: 100 } }]),
  ).rejects.toThrow('visual_frame_requires_anchor')
})

it('adds and edits session shapes alongside chart layers with exact inverse roots', async () => {
  const f = fixture()
  await f.apply([
    { op: 'add_chart', sheetId: f.sheetId, chartType: 'column', dataRange: 'A1:B3' },
    { op: 'add_shape', sheetId: f.sheetId, shapeType: 'rect', anchorCell: 'D4', text: 'First' },
  ])
  const visuals = () => f.state.sheets[0]!.features!.visuals as any[]
  const chart = visuals()[0],
    shape = visuals()[1]
  expect(shape.id).toMatch(/^added-shape-/)
  expect(shape.anchor).toMatchObject({ fromRow: 3, toRow: 9, fromColumn: 3, toColumn: 6 })
  expect((f.state.sheets[0]!.features!.charts as any[]).map((v) => v.id)).toEqual([chart.id])
  await f.apply([
    { op: 'edit_shape', visualId: shape.id, text: 'Edited & saved', fillColor: '#112233' },
    {
      op: 'edit_visual',
      visualId: shape.id,
      flipH: true,
      rotation: 90,
      lineColor: 'none',
      altText: 'Description',
      zIndex: 0,
    },
  ])
  expect(visuals().map((v) => v.id)).toEqual([shape.id, chart.id])
  expect(visuals()[0]).toMatchObject({
    text: 'Edited & saved',
    fillColor: '#112233',
    flipH: true,
    rotation: 90,
    lineColor: 'none',
    altText: 'Description',
  })
  await f.apply([
    { op: 'delete_visual', visualId: shape.id },
    { op: 'add_shape', sheetId: f.sheetId, shapeType: 'ellipse', anchorCell: 'D4' },
  ])
  expect(visuals()[1].id).not.toBe(shape.id)
  await f.apply([{ op: 'add_shape', sheetId: f.sheetId, shapeType: 'textbox', anchorCell: 'F5' }])
  expect(visuals().at(-1)).toMatchObject({
    kind: 'shape',
    shapeType: 'rect',
    isTextBox: true,
    text: 'Text',
    fillColor: '#FFFFFF',
    anchor: { fromRow: 4, toRow: 8, fromColumn: 5, toColumn: 9 },
  })
})
it('validates full drawn anchors and rejects shape edits to file objects atomically', async () => {
  const f = fixture(),
    before = structuredClone(f.state)
  await expect(
    f.apply([
      {
        op: 'add_shape',
        sheetId: f.sheetId,
        shapeType: 'rect',
        anchorCell: 'A1',
        anchor: {
          fromRow: 0,
          fromColumn: 0,
          fromRowOffset: 0,
          fromColumnOffset: 0,
          toRow: 0,
          toColumn: 0,
          toRowOffset: 0,
          toColumnOffset: 0,
        },
      },
    ]),
  ).rejects.toThrow()
  expect(f.state).toEqual(before)
  f.state.sheets[0]!.protected = true
  await expect(
    f.apply([{ op: 'add_shape', sheetId: f.sheetId, shapeType: 'rect', anchorCell: 'A1' }]),
  ).rejects.toThrow('protected_target')
  const source = originalChartFixture()
  await expect(
    source.apply([{ op: 'edit_shape', visualId: 'original-chart', text: 'forged' }]),
  ).rejects.toThrow('unknown_shape')
})

it('adds inline pictures with stable identities and canonical transforms, order and inverses', async () => {
  const f = fixture(),
    base64 =
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII='
  const picture = {
    op: 'add_picture' as const,
    sheetId: f.sheetId,
    name: 'pixel.png',
    mediaType: 'image/png' as const,
    base64,
    anchor: {
      fromRow: 2,
      fromColumn: 3,
      fromRowOffset: 0,
      fromColumnOffset: 0,
      toRow: 4,
      toColumn: 5,
      toRowOffset: 0,
      toColumnOffset: 0,
    },
  }
  await f.apply([
    picture,
    { op: 'add_shape', sheetId: f.sheetId, shapeType: 'rect', anchorCell: 'B5' },
  ])
  const image = (f.state.sheets[0]!.features!.visuals as any[])[0]
  expect(image).toMatchObject({
    kind: 'image',
    mediaDataUrl: `data:image/png;base64,${base64}`,
    name: 'pixel.png',
  })
  expect(image.id).toMatch(/^added-image-/)
  await f.apply([
    {
      op: 'edit_visual',
      visualId: image.id,
      altText: 'Manual picture',
      rotation: 90,
      flipH: true,
      frameSize: { width: 100, height: 80 },
      zIndex: 1,
    },
  ])
  expect((f.state.sheets[0]!.features!.visuals as any[])[1]).toMatchObject({
    id: image.id,
    rotation: 90,
    flipH: true,
    frameWidth: 100,
    frameHeight: 80,
  })
  const before = structuredClone(f.state)
  await expect(
    f.apply([{ op: 'edit_visual', visualId: image.id, fillColor: '#FF0000' }]),
  ).rejects.toThrow('unsupported_visual_properties')
  await expect(f.apply([{ ...picture, mediaType: 'image/jpeg' }])).rejects.toThrow()
  await expect(f.apply([{ ...picture, base64: 'iVBORw0KGgo===!' }])).rejects.toThrow()
  expect(f.state).toEqual(before)
  await f.apply([{ op: 'delete_visual', visualId: image.id }, picture])
  const replacement = (f.state.sheets[0]!.features!.visuals as any[]).find(
    (v) => v.kind === 'image',
  )
  expect(replacement.id).not.toBe(image.id)
  f.state.sheets[0]!.protected = true
  await expect(f.apply([picture])).rejects.toThrow('protected_target')
})

it('pastes typed snapshots with literal formulas, errors, dates and replacement formats as one reversible block', async () => {
  const f = fixture()
  f.state.dateSystem = '1900'
  await f.apply([
    {
      op: 'format_range',
      sheetId: f.sheetId,
      range: 'D5:G5',
      format: { bold: true, border: { type: 'all', style: 'thick', color: '#FF0000' } },
    },
  ])
  const before = structuredClone(f.state)
  await f.apply([
    {
      op: 'paste_cells',
      sheetId: f.sheetId,
      start: 'D5',
      cells: [
        [
          {
            kind: 'string',
            rawValue: '=literal',
            style: { borderLeft: { style: 'mediumDashDotDot', color: '#123456' } },
          },
          { kind: 'error', rawValue: '#DIV/0!' },
          { kind: 'date', rawValue: 45000, dateSystem: '1900', numberFormat: 'yyyy-mm-dd' },
          {
            kind: 'number',
            rawValue: 12,
            formula: '=SUM(A1:A3)',
            calculationState: 'ready',
            cachedValue: 12,
          },
        ],
      ],
    },
  ])
  const cells = f.state.sheets[0]!.cells
  expect(cells.D5).toMatchObject({
    kind: 'string',
    rawValue: '=literal',
    styleReset: true,
    style: { borderLeft: { style: 'mediumDashDotDot', color: '#123456' } },
  })
  expect(cells.D5!.formula).toBeUndefined()
  expect(cells.D5!.style?.bold).toBeUndefined()
  expect(cells.D5!.style?.borderTop).toBeUndefined()
  expect(cells.E5).toMatchObject({ kind: 'error', rawValue: '#DIV/0!', styleReset: true })
  expect(cells.F5).toMatchObject({
    kind: 'date',
    rawValue: 45000,
    dateSystem: '1900',
    numberFormat: 'yyyy-mm-dd',
  })
  expect(cells.G5).toMatchObject({
    formula: '=SUM(A1:A3)',
    calculationState: 'pending',
    rawValue: null,
  })
  expect(cells.G5!.cachedValue).toBeUndefined()
  expect(before.sheets[0]!.cells.D5!.style!.bold).toBe(true)
  const current = structuredClone(f.state)
  await expect(
    f.apply([
      {
        op: 'paste_cells',
        sheetId: f.sheetId,
        start: 'XFD1048576',
        cells: [
          [
            { kind: 'number', rawValue: 1 },
            { kind: 'number', rawValue: 2 },
          ],
        ],
      },
    ]),
  ).rejects.toThrow()
  await expect(
    f.apply([
      {
        op: 'paste_cells',
        sheetId: f.sheetId,
        start: 'A1',
        cells: [[{ kind: 'number', rawValue: 1 }], []],
      },
    ]),
  ).rejects.toThrow()
  expect(f.state).toEqual(current)
})

it('replaces contained destination merges atomically and rejects partial merges, protected targets and mixed date systems', async () => {
  const f = fixture()
  f.state.sheets[0]!.features = {
    merges: [
      { startRow: 4, endRow: 4, startColumn: 3, endColumn: 4 },
      { startRow: 8, endRow: 8, startColumn: 0, endColumn: 1 },
    ],
  }
  const before = structuredClone(f.state)
  await expect(
    f.apply([
      {
        op: 'paste_cells',
        sheetId: f.sheetId,
        start: 'E5',
        cells: [[{ kind: 'number', rawValue: 1 }]],
        merges: [],
      },
    ]),
  ).rejects.toThrow('merge_partial_overlap')
  expect(f.state).toEqual(before)
  await f.apply([
    {
      op: 'paste_cells',
      sheetId: f.sheetId,
      start: 'D5',
      cells: [
        [
          { kind: 'string', rawValue: 'Merged' },
          { kind: 'blank', rawValue: null },
        ],
      ],
      merges: ['D5:E5'],
    },
  ])
  expect(f.state.sheets[0]!.features!.mergeEdits).toEqual([
    { startRow: 8, endRow: 8, startColumn: 0, endColumn: 1 },
    { startRow: 4, endRow: 4, startColumn: 3, endColumn: 4 },
  ])
  await f.apply([
    {
      op: 'paste_cells',
      sheetId: f.sheetId,
      start: 'D5',
      cells: [
        [
          { kind: 'number', rawValue: 1 },
          { kind: 'number', rawValue: 2 },
        ],
      ],
      merges: [],
    },
  ])
  expect(f.state.sheets[0]!.features!.mergeEdits).toEqual([
    { startRow: 8, endRow: 8, startColumn: 0, endColumn: 1 },
  ])
  const current = structuredClone(f.state)
  await expect(
    f.apply([
      {
        op: 'paste_cells',
        sheetId: f.sheetId,
        start: 'D5',
        cells: [[{ kind: 'date', rawValue: 1, dateSystem: '1904' }]],
      },
    ]),
  ).rejects.toThrow('date_system_mismatch')
  expect(f.state).toEqual(current)
  f.state.sheets[0]!.protected = true
  await expect(
    f.apply([
      {
        op: 'paste_cells',
        sheetId: f.sheetId,
        start: 'D5',
        cells: [[{ kind: 'number', rawValue: 1 }]],
      },
    ]),
  ).rejects.toThrow('protected_target')
})

it('keeps existing chart caches unchanged when only print metadata changes', async () => {
  const f = fixture()
  const chart = buildChartVisual({
    id: 'page-chart',
    sheetId: f.sheetId,
    sheetName: 'Data',
    chartType: 'column',
    dataRange: 'B2:B3',
    values: [[12.5], [8]],
  })
  f.state.sheets[0]!.features = JSON.parse(JSON.stringify({ visuals: [chart], charts: [chart] }))
  f.state.sheets[0]!.cells.B2 = {
    kind: 'blank',
    rawValue: null,
    formula: '=1+1',
    calculationState: 'pending',
  }
  const before = structuredClone(f.state.sheets[0]!.features!)
  await f.apply([{ op: 'set_page_setup', sheetId: f.sheetId, orientation: 'landscape' }])
  expect(f.state.sheets[0]!.features!.visuals).toEqual(before.visuals)
  expect(f.state.sheets[0]!.features!.charts).toEqual(before.charts)
})

it('retains journaled original-chart caches when print settings change with pending formulas', async () => {
  const f = originalChartFixture()
  const chartEdits = {
    'xl/charts/chart1.xml': {
      series: [{ index: 0, valuesRef: "'Data'!$B$2:$B$3", values: [12.5, 8] }],
    },
  }
  f.state.sheets[0]!.features = { chartEdits: structuredClone(chartEdits) }
  f.state.sheets[0]!.cells.B2 = {
    kind: 'blank',
    rawValue: null,
    formula: '=1+1',
    calculationState: 'pending',
  }
  await f.apply([{ op: 'set_page_setup', sheetId: f.sheetId, paperSize: 8 }])
  expect(f.state.sheets[0]!.features!.chartEdits).toEqual(chartEdits)
})

it('journals workbook and sheet-scoped names atomically without changing source metadata or cell caches', async () => {
  const f = fixture()
  f.state.features = {
    originalMetadata: { definedNames: [{ name: 'Source', formula: 'Data!$B$2' }] },
  }
  const metadata = structuredClone(f.state.features.originalMetadata)
  const cells = structuredClone(f.state.sheets[0]!.cells)
  await f.apply([
    { op: 'add_defined_name', name: 'Total', ref: '=SUM(Data!B2:B3)' },
    { op: 'add_defined_name', name: 'Total', ref: 'Data!$B$3', scopeSheetId: f.sheetId },
  ])
  expect(f.state.features!.definedNamesEdits).toEqual({
    names: [
      { name: 'Source', formula: 'Data!$B$2' },
      { name: 'Total', formula: 'SUM(Data!B2:B3)' },
      { name: 'Total', formula: 'Data!$B$3', scopeSheetId: f.sheetId },
    ],
    preserveNames: [],
  })
  expect(f.state.features!.originalMetadata).toEqual(metadata)
  expect(f.state.sheets[0]!.cells).toEqual(cells)
  await f.apply([
    { op: 'delete_defined_name', name: 'total', scopeSheetId: f.sheetId },
    { op: 'add_defined_name', name: 'Scoped', ref: 'Data!$B$2', scopeSheetId: f.sheetId },
  ])
  const before = structuredClone(f.state)
  await expect(
    f.apply([
      { op: 'delete_defined_name', name: 'Total' },
      { op: 'add_defined_name', name: 'source', ref: '1' },
    ]),
  ).rejects.toThrow('duplicate_defined_name')
  expect(f.state).toEqual(before)
  await expect(
    f.apply([{ op: 'add_defined_name', name: 'Bad', ref: '1', scopeSheetId: 'missing' }]),
  ).rejects.toThrow('unknown_name_scope')
  await expect(f.apply([{ op: 'add_defined_name', name: 'A1', ref: '1' }])).rejects.toThrow(
    'invalid_defined_name',
  )
})

it('rejects names in protected workbooks and guards coordinate changes after new names', async () => {
  const f = fixture()
  await f.apply([{ op: 'add_defined_name', name: 'NewName', ref: 'Data!$B$2' }])
  await expect(
    f.apply([{ op: 'insert_rows', sheetId: f.sheetId, row: 1, count: 1 }]),
  ).rejects.toThrow('definedNames')
  await expect(
    f.apply([
      {
        op: 'relocate_range',
        sheetId: f.sheetId,
        sourceSheetId: f.sheetId,
        source: 'B2:B2',
        start: 'B3',
      },
    ]),
  ).rejects.toThrow('definedNames')
  f.state.features!.originalMetadata = { workbookProtection: { protected: true } }
  await expect(f.apply([{ op: 'delete_defined_name', name: 'NewName' }])).rejects.toThrow(
    'protected_workbook',
  )
})

it('commits reference-aware name changes without reapplying older renames to later formulas', async () => {
  const f = fixture()
  f.state.features = {
    originalMetadata: {
      nameBindingsComplete: true,
      definedNames: [{ name: 'Revenue', formula: 'Data!$B$2' }],
    },
  }
  f.state.sheets[0]!.cells.C2 = {
    kind: 'blank',
    rawValue: null,
    formula: '=Revenue+1',
    calculationState: 'pending',
  }
  await f.apply([
    {
      op: 'manage_defined_name',
      name: 'Revenue',
      replacement: { name: 'Sales', ref: 'Data!$B$2' },
    },
  ])
  expect(f.state.sheets[0]!.cells.C2!.formula).toBe('=Sales+1')
  await f.apply([{ op: 'set_formula', sheetId: f.sheetId, address: 'C3', formula: '=Revenue+2' }])
  await f.apply([{ op: 'add_defined_name', name: 'Revenue', ref: '1' }])
  await f.apply([{ op: 'manage_defined_name', name: 'Sales', replacement: null }])
  expect(f.state.sheets[0]!.cells.C2!.formula).toBe('=#REF!+1')
  expect(f.state.sheets[0]!.cells.C3!.formula).toBe('=Revenue+2')
  expect(f.state.features!.definedNameReferences).toEqual([
    { name: 'Revenue', replacement: 'Sales', shadowedSheetIds: [] },
    { name: 'Sales', replacement: null, shadowedSheetIds: [] },
  ])
  const before = structuredClone(f.state)
  await expect(
    f.apply([
      { op: 'manage_defined_name', name: 'Revenue', replacement: { name: 'A1', ref: '1' } },
    ]),
  ).rejects.toThrow('invalid_defined_name')
  expect(f.state).toEqual(before)
})

it('keeps structural guards after the last visible name is removed and for file-only bindings', async () => {
  for (const feature of ['definedNameReferences', 'reservedDefinedNames'] as const) {
    const f = fixture()
    f.state.features ??= {}
    if (feature === 'definedNameReferences')
      f.state.features.definedNameReferences = [
        { name: 'Removed', replacement: null, shadowedSheetIds: [] },
      ]
    else
      f.state.features.originalMetadata = {
        nameBindingsComplete: true,
        reservedDefinedNames: [{ name: 'Hidden', sheetIndex: 0, sheetId: 'base' }],
      }
    const before = structuredClone(f.state)
    for (const operation of [
      { op: 'insert_rows', sheetId: f.sheetId, row: 1, count: 1 },
      { op: 'rename_sheet', sheetId: f.sheetId, name: 'Renamed' },
    ] satisfies WorkbookOperation[])
      await expect(f.apply([operation])).rejects.toThrow(
        'unsupported_structural_feature: definedNames',
      )
    expect(f.state).toEqual(before)
  }
})

it('composes custom AND/OR conditions with other columns and keeps criteria out of cell history', async () => {
  const f = fixture(),
    cells = structuredClone(f.state.sheets[0]!.cells)
  await f.apply([{ op: 'set_filter', sheetId: f.sheetId, range: 'A1:B3' }])
  await f.apply([
    {
      op: 'set_filter_criteria',
      sheetId: f.sheetId,
      column: 'B',
      customs: {
        and: true,
        filters: [
          { operator: 'greaterThanOrEqual', val: '10' },
          { operator: 'lessThan', val: '20' },
        ],
      },
    },
  ])
  expect(f.state.sheets[0]!.features?.filterEdits).toMatchObject({
    hiddenRows: [2],
    filter: { columns: [{ colId: 1, customs: { and: true } }] },
  })
  await f.apply([{ op: 'set_filter_criteria', sheetId: f.sheetId, column: 'A', values: ['South'] }])
  expect(f.state.sheets[0]!.features?.filterEdits).toMatchObject({ hiddenRows: [1, 2] })
  await f.apply([{ op: 'set_filter_criteria', sheetId: f.sheetId, column: 'B', values: null }])
  expect(f.state.sheets[0]!.features?.filterEdits).toMatchObject({
    hiddenRows: [1],
    filter: { columns: [{ colId: 0, values: ['South'] }] },
  })
  expect(f.state.sheets[0]!.cells).toEqual(cells)
})

it('rejects missing formula results before custom filtering or content/filter mixed evaluation', async () => {
  const f = fixture()
  f.state.sheets[0]!.cells.B2 = {
    kind: 'number',
    rawValue: 12.5,
    formula: '=10+2.5',
    calculationState: 'pending',
  }
  await f.apply([{ op: 'set_filter', sheetId: f.sheetId, range: 'A1:B3' }])
  const operation = {
    op: 'set_filter_criteria' as const,
    sheetId: f.sheetId,
    column: 'B',
    customs: { filters: [{ operator: 'greaterThan' as const, val: '10' }] },
  }
  await expect(f.apply([operation])).rejects.toThrow('filter_formula_pending')
  f.state.sheets[0]!.cells.B2 = {
    ...f.state.sheets[0]!.cells.B2!,
    cachedValue: 12.5,
    calculationState: 'ready',
  }
  await expect(
    f.apply([{ op: 'set_cell', sheetId: f.sheetId, address: 'A2', value: 100 }, operation]),
  ).rejects.toThrow('filter_formula_pending')
  await f.apply([operation])
  expect(f.state.sheets[0]!.features?.filterEdits).toMatchObject({ hiddenRows: [2] })
})

it('clears criteria without removing the filter and reapplies unchanged definitions to edited data', async () => {
  const f = fixture()
  await f.apply([
    { op: 'set_filter', sheetId: f.sheetId, range: 'A1:B3' },
    {
      op: 'set_filter_criteria',
      sheetId: f.sheetId,
      column: 'B',
      customs: { filters: [{ operator: 'greaterThan', val: 10 }] },
    },
  ])
  expect(f.state.sheets[0]!.features?.filterEdits).toMatchObject({ hiddenRows: [2] })
  await f.apply([{ op: 'set_cell', sheetId: f.sheetId, address: 'B3', value: 20 }])
  expect(f.state.sheets[0]!.features?.filterEdits).toMatchObject({ hiddenRows: [2] })
  const criteria = structuredClone((f.state.sheets[0]!.features!.filterEdits as any).filter)
  await f.apply([{ op: 'reapply_filter', sheetId: f.sheetId }])
  expect(f.state.sheets[0]!.features?.filterEdits).toMatchObject({
    hiddenRows: [],
    filter: criteria,
  })
  await f.apply([{ op: 'clear_filter_criteria', sheetId: f.sheetId }])
  expect(f.state.sheets[0]!.features?.filterEdits).toMatchObject({
    hiddenRows: [],
    filter: { range: criteria.range, columns: [] },
  })
  await f.apply([{ op: 'clear_filter', sheetId: f.sheetId }])
  expect(f.state.sheets[0]!.features?.filterEdits).toMatchObject({ filter: null })
})

it('keeps filtered and manually hidden rows stationary while moving typed payloads and styles across visible slots', async () => {
  const f = fixture(),
    sheet = f.state.sheets[0]!
  sheet.rows = 6
  sheet.cells = {
    A1: { kind: 'string', rawValue: 'Group' },
    B1: { kind: 'string', rawValue: 'Amount' },
    A2: { kind: 'string', rawValue: 'Keep' },
    B2: { kind: 'number', rawValue: 30, style: { bold: true } },
    A3: { kind: 'string', rawValue: 'Drop' },
    B3: { kind: 'number', rawValue: 999, style: { italic: true } },
    A4: { kind: 'string', rawValue: 'Keep' },
    B4: { kind: 'number', rawValue: 10, style: { fillColor: '#FF0000' } },
    A5: { kind: 'string', rawValue: 'Keep' },
    B5: { kind: 'number', rawValue: 5, style: { underline: true } },
    A6: { kind: 'string', rawValue: 'Keep' },
    B6: { kind: 'number', rawValue: 20, numberFormat: '0.00', style: { numberFormat: '0.00' } },
  }
  await f.apply([
    { op: 'set_filter', sheetId: f.sheetId, range: 'A1:B6' },
    { op: 'set_filter_criteria', sheetId: f.sheetId, column: 'A', values: ['Keep'] },
    { op: 'set_rows_hidden', sheetId: f.sheetId, row: 5, count: 1, hidden: true },
  ])
  const before = structuredClone(f.state.sheets[0]!)
  await f.apply([
    {
      op: 'sort_range',
      sheetId: f.sheetId,
      range: 'A1:B6',
      byColumn: 'B',
      order: 'asc',
      hasHeader: true,
    },
  ])
  const sorted = f.state.sheets[0]!
  expect(['B2', 'B3', 'B4', 'B5', 'B6'].map((address) => sorted.cells[address]!.rawValue)).toEqual([
    10, 999, 20, 5, 30,
  ])
  expect(sorted.cells.B2!.style!.fillColor).toBe('#FF0000')
  expect(sorted.cells.B4!.numberFormat).toBe('0.00')
  expect(sorted.cells.B6!.style!.bold).toBe(true)
  expect(sorted.cells.B3).toEqual(before.cells.B3)
  expect(sorted.cells.B5).toEqual(before.cells.B5)
  expect(sorted.features).toEqual(before.features)
  await f.apply(
    [
      { op: 'clear_filter_criteria', sheetId: f.sheetId },
      {
        op: 'sort_range',
        sheetId: f.sheetId,
        range: 'A1:B6',
        byColumn: 'B',
        order: 'asc',
        hasHeader: true,
      },
    ],
    [2, 4],
  )
  expect(
    ['B2', 'B3', 'B4', 'B5', 'B6'].map((address) => f.state.sheets[0]!.cells[address]!.rawValue),
  ).toEqual([10, 20, 30, 5, 999])
  expect(f.state.sheets[0]!.features!.filterEdits).toMatchObject({
    hiddenRows: [],
    filter: { columns: [] },
  })
})

it('combines color and text columns while color criteria do not require formula values to be ready', async () => {
  const f = fixture()
  f.state.sheets[0]!.cells.A2!.style = { fillColor: '#FF0000' }
  f.state.sheets[0]!.cells.A3 = {
    kind: 'blank',
    rawValue: null,
    formula: '=UnknownFunction()',
    calculationState: 'pending',
    style: { fillColor: '#0000FF' },
  }
  await f.apply([
    { op: 'set_filter', sheetId: f.sheetId, range: 'A1:B3' },
    {
      op: 'set_filter_criteria',
      sheetId: f.sheetId,
      column: 'A',
      colors: { fill: ['#ff0000', '#0000ff'] },
    },
    { op: 'set_filter_criteria', sheetId: f.sheetId, column: 'B', values: ['12.50'] },
  ])
  expect(f.state.sheets[0]!.features!.filterEdits).toMatchObject({
    hiddenRows: [2],
    filter: {
      columns: [
        { colId: 0, colors: { fill: ['#FF0000', '#0000FF'] } },
        { colId: 1, values: ['12.50'] },
      ],
    },
  })
  expect(f.state.sheets[0]!.cells.A3!.formula).toBe('=UnknownFunction()')
  await f.apply([
    { op: 'format_range', sheetId: f.sheetId, range: 'A2', format: { fillColor: '#00FF00' } },
    { op: 'reapply_filter', sheetId: f.sheetId },
  ])
  expect(f.state.sheets[0]!.features!.filterEdits).toMatchObject({ hiddenRows: [1, 2] })
})
