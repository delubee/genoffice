import { useEffect, useRef, useState } from 'react'
import type { EditorHost } from './ports'
import { WorkbookRuntime } from './runtime'
import { editorText } from './i18n'
import { editorError } from './editor-error'
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { columnLabel } from '@genoffice/xlsx-gateway/domain/cell-address'
import { BasicChart } from './BasicChart'
import { WorkbookRibbon, RibbonGroup } from './WorkbookRibbon'
import type { SheetVisual } from '@genoffice/xlsx-gateway/domain/chart-visual'
import type { WorkbookAppearance } from './appearance'

export function WorkbookEditor({
  host,
  onReady,
  locale,
  appearance = 'vscode',
  onAppearanceChange,
}: {
  host: EditorHost
  onReady(runtime: WorkbookRuntime): void
  locale?: string
  appearance?: WorkbookAppearance
  onAppearanceChange?(appearance: WorkbookAppearance): void
}) {
  const container = useRef<HTMLDivElement>(null),
    runtime = useRef<WorkbookRuntime | null>(null)
  const [error, setError] = useState(''),
    [name, setName] = useState(''),
    [format, setFormat] = useState('0.00'),
    [sourceRange, setSourceRange] = useState('A1'),
    [property, setProperty] = useState<
      | 'fontFamily'
      | 'fontSize'
      | 'fontColor'
      | 'fillColor'
      | 'horizontalAlign'
      | 'verticalAlign'
      | 'wrapText'
      | 'border'
    >('fontFamily'),
    [styleValue, setStyleValue] = useState('Calibri')
  const [charts, setCharts] = useState<readonly SheetVisual[]>([]),
    [chartType, setChartType] = useState<'column' | 'bar' | 'line' | 'pie'>('column')
  const [, setSelectionVersion] = useState(0)
  const [rowHeight, setRowHeight] = useState(15),
    [columnWidth, setColumnWidth] = useState(96)
  const text = editorText(locale)
  const initialAppearance = useRef(appearance)
  useEffect(() => {
    const projection = new WorkbookRuntime(
      container.current!,
      host,
      (failure) => setError(editorError(failure, editorText(locale))),
      () => {
        setCharts(runtime.current?.charts ?? [])
        setSelectionVersion((version) => version + 1)
      },
      locale,
      initialAppearance.current,
    )
    runtime.current = projection
    onReady(projection)
    return () => {
      runtime.current = null
      projection.dispose()
    }
  }, [host, onReady, locale])
  useEffect(() => {
    runtime.current?.setAppearance(appearance)
  }, [appearance])
  const selected = () => {
    const selection = runtime.current?.selection
    if (!selection) throw new Error(text.selectFirst)
    return selection
  }
  const submit = (make: () => WorkbookOperation) => {
    try {
      setError('')
      const op = make()
      void runtime.current!.enqueue([op]).catch((failure) => setError(editorError(failure, text)))
    } catch (failure) {
      setError(editorError(failure, text))
    }
  }
  return (
    <section className="workbook-editor">
      {onAppearanceChange && appearance === 'excel' && (
        <div className="workbook-titlebar">
          <strong>GenOffice</strong>
          <span>{text.title}</span>
        </div>
      )}
      <div
        className="workbook-quick-access workbook-toolbar"
        role="toolbar"
        aria-label={text.title}
      >
        {onAppearanceChange && (
          <label className="workbook-appearance">
            {text.appearance}
            <select
              aria-label={text.appearance}
              value={appearance}
              onChange={(event) => onAppearanceChange(event.target.value as WorkbookAppearance)}
            >
              <option value="excel">{text.excelAppearance}</option>
              <option value="vscode">{text.vscodeAppearance}</option>
            </select>
          </label>
        )}
        <button
          onClick={() => {
            void runtime.current
              ?.requestHistory('undo')
              .catch((failure) => setError(editorError(failure, text)))
          }}
        >
          {text.undo}
        </button>
        <button
          onClick={() => {
            void runtime.current
              ?.requestHistory('redo')
              .catch((failure) => setError(editorError(failure, text)))
          }}
        >
          {text.redo}
        </button>
      </div>
      <WorkbookRibbon
        runtime={runtime.current}
        locale={locale}
        onError={(message) => setError(editorError(message, text))}
        home={
          <>
            <RibbonGroup title={text.fontAndAlignment}>
              <button
                onClick={() =>
                  submit(() => {
                    const s = selected()
                    return {
                      op: 'format_range',
                      sheetId: s.sheetId,
                      range: s.address,
                      format: { bold: s.style.bl !== 1 },
                    }
                  })
                }
              >
                {text.bold}
              </button>
              <button
                onClick={() =>
                  submit(() => {
                    const s = selected()
                    return {
                      op: 'format_range',
                      sheetId: s.sheetId,
                      range: s.address,
                      format: { italic: s.style.it !== 1 },
                    }
                  })
                }
              >
                {text.italic}
              </button>
              <input
                aria-label={text.format}
                value={format}
                onChange={(event) => setFormat(event.target.value)}
              />
              <button
                onClick={() =>
                  submit(() => {
                    const s = selected()
                    return {
                      op: 'format_range',
                      sheetId: s.sheetId,
                      range: s.address,
                      format: { numberFormat: format },
                    }
                  })
                }
              >
                {text.apply}
              </button>
              <details>
                <summary>{text.moreFormats}</summary>
                <select
                  aria-label={text.moreFormats}
                  value={property}
                  onChange={(event) => {
                    const next = event.target.value as typeof property
                    setProperty(next)
                    setStyleValue(
                      {
                        fontFamily: 'Calibri',
                        fontSize: '11',
                        fontColor: '#000000',
                        fillColor: '#FFFFFF',
                        horizontalAlign: 'left',
                        verticalAlign: 'top',
                        wrapText: 'true',
                        border: 'all',
                      }[next],
                    )
                  }}
                >
                  {(
                    [
                      'fontFamily',
                      'fontSize',
                      'fontColor',
                      'fillColor',
                      'horizontalAlign',
                      'verticalAlign',
                      'wrapText',
                      'border',
                    ] as const
                  ).map((key) => (
                    <option key={key} value={key}>
                      {text[key]}
                    </option>
                  ))}
                </select>
                {property === 'horizontalAlign' ||
                property === 'verticalAlign' ||
                property === 'wrapText' ||
                property === 'border' ? (
                  <select
                    aria-label={text[property]}
                    value={styleValue}
                    onChange={(event) => setStyleValue(event.target.value)}
                  >
                    {(property === 'horizontalAlign'
                      ? ['left', 'center', 'right']
                      : property === 'verticalAlign'
                        ? ['top', 'center', 'bottom']
                        : property === 'border'
                          ? ['all', 'none', 'top', 'bottom', 'left', 'right']
                          : ['true', 'false']
                    ).map((value) => (
                      <option key={value} value={value}>
                        {value === 'true'
                          ? text.enabled
                          : value === 'false'
                            ? text.disabled
                            : text[value as 'left']}
                      </option>
                    ))}
                  </select>
                ) : (
                  <input
                    aria-label={text[property]}
                    type={
                      property === 'fontSize'
                        ? 'number'
                        : property === 'fontColor' || property === 'fillColor'
                          ? 'color'
                          : 'text'
                    }
                    value={styleValue}
                    onChange={(event) => setStyleValue(event.target.value)}
                  />
                )}
                <button
                  onClick={() =>
                    submit(() => {
                      const s = selected()
                      const value =
                        property === 'fontSize'
                          ? Number(styleValue)
                          : property === 'wrapText'
                            ? styleValue === 'true'
                            : property === 'border'
                              ? { type: styleValue }
                              : styleValue
                      return {
                        op: 'format_range',
                        sheetId: s.sheetId,
                        range: s.address,
                        format: { [property]: value },
                      } as WorkbookOperation
                    })
                  }
                >
                  {text.apply}
                </button>
              </details>
            </RibbonGroup>
            <RibbonGroup title={text.clipboard}>
              <input
                aria-label={text.sourceRange}
                title={text.sourceRange}
                value={sourceRange}
                onChange={(event) => setSourceRange(event.target.value.toUpperCase())}
              />
              <button
                onClick={() =>
                  submit(() => ({
                    op: 'copy_range',
                    sheetId: selected().sheetId,
                    source: sourceRange,
                    target: selected().address,
                  }))
                }
              >
                {text.copyRange}
              </button>
              <button
                onClick={() =>
                  submit(() => ({
                    op: 'fill_range',
                    sheetId: selected().sheetId,
                    source: sourceRange,
                    target: selected().address,
                  }))
                }
              >
                {text.fillRange}
              </button>
            </RibbonGroup>
            <RibbonGroup title={text.cells}>
              <input
                aria-label={text.sheetName}
                value={name}
                onChange={(event) => setName(event.target.value)}
              />
              <button onClick={() => submit(() => ({ op: 'add_sheet', name }))}>
                {text.addSheet}
              </button>
              <button
                onClick={() =>
                  submit(() => ({ op: 'rename_sheet', sheetId: selected().sheetId, name }))
                }
              >
                {text.renameSheet}
              </button>
              <button
                onClick={() => submit(() => ({ op: 'delete_sheet', sheetId: selected().sheetId }))}
              >
                {text.deleteSheet}
              </button>
              <button
                onClick={() =>
                  submit(() => {
                    const s = selected()
                    return {
                      op: 'insert_rows',
                      sheetId: s.sheetId,
                      row: s.range.startRow + 1,
                      count: 1,
                    }
                  })
                }
              >
                {text.insertRow}
              </button>
              <button
                onClick={() =>
                  submit(() => {
                    const s = selected()
                    return {
                      op: 'delete_rows',
                      sheetId: s.sheetId,
                      row: s.range.startRow + 1,
                      count: 1,
                    }
                  })
                }
              >
                {text.deleteRow}
              </button>
              <button
                onClick={() =>
                  submit(() => {
                    const s = selected()
                    return {
                      op: 'insert_cols',
                      sheetId: s.sheetId,
                      column: columnLabel(s.range.startColumn),
                      count: 1,
                    }
                  })
                }
              >
                {text.insertColumn}
              </button>
              <button
                onClick={() =>
                  submit(() => {
                    const s = selected()
                    return {
                      op: 'delete_cols',
                      sheetId: s.sheetId,
                      column: columnLabel(s.range.startColumn),
                      count: 1,
                    }
                  })
                }
              >
                {text.deleteColumn}
              </button>
            </RibbonGroup>
            <RibbonGroup title={text.cellSize}>
              <label>
                {text.rowHeight}
                <input
                  aria-label={text.rowHeight}
                  type="number"
                  min="2"
                  max="409"
                  value={rowHeight}
                  onChange={(event) => setRowHeight(Number(event.target.value))}
                />
              </label>
              <button
                onClick={() =>
                  submit(() => {
                    const s = selected()
                    return {
                      op: 'set_row_height',
                      sheetId: s.sheetId,
                      row: s.range.startRow + 1,
                      count: s.range.endRow - s.range.startRow + 1,
                      heightPoints: rowHeight,
                    }
                  })
                }
              >
                {text.setRowHeight}
              </button>
              <label>
                {text.columnWidth}
                <input
                  aria-label={text.columnWidth}
                  type="number"
                  min="10"
                  max="2000"
                  value={columnWidth}
                  onChange={(event) => setColumnWidth(Number(event.target.value))}
                />
              </label>
              <button
                onClick={() =>
                  submit(() => {
                    const s = selected()
                    return {
                      op: 'set_col_width',
                      sheetId: s.sheetId,
                      column: columnLabel(s.range.startColumn),
                      count: s.range.endColumn - s.range.startColumn + 1,
                      widthPx: columnWidth,
                    }
                  })
                }
              >
                {text.setColumnWidth}
              </button>
            </RibbonGroup>
          </>
        }
        insert={
          <RibbonGroup title={text.chart}>
            <select
              aria-label={text.chart}
              value={chartType}
              onChange={(event) => setChartType(event.target.value as typeof chartType)}
            >
              {(['column', 'bar', 'line', 'pie'] as const).map((type) => (
                <option key={type} value={type}>
                  {text[type]}
                </option>
              ))}
            </select>
            <button
              onClick={() =>
                submit(() => ({
                  op: 'add_chart',
                  sheetId: selected().sheetId,
                  dataRange: selected().address,
                  chartType,
                }))
              }
            >
              {text.addChart}
            </button>
          </RibbonGroup>
        }
      />
      {error && (
        <div className="workbook-error" role="alert">
          {error}
        </div>
      )}
      <div className="workbook-content">
        <div ref={container} className="workbook-grid" aria-label={text.title} />
        {charts.length > 0 && (
          <aside className="workbook-charts" aria-label={text.chart}>
            {charts.map((visual) => (
              <section key={visual.id}>
                <BasicChart chart={visual.chart} />
                <form
                  key={`${visual.id}:${visual.chart.title}`}
                  onSubmit={(event) => {
                    event.preventDefault()
                    const fields = new FormData(event.currentTarget)
                    submit(() => ({
                      op: 'edit_chart',
                      chartPath: visual.id,
                      title: String(fields.get('title')),
                      chartType: String(fields.get('type')) as typeof chartType,
                    }))
                  }}
                >
                  <input
                    name="title"
                    aria-label={text.chartTitle}
                    defaultValue={visual.chart.title}
                    maxLength={255}
                  />
                  <select
                    name="type"
                    aria-label={text.chart}
                    defaultValue={
                      visual.chart.chartTypes.includes('pieChart')
                        ? 'pie'
                        : visual.chart.chartTypes.includes('lineChart')
                          ? 'line'
                          : visual.chart.barDirection === 'bar'
                            ? 'bar'
                            : 'column'
                    }
                  >
                    {(['column', 'bar', 'line', 'pie'] as const).map((type) => (
                      <option key={type} value={type}>
                        {text[type]}
                      </option>
                    ))}
                  </select>
                  <button type="submit">{text.apply}</button>
                  <button
                    type="button"
                    onClick={() => submit(() => ({ op: 'delete_visual', visualId: visual.id }))}
                  >
                    {text.deleteChart}
                  </button>
                </form>
              </section>
            ))}
          </aside>
        )}
      </div>
    </section>
  )
}
