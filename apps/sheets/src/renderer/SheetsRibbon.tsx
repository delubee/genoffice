/** Desktop and embedded spreadsheet command band. Host actions are injected callbacks. */
import { platformShortcuts } from '@genoffice/i18n'
import { SHAPE_GALLERY_GROUPS, ShapePreview, useDismissablePopover } from '@genoffice/ui/chrome'
import type { ChartSeriesVisualState } from '@genoffice/xlsx-gateway/domain/chart-visual'
import { useEffect, useRef, useState } from 'react'
import { ColorDropdown } from './ColorDropdown'
import type { HeaderFooterParts } from './ribbon-model'
import {
  accountingStyle,
  COMMA_STYLE,
  PERCENT_STYLE,
  RIBBON_CURRENCY_SYMBOL,
} from './excel-format-shortcuts'
import { stepFontSize } from './font-size-ladder'
import { useI18n, type StringKey } from './i18n/locale'
import {
  categoryOptionForPattern,
  NUMBER_FORMAT_LABEL,
  numberFormatCategories,
} from './number-format'
import { PASTE_SPECIAL_OPEN_SHELL_COMMAND } from './paste-special-model'
import {
  BorderAllIcon,
  BorderBottomIcon,
  BorderLeftIcon,
  BorderNoneIcon,
  BorderOuterIcon,
  BorderRightIcon,
  BorderThickOuterIcon,
  BorderTopIcon,
  CaretIcon,
  GensparkMark,
  RIBBON_GLYPH_ICONS,
} from './ribbon-icons'
import { type SelectionFormat } from './selection-format'
import { fontFamilyGroups, useSystemFontFamilies } from './system-fonts'
import type { TableOptionKey } from './ribbon-model'
import type { SelectedTableRibbon } from './ribbon-model'
import { builtinTablePalette, tableStyleNames, type TableStyleFamily } from './table-styles'
import { COLOR_SCHEMES, FONT_SCHEMES, THEME_PRESETS } from './themes'

// No File tab: file commands live in the macOS
// application menu (File → Open/Save/Save As) and the toolbar icons.
export const ribbonTabs = [
  'Home',
  'Insert',
  'Page Layout',
  'Formulas',
  'Data',
  'Review',
  'View',
] as const

/// 'Chart Design' is contextual: it exists only while a chart is selected,
/// and appears without stealing the active tab.
export type RibbonTab =
  (typeof ribbonTabs)[number] | 'Chart Design' | 'Table Design' | 'Shape Format' | 'Picture Format'

/// Paint of the selected editable shape; `'none'` is an explicit no fill/line.
export interface SelectedShapeRibbon {
  readonly fillColor: string | undefined
  readonly lineColor: string | undefined
}

export interface SelectedChartRibbon {
  readonly title: string
  /// Non-null when the chart's type can be rewritten (single-plot
  /// column/bar/line/area/pie/doughnut); scatter/combo/3D keep their type.
  readonly convertible: string | null
  readonly currentType: string | null
  readonly canEdit: boolean
  readonly isPie: boolean
  readonly hasAxes: boolean
  /// Whether the save pipeline can write plot-level data labels for this
  /// chart family (scatter/radar cannot).
  readonly canLabel: boolean
  readonly seriesCount: number
  readonly categoryCount: number
  /// Current series (pending edits applied) for Switch Row/Column.
  readonly series: readonly ChartSeriesVisualState[]
  readonly legend?: string | undefined
  readonly axisTitles?:
    | {
        category?: string | null | undefined
        value?: string | null | undefined
      }
    | undefined
  readonly dataLabels?: string | undefined
  readonly grouping?: string | undefined
}

/// Review > Translate targets, shown in their native names (never localized).
const TRANSLATE_LANGUAGES = [
  'English',
  '简体中文',
  '繁體中文',
  '日本語',
  '한국어',
  'Español',
  'Français',
  'Deutsch',
  'Português',
  'Русский',
  'العربية',
] as const

/// Glyph icons live in ribbon-icons.tsx, drawn to the shared icon standard
/// (24×24 canvas, 1.5px strokes, round caps/joins); unmapped glyphs render
/// as text (letterforms such as $, ?, θ, ƒx are typography, not icons).
function ToolSymbol({ symbol }: { readonly symbol: string }): React.JSX.Element {
  return (
    <span className="tool-symbol" aria-hidden="true">
      {RIBBON_GLYPH_ICONS[symbol] ?? symbol}
    </span>
  )
}

export interface PageLayoutEcho {
  readonly orientation?: 'portrait' | 'landscape' | undefined
  readonly paperSize?: number | undefined
  readonly scale?: number | undefined
  readonly fitToWidth?: number | undefined
  readonly fitToHeight?: number | undefined
  readonly margins?: 'normal' | 'wide' | 'narrow' | undefined
  readonly printGridlines?: boolean | undefined
  readonly printHeadings?: boolean | undefined
  readonly showGridlines: boolean
  readonly showHeadings: boolean
  /// Mirrored grid (column A at the right edge) for the active sheet.
  readonly rightToLeft: boolean
  readonly printArea?: string | null | undefined
  readonly printTitles?: string | null | undefined
  readonly header?: HeaderFooterParts | null | undefined
  readonly footer?: HeaderFooterParts | null | undefined
  /// Page Break Preview overlay on for the active sheet (View tab echo).
  readonly pageBreakPreview?: boolean | undefined
}

/// Table Name box: commits on Enter / blur, Esc restores.
function TableNameInput({
  name,
  onCommit,
}: {
  readonly name: string
  readonly onCommit: (name: string) => void
}): React.JSX.Element {
  const [value, setValue] = useState(name)
  // Escape blurs the box; the blur must then restore, not commit the stale text.
  const cancelledRef = useRef(false)
  const commit = (): void => {
    if (cancelledRef.current) {
      cancelledRef.current = false
      setValue(name)
      return
    }
    const next = value.trim()
    if (next.length > 0 && next !== name) onCommit(next)
    else setValue(name)
  }
  return (
    <input
      className="table-name-input"
      value={value}
      spellCheck={false}
      onChange={(event) => setValue(event.target.value)}
      onBlur={commit}
      onKeyDown={(event) => {
        if (event.key === 'Enter') {
          event.preventDefault()
          commit()
          event.currentTarget.blur()
        } else if (event.key === 'Escape') {
          cancelledRef.current = true
          setValue(name)
          event.currentTarget.blur()
        }
        event.stopPropagation()
      }}
    />
  )
}

/// Excel's Table Styles gallery: Light / Medium / Dark rows of swatches
/// drawn from the same palette rules the grid renders with.
function TableStyleGallery({
  current,
  onPick,
}: {
  readonly current: string | null
  readonly onPick: (style: string) => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [open, setOpen] = useState(false)
  const wrapRef = useRef<HTMLDivElement>(null)
  useDismissablePopover(open, () => setOpen(false), { inside: () => [wrapRef.current] })
  useEscapeClose(open, () => setOpen(false))
  const families: { family: TableStyleFamily; label: string }[] = [
    { family: 'Light', label: t('appTableStyleFamilyLight') },
    { family: 'Medium', label: t('appTableStyleFamilyMedium') },
    { family: 'Dark', label: t('appTableStyleFamilyDark') },
  ]
  const swatch = (style: string, label: string): React.JSX.Element => {
    const palette = builtinTablePalette(style, undefined)
    const rowFill = (index: number): string | undefined =>
      index % 2 === 0 ? palette.stripeFill : (palette.secondRowStripeFill ?? palette.wholeTableFill)
    return (
      <button
        key={style}
        className={`table-style-swatch ${current === style ? 'active' : ''}`}
        data-tip={label}
        onClick={() => {
          onPick(style)
          setOpen(false)
        }}
      >
        <span
          className="swatch-row header"
          style={{
            background: palette.headerFill,
            borderColor: palette.borderColor ?? palette.wholeTableBorderColor ?? palette.headerFill,
          }}
        >
          {[0, 1, 2].map((cell) => (
            <i key={cell} style={{ background: palette.headerFontColor }} />
          ))}
        </span>
        {[0, 1, 2, 3].map((row) => (
          <span
            key={row}
            className="swatch-row"
            style={{ background: rowFill(row), borderColor: palette.innerHorizontalBorderColor }}
          >
            {[0, 1, 2].map((cell) => (
              <i key={cell} style={{ background: palette.bodyFontColor }} />
            ))}
          </span>
        ))}
      </button>
    )
  }
  const currentLabel = (() => {
    const match = current ? /^TableStyle(Light|Medium|Dark)(\d+)$/.exec(current) : null
    if (!match) return t('appGroupTableStyles')
    const family = families.find((entry) => entry.family === match[1])
    return `${family?.label ?? match[1]} ${match[2]}`
  })()
  return (
    <div className="table-style-gallery" ref={wrapRef}>
      <button
        className="styles-row as-button"
        data-tip={t('appTableStylesGalleryTip')}
        aria-haspopup="dialog"
        aria-expanded={open}
        onClick={() => setOpen((value) => !value)}
      >
        <ToolSymbol symbol="▦" />
        {currentLabel}
        <CaretIcon />
      </button>
      {open && (
        <div className="table-style-panel" role="dialog" aria-label={t('appGroupTableStyles')}>
          {families.map(({ family, label }) => (
            <section key={family}>
              <header>{label}</header>
              <div className="table-style-grid">
                {tableStyleNames(family).map((style, index) =>
                  swatch(style, `${label} ${index + 1}`),
                )}
              </div>
            </section>
          ))}
        </div>
      )}
    </div>
  )
}

export function SheetsRibbon({
  activeTab,
  selectionFormat,
  formatPainterActive,
  sheetHasContent,
  sheetProtected,
  workbookProtected,
  formulaBarVisible,
  crossHighlightVisible,
  pageLayout,
  selectedChart,
  selectedShape,
  selectedTable,
  outlineSummary,
  onCommand,
  onAiRun,
  aiOpen,
  onAiToggle,
  onListNames,
  calcManual,
  onRefreshPivot,
  onIsSelectionInPivot,
}: {
  readonly activeTab: RibbonTab
  readonly selectionFormat: SelectionFormat | null
  readonly formatPainterActive: boolean
  readonly sheetHasContent: boolean
  readonly sheetProtected: boolean | null
  readonly workbookProtected: boolean | null
  /// View > Show echo for the formula bar toggle (app-level, not per sheet).
  readonly formulaBarVisible: boolean
  readonly crossHighlightVisible: boolean
  readonly pageLayout: PageLayoutEcho
  readonly selectedChart: SelectedChartRibbon | null
  readonly selectedShape: SelectedShapeRibbon | null
  readonly selectedTable: SelectedTableRibbon | null
  readonly outlineSummary: { below: boolean; right: boolean }
  readonly onCommand: (command: string) => void
  /** Defined names for the Use in Formula menu. */
  readonly onListNames: () => readonly string[]
  /** Manual-recalc mode echo for the Calculation Options menu. */
  readonly calcManual: boolean
  /** Open the AI panel and immediately send the given prompt */
  readonly onAiRun: (prompt: string) => void
  /** AI side panel visibility (docs/slides parity: the entry button toggles it) */
  readonly aiOpen: boolean
  readonly onAiToggle: () => void
  readonly onRefreshPivot: () => string | null
  readonly onIsSelectionInPivot: () => boolean
}): React.JSX.Element {
  const { t } = useI18n()
  const [fontColor, setFontColor] = useState('#C00000')
  const [fillColor, setFillColor] = useState('#FFF2CC')
  const [borderColor, setBorderColor] = useState('#000000')
  const { families: systemFontFamilies, load: loadSystemFonts } = useSystemFontFamilies()
  // Large menu button: a native select stretched invisibly over the tool,
  // each option carrying its full command string.
  const largeMenu = (
    label: string,
    symbol: string,
    title: string,
    options: readonly { value: string; label: string }[],
  ): React.JSX.Element => (
    <div className="ribbon-tool large" data-tip={title}>
      <span className="tool-icon-row">
        <ToolSymbol symbol={symbol} />
        <CaretIcon />
      </span>
      <span>
        <strong>{label}</strong>
      </span>
      <MenuSelect cover label={label} options={options} onPick={onCommand} />
    </div>
  )
  if (activeTab === 'Chart Design') {
    const canConvert = Boolean(selectedChart?.convertible && selectedChart.canEdit)
    const canEditChart = Boolean(selectedChart?.canEdit)
    const labelValue = selectedChart?.isPie ? 'percent' : 'value'
    // ✓ mirrors the chart's current state (parsed from the file + pending
    // edits), so the menu reads as checked toggles.
    const activeLabels = selectedChart?.dataLabels ?? 'none'
    const activeLegend = selectedChart?.legend ?? 'none'
    const check = (on: boolean): string => (on ? ' ✓' : '')
    const elementOptions = [
      { value: 'chart-element-title', label: t('appChartElTitle') },
      ...(selectedChart?.hasAxes
        ? [
            {
              value: 'chart-element-axis-cat',
              label: t('appChartElAxisCat') + check(Boolean(selectedChart.axisTitles?.category)),
            },
            {
              value: 'chart-element-axis-val',
              label: t('appChartElAxisVal') + check(Boolean(selectedChart.axisTitles?.value)),
            },
          ]
        : []),
      ...(selectedChart?.canLabel
        ? [
            ...(selectedChart.isPie
              ? [
                  {
                    value: 'chart-labels:category-percent',
                    label:
                      t('appChartElLabelsNamePct') + check(activeLabels === 'category-percent'),
                  },
                ]
              : []),
            {
              value: `chart-labels:${labelValue}`,
              label:
                t(selectedChart.isPie ? 'appChartElLabelsPct' : 'appChartElLabelsValues') +
                check(activeLabels === labelValue),
            },
            {
              value: 'chart-labels:none',
              label: t('appChartElLabelsNone') + check(activeLabels === 'none'),
            },
          ]
        : []),
      {
        value: 'chart-legend:right',
        label: t('appChartElLegendRight') + check(activeLegend === 'right'),
      },
      {
        value: 'chart-legend:top',
        label: t('appChartElLegendTop') + check(activeLegend === 'top'),
      },
      {
        value: 'chart-legend:bottom',
        label: t('appChartElLegendBottom') + check(activeLegend === 'bottom'),
      },
      {
        value: 'chart-legend:left',
        label: t('appChartElLegendLeft') + check(activeLegend === 'left'),
      },
      {
        value: 'chart-legend:none',
        label: t('appChartElLegendNone') + check(activeLegend === 'none'),
      },
    ]
    const layoutLabels = t(selectedChart?.isPie ? 'appLayoutLabelsNamePct' : 'appLayoutLabelsValue')
    const layoutOptions = [
      { value: 'chart-layout:1', label: t('appLayout1', { labels: layoutLabels }) },
      { value: 'chart-layout:2', label: t('appLayout2', { labels: layoutLabels }) },
      { value: 'chart-layout:3', label: t('appLayout3') },
      { value: 'chart-layout:4', label: t('appLayout4', { labels: layoutLabels }) },
    ]
    const colorOptions = [
      { value: 'chart-colors:office', label: t('appColorsOffice') },
      { value: 'chart-colors:blue', label: t('appColorsBlue') },
      { value: 'chart-colors:green', label: t('appColorsGreen') },
      { value: 'chart-colors:warm', label: t('appColorsWarm') },
      { value: 'chart-colors:gray', label: t('appColorsGray') },
    ]
    const canRecolor =
      canEditChart &&
      ((selectedChart?.isPie ? selectedChart.categoryCount : selectedChart?.seriesCount) ?? 0) > 0
    const stackable =
      canEditChart && ['column', 'bar', 'line', 'area'].includes(selectedChart?.currentType ?? '')
    // 'standard' is how line/area spell side-by-side in OOXML.
    const activeGrouping =
      selectedChart?.grouping === 'stacked' || selectedChart?.grouping === 'percentStacked'
        ? selectedChart.grouping
        : 'clustered'
    const groupingOptions = [
      {
        value: 'chart-grouping:clustered',
        label: t('appGroupingClustered') + check(activeGrouping === 'clustered'),
      },
      {
        value: 'chart-grouping:stacked',
        label: t('appGroupingStacked') + check(activeGrouping === 'stacked'),
      },
      {
        value: 'chart-grouping:percentStacked',
        label: t('appGroupingPercent') + check(activeGrouping === 'percentStacked'),
      },
    ]
    return (
      <div className="ribbon" data-ribbon-body="">
        <RibbonGroup label={t('appGroupChartLayouts')}>
          {canEditChart ? (
            largeMenu(t('appAddChartElement'), '📊', t('appAddChartElementTitle'), elementOptions)
          ) : (
            <RibbonButton
              large
              menu
              label={t('appAddChartElement')}
              detail={t('appSelectEditableChart')}
              symbol="📊"
              disabled
              onClick={() => {}}
            />
          )}
          {canEditChart ? (
            largeMenu(t('appQuickLayout'), '▦', t('appQuickLayoutTitle'), layoutOptions)
          ) : (
            <RibbonButton
              large
              menu
              label={t('appQuickLayout')}
              detail={t('appSelectEditableChart')}
              symbol="▦"
              disabled
              onClick={() => {}}
            />
          )}
          {canRecolor ? (
            largeMenu(
              t('appChangeColors'),
              '🎨',
              t(selectedChart?.isPie ? 'appRecolorSlices' : 'appRecolorSeries'),
              colorOptions,
            )
          ) : (
            <RibbonButton
              large
              label={t('appChangeColors')}
              detail={t('appSelectEditableChart')}
              symbol="🎨"
              disabled
              onClick={() => {}}
            />
          )}
        </RibbonGroup>
        <RibbonGroup label={t('appChangeChartType')}>
          {(
            [
              ['column', 'appChartColumn'],
              ['bar', 'appChartBar'],
              ['line', 'appChartLine'],
              ['area', 'appChartTypeArea'],
              ['pie', 'appChartPie'],
              ['doughnut', 'appChartDoughnut'],
            ] as const
          ).map(([type, labelKey]) => (
            <RibbonButton
              key={type}
              large
              label={t(labelKey)}
              detail={
                selectedChart?.currentType === type
                  ? t('appCurrentType')
                  : canConvert
                    ? t('appChangeChartType')
                    : t('appScatterKeepType')
              }
              symbol={type === 'line' || type === 'area' ? '🗠' : '📊'}
              active={selectedChart?.currentType === type}
              disabled={!canConvert}
              onClick={() => onCommand(`chart-type-${type}`)}
            />
          ))}
          {stackable
            ? largeMenu(t('appGroupingBtn'), '▤', t('appGroupingTitle'), groupingOptions)
            : null}
        </RibbonGroup>
        <RibbonGroup label={t('appGroupData')}>
          <RibbonButton
            large
            label={t('appSelectDataBtn')}
            detail={t('appSeriesNamesRanges')}
            symbol="📊"
            disabled={!canEditChart}
            onClick={() => onCommand('chart-select-data')}
          />
          <RibbonButton
            large
            label={t('appSwitchRowColumn')}
            detail={t('appSwitchRowColumnDetail')}
            symbol="⇄"
            disabled={!canEditChart || (selectedChart?.categoryCount ?? 0) === 0}
            onClick={() => onCommand('chart-switch-row-col')}
          />
          <RibbonButton
            large
            label={t('appFormatPane')}
            detail={t('appFormatPaneDetail')}
            symbol="🎨"
            disabled={!canEditChart}
            onClick={() => onCommand('chart-format-pane')}
          />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupChartActions')}>
          <RibbonButton
            large
            label={t('appDeleteChart')}
            detail={selectedChart?.title ?? t('appRemoveFromSheet')}
            symbol="🗑"
            onClick={() => onCommand('chart-delete')}
          />
        </RibbonGroup>
      </div>
    )
  }

  if (activeTab === 'Shape Format' || activeTab === 'Picture Format') {
    const paintSwatch = (color: string | undefined): string =>
      color === undefined || color === 'none' ? 'transparent' : color
    // No current paint: the glyph opens the palette instead of re-applying
    // a placeholder color.
    const hasPaint = (color: string | undefined): color is string =>
      color !== undefined && color !== 'none'
    const zOrder = (prefix: 'front' | 'back') => [
      {
        value: `visual:z:${prefix}`,
        label: t(prefix === 'front' ? 'appVisualBringToFront' : 'appVisualSendToBack'),
      },
      {
        value: `visual:z:${prefix === 'front' ? 'forward' : 'backward'}`,
        label: t(prefix === 'front' ? 'appVisualBringForward' : 'appVisualSendBackward'),
      },
    ]
    const alignOptions = [
      { value: 'visual:align:left', label: t('appVisualAlignLeft') },
      { value: 'visual:align:center', label: t('appVisualAlignCenter') },
      { value: 'visual:align:right', label: t('appVisualAlignRight') },
      { value: 'visual:align:top', label: t('appVisualAlignTop') },
      { value: 'visual:align:middle', label: t('appVisualAlignMiddle') },
      { value: 'visual:align:bottom', label: t('appVisualAlignBottom') },
      { value: 'visual:distribute:horizontal', label: t('appVisualDistributeH') },
      { value: 'visual:distribute:vertical', label: t('appVisualDistributeV') },
    ]
    const rotateOptions = [
      { value: 'visual:rotate:cw', label: t('appVisualRotateCw') },
      { value: 'visual:rotate:ccw', label: t('appVisualRotateCcw') },
      { value: 'visual:flip:v', label: t('appVisualFlipV') },
      { value: 'visual:flip:h', label: t('appVisualFlipH') },
    ]
    return (
      <div className="ribbon" data-ribbon-body="">
        {selectedShape && (
          <RibbonGroup label={t('appGroupShapeStyles')}>
            <ColorDropdown
              label="Shape fill"
              data-tip={t('appShapeFill')}
              split={hasPaint(selectedShape.fillColor)}
              display={
                <span className="swatch-letter">
                  <ToolSymbol symbol="◧" />
                  <i style={{ background: paintSwatch(selectedShape.fillColor) }} />
                </span>
              }
              value={hasPaint(selectedShape.fillColor) ? selectedShape.fillColor : '#FFFFFF'}
              auto={t('dlgFcNoFill')}
              onPick={(hex) => onCommand(`shape-fill:${hex ?? 'none'}`)}
            />
            <ColorDropdown
              label="Shape outline"
              data-tip={t('appShapeOutline')}
              split={hasPaint(selectedShape.lineColor)}
              display={
                <span className="swatch-letter">
                  <ToolSymbol symbol="◻" />
                  <i style={{ background: paintSwatch(selectedShape.lineColor) }} />
                </span>
              }
              value={hasPaint(selectedShape.lineColor) ? selectedShape.lineColor : '#FFFFFF'}
              auto={t('appShapeNoOutline')}
              onPick={(hex) => onCommand(`shape-outline:${hex ?? 'none'}`)}
            />
            <RibbonButton
              large
              label={t('appDeleteShape')}
              detail={t('appRemoveFromSheet')}
              symbol="🗑"
              onClick={() => onCommand('shape-delete')}
            />
          </RibbonGroup>
        )}
        <RibbonGroup label={t('appGroupArrange')}>
          {largeMenu(t('appVisualBringToFront'), '⬆', t('appVisualBringToFront'), zOrder('front'))}
          {largeMenu(t('appVisualSendToBack'), '⬇', t('appVisualSendToBack'), zOrder('back'))}
          {largeMenu(t('appVisualAlign'), '⫴', t('appVisualAlignTitle'), alignOptions)}
          {largeMenu(t('appVisualRotate'), '⟳', t('appVisualRotateTitle'), rotateOptions)}
        </RibbonGroup>
        <RibbonGroup label={t('appGroupSize')}>
          <RibbonButton
            large
            label={t('appVisualSizeProperties')}
            detail=""
            symbol="⤢"
            onClick={() => onCommand('visual:size')}
          />
          <RibbonButton
            large
            label={t('appVisualEditAltText')}
            detail=""
            symbol="Alt"
            onClick={() => onCommand('visual:alt-text')}
          />
        </RibbonGroup>
      </div>
    )
  }
  if (activeTab === 'Table Design') {
    const table = selectedTable
    const option = (key: TableOptionKey, label: string): React.JSX.Element => (
      <button
        className="check-item"
        data-tip={label}
        disabled={!table}
        onClick={() => onCommand(`table-option:${key}:${table?.[key] ? 0 : 1}`)}
      >
        <i className="check-box">{table?.[key] ? '\u2713' : ''}</i>
        {label}
      </button>
    )
    return (
      <div className="ribbon" data-ribbon-body="">
        <RibbonGroup label={t('appGroupTableProperties')}>
          <div className="row-stack table-name-field">
            <label className="table-name-label">{t('appTableNameLabel')}</label>
            <TableNameInput
              key={table?.key ?? ''}
              name={table?.name ?? ''}
              onCommit={(name) => onCommand(`table-rename:${name}`)}
            />
          </div>
          <RibbonButton
            large
            label={t('appResizeTable')}
            detail={table?.rangeA1 ?? ''}
            symbol="⤢"
            disabled={!table}
            onClick={() => onCommand('table-resize-open')}
          />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupTableTools')}>
          <RibbonButton
            large
            label={t('appRemoveDuplicates')}
            detail={t('appInSelection')}
            symbol="⧉"
            onClick={() => onCommand('remove-duplicates-open')}
          />
          <RibbonButton
            large
            label={t('appConvertToRange')}
            detail={table?.name ?? ''}
            symbol="▦"
            disabled={!table}
            onClick={() => onCommand('table-convert-range')}
          />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupTableStyleOptions')}>
          <div className="row-stack table-options">
            {option('headerRow', t('appTableOptHeaderRow'))}
            {option('totalsRow', t('appTableOptTotalRow'))}
            {option('bandedRows', t('appTableOptBandedRows'))}
          </div>
          <div className="row-stack table-options">
            {option('firstColumn', t('appTableOptFirstColumn'))}
            {option('lastColumn', t('appTableOptLastColumn'))}
            {option('bandedColumns', t('appTableOptBandedColumns'))}
          </div>
          <div className="row-stack table-options">
            {option('filterButton', t('appTableOptFilterButton'))}
          </div>
        </RibbonGroup>
        <RibbonGroup label={t('appGroupTableStyles')}>
          <TableStyleGallery
            current={table?.style ?? null}
            onPick={(style) => onCommand(`table-style:${style}`)}
          />
        </RibbonGroup>
      </div>
    )
  }

  if (activeTab === 'Insert') {
    return (
      <div className="ribbon" data-ribbon-body="">
        <RibbonGroup label={t('appGroupTables')}>
          <RibbonButton
            large
            label={t('appPivotTable')}
            detail={t('appFromSelection')}
            symbol="⊞"
            onClick={() => onCommand('pivot-open')}
          />
          <RibbonButton
            large
            label={t('appEditPivotTable')}
            detail={t('appChangeFields')}
            symbol="⊞"
            onClick={() => onCommand('pivot-edit')}
          />
          <RibbonButton
            large
            label={t('appTableBtn')}
            detail={t('appRealExcelTable')}
            symbol="▦"
            onClick={() => onCommand('format-as-table')}
          />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupIllustrations')}>
          <RibbonButton
            large
            label={t('appPictures')}
            detail={t('appPictureTypes')}
            symbol="🖼"
            onClick={() => onCommand('insert-picture')}
          />
          <div className="row-stack">
            <span className="styles-row" data-tip={t('appInsertShapeTitle')}>
              <ToolSymbol symbol="◇" />
              {t('appShapes')}
              <CaretIcon />
              <ShapeGallerySelect
                label="Shapes"
                onPick={(prst) => onCommand(`insert-shape:${prst}`)}
              />
            </span>
            <button
              className="styles-row as-button"
              data-tip={t('appIcons')}
              onClick={() => onCommand('insert-icons')}
            >
              <ToolSymbol symbol="✧" />
              {t('appIcons')}
            </button>
            <button
              className="styles-row as-button"
              data-tip={t('appScreenshot')}
              onClick={() => onCommand('insert-screenshot')}
            >
              <ToolSymbol symbol="⧉" />
              {t('appScreenshot')}
            </button>
          </div>
        </RibbonGroup>
        <RibbonGroup label={t('appGroupCheckbox')}>
          <RibbonButton
            large
            label={t('appGroupCheckbox')}
            detail={t('appAtSelection')}
            symbol="☑"
            onClick={() => onCommand('insert-checkbox')}
          />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupCharts')}>
          <RibbonButton
            large
            label={t('appRecommendedCharts')}
            detail={t('appFromSelection')}
            symbol="📊"
            onClick={() => onCommand('recommended-charts-open')}
          />
          <div className="chart-grid">
            <button
              data-tip={t('appChartGridTitle', { type: t('appChartColumn') })}
              aria-label={t('appChartGridTitle', { type: t('appChartColumn') })}
              onClick={() => onCommand('insert-chart:column')}
            >
              <ToolSymbol symbol="▮▬" />
            </button>
            <button
              data-tip={t('appChartGridTitle', { type: t('appChartBar') })}
              aria-label={t('appChartGridTitle', { type: t('appChartBar') })}
              onClick={() => onCommand('insert-chart:bar')}
            >
              <ToolSymbol symbol="▤" />
            </button>
            <button
              data-tip={t('appChartGridTitle', { type: t('appChartLine') })}
              aria-label={t('appChartGridTitle', { type: t('appChartLine') })}
              onClick={() => onCommand('insert-chart:line')}
            >
              <ToolSymbol symbol="📈" />
            </button>
            <button
              data-tip={t('appChartGridTitle', { type: t('appChartTypeArea') })}
              aria-label={t('appChartGridTitle', { type: t('appChartTypeArea') })}
              onClick={() => onCommand('insert-chart:area')}
            >
              <ToolSymbol symbol="◪" />
            </button>
            <button
              data-tip={t('appChartGridTitle', { type: t('appChartPie') })}
              aria-label={t('appChartGridTitle', { type: t('appChartPie') })}
              onClick={() => onCommand('insert-chart:pie')}
            >
              <ToolSymbol symbol="◔" />
            </button>
            <button
              data-tip={t('appChartGridTitle', { type: t('appChartScatter') })}
              aria-label={t('appChartGridTitle', { type: t('appChartScatter') })}
              onClick={() => onCommand('insert-chart:scatter')}
            >
              <ToolSymbol symbol="∴" />
            </button>
            <button
              data-tip={t('appChartGridTitle', { type: t('appChartRadar') })}
              aria-label={t('appChartGridTitle', { type: t('appChartRadar') })}
              onClick={() => onCommand('insert-chart:radar')}
            >
              <ToolSymbol symbol="✳" />
            </button>
            <button
              data-tip={t('appChartGridTitle', { type: t('appChartDoughnut') })}
              aria-label={t('appChartGridTitle', { type: t('appChartDoughnut') })}
              onClick={() => onCommand('insert-chart:doughnut')}
            >
              <ToolSymbol symbol="◍" />
            </button>
            <button
              data-tip={t('appChartGridTitle', { type: t('appChartCombo') })}
              aria-label={t('appChartGridTitle', { type: t('appChartCombo') })}
              onClick={() => onCommand('insert-chart:combo')}
            >
              <ToolSymbol symbol="𝄜" />
            </button>
          </div>
          {largeMenu(
            t('appPivotChart'),
            '🗠',
            onIsSelectionInPivot() ? t('appPivotChartHintIn') : t('appPivotChartHintOut'),
            [
              { value: 'insert-pivot-chart:column', label: t('appChartColumn') },
              { value: 'insert-pivot-chart:bar', label: t('appChartBar') },
              { value: 'insert-pivot-chart:line', label: t('appChartLine') },
              { value: 'insert-pivot-chart:pie', label: t('appChartPie') },
              { value: 'insert-pivot-chart:doughnut', label: t('appChartDoughnut') },
              { value: 'insert-pivot-chart:radar', label: t('appChartRadar') },
            ],
          )}
        </RibbonGroup>
        <RibbonGroup label={t('appGroupSparklines')}>
          {largeMenu(t('appGroupSparklines'), '〜', t('appSparklinesTitle'), [
            { value: 'sparkline:line', label: t('appSparkLine') },
            { value: 'sparkline:column', label: t('appSparkColumn') },
            { value: 'sparkline:stacked', label: t('appSparkWinLoss') },
          ])}
        </RibbonGroup>
        <RibbonGroup label={t('appGroupFilters')}>
          <div className="row-stack">
            <button
              className="styles-row as-button"
              data-tip={onIsSelectionInPivot() ? t('appSlicerHintIn') : t('appSlicerHintOut')}
              onClick={() => onCommand('slicer-open')}
            >
              <ToolSymbol symbol="▥" />
              {t('appSlicer')}
            </button>
            <button
              className="styles-row as-button"
              data-tip={onIsSelectionInPivot() ? t('appTimelineHintIn') : t('appTimelineHintOut')}
              onClick={() => onCommand('timeline-open')}
            >
              <ToolSymbol symbol="🕒" />
              {t('appTimeline')}
            </button>
          </div>
        </RibbonGroup>
        <RibbonGroup label={t('appGroupLinks')}>
          <RibbonButton
            large
            label={t('appLink')}
            detail={t('appAtSelection')}
            symbol="🔗"
            onClick={() => onCommand('link-open')}
          />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupComments')}>
          <RibbonButton
            large
            label={t('appNewComment')}
            detail={t('appAtSelection')}
            symbol="🗨"
            onClick={() => onCommand('note-open')}
          />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupText')}>
          <RibbonButton
            large
            label={t('appTextBox')}
            detail={t('appAtSelection')}
            symbol="A"
            onClick={() => onCommand('insert-textbox')}
          />
          <RibbonButton
            large
            label={t('appHeaderFooter')}
            detail={t('appPrintedPages')}
            symbol="🗎"
            onClick={() => onCommand('header-footer-open')}
          />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupSymbols')}>
          <div className="row-stack">
            <button
              className="styles-row as-button"
              data-tip={t('appEquation')}
              onClick={() => onCommand('insert-equation')}
            >
              <ToolSymbol symbol="π" />
              {t('appEquation')}
            </button>
            <button
              className="styles-row as-button"
              data-tip={t('appSymbol')}
              onClick={() => onCommand('insert-symbol')}
            >
              <ToolSymbol symbol="Ω" />
              {t('appSymbol')}
            </button>
          </div>
        </RibbonGroup>
      </div>
    )
  }

  if (activeTab === 'Page Layout') {
    const fitOptions = (axis: 'width' | 'height') => {
      const current = String(
        (axis === 'width' ? pageLayout.fitToWidth : pageLayout.fitToHeight) ?? 0,
      )
      const options = [
        { value: '0', label: t('appFitAutomaticOption') },
        ...[1, 2, 3, 4, 5].map((pages) => ({
          value: String(pages),
          label: t(pages === 1 ? 'appFitPage1' : 'appFitPagesN', { count: pages }),
        })),
      ]
      return (
        <MenuSelect
          className="select-like compact"
          label={`Fit to ${axis}`}
          data-tip={axis === 'width' ? t('appWidth') : t('appHeight')}
          value={current}
          display={options.find((option) => option.value === current)?.label ?? current}
          options={options}
          onPick={(value) => onCommand(`page-layout:fit-${axis}:${value}`)}
        />
      )
    }
    const marginLabels = {
      normal: t('appMarginNormal'),
      wide: t('appMarginWide'),
      narrow: t('appMarginNarrow'),
    } as const
    return (
      <div className="ribbon" data-ribbon-body="">
        <RibbonGroup label={t('appGroupThemes')}>
          {largeMenu(
            t('appGroupThemes'),
            '🎨',
            t('appThemesTitle'),
            THEME_PRESETS.map((preset) => ({
              value: `page-layout:theme:${preset.id}`,
              label: preset.name,
            })),
          )}
          <div className="row-stack">
            <span className="styles-row" data-tip={t('appThemeColorsTitle')}>
              <ToolSymbol symbol="▤" />
              {t('appColors')}
              <CaretIcon />
              <MenuSelect
                cover
                label={t('appColors')}
                options={COLOR_SCHEMES.map((scheme) => ({
                  value: `page-layout:theme-colors:${scheme.id}`,
                  label: scheme.name,
                }))}
                onPick={onCommand}
              />
            </span>
            <span className="styles-row" data-tip={t('appThemeFontsTitle')}>
              <ToolSymbol symbol="A" />
              {t('appFonts')}
              <CaretIcon />
              <MenuSelect
                cover
                label={t('appFonts')}
                options={FONT_SCHEMES.map((scheme) => ({
                  value: `page-layout:theme-fonts:${scheme.id}`,
                  label: scheme.name,
                }))}
                onPick={onCommand}
              />
            </span>
          </div>
        </RibbonGroup>
        <RibbonGroup label={t('appGroupPageSetup')}>
          {largeMenu(
            t('appMargins'),
            '⿴',
            t('appMarginsTitle', {
              value: pageLayout.margins ? marginLabels[pageLayout.margins] : t('appAsSavedInFile'),
            }),
            [
              { value: 'page-layout:margins:normal', label: t('appMarginNormal') },
              { value: 'page-layout:margins:wide', label: t('appMarginWide') },
              { value: 'page-layout:margins:narrow', label: t('appMarginNarrow') },
            ],
          )}
          {largeMenu(
            t('appOrientationLabel'),
            '⤢',
            t('appOrientationTitle', {
              value: pageLayout.orientation
                ? t(pageLayout.orientation === 'portrait' ? 'appPortrait' : 'appLandscape')
                : t('appAsSavedInFile'),
            }),
            [
              { value: 'page-layout:orientation:portrait', label: t('appPortrait') },
              { value: 'page-layout:orientation:landscape', label: t('appLandscape') },
            ],
          )}
          {largeMenu(t('appSizeLabel'), '▭', t('appPaperSizeTitle'), [
            { value: 'page-layout:paper:1', label: 'Letter' },
            { value: 'page-layout:paper:5', label: 'Legal' },
            { value: 'page-layout:paper:3', label: 'Tabloid' },
            { value: 'page-layout:paper:7', label: 'Executive' },
            { value: 'page-layout:paper:8', label: 'A3' },
            { value: 'page-layout:paper:9', label: 'A4' },
            { value: 'page-layout:paper:11', label: 'A5' },
          ])}
          {largeMenu(
            t('appPrintArea'),
            '⬚',
            pageLayout.printArea
              ? t('appPrintAreaTitle', { area: pageLayout.printArea })
              : t('appPrintAreaFromSelection'),
            [
              { value: 'page-layout:print-area:set', label: t('appSetPrintArea') },
              { value: 'page-layout:print-area:clear', label: t('appClearPrintArea') },
            ],
          )}
          {largeMenu(t('appBreaks'), '┆', t('appBreaksTitle'), [
            { value: 'page-layout:breaks:insert', label: t('appInsertPageBreak') },
            { value: 'page-layout:breaks:remove', label: t('appRemovePageBreak') },
            { value: 'page-layout:breaks:reset', label: t('appResetAllPageBreaks') },
          ])}
          {largeMenu(
            t('appPrintTitlesLabel'),
            '▤',
            pageLayout.printTitles
              ? t('appPrintTitlesTitle', { rows: pageLayout.printTitles })
              : t('appPrintTitlesHint'),
            [
              { value: 'page-layout:print-titles:first-row', label: t('appRepeatRow1') },
              { value: 'page-layout:print-titles:selection', label: t('appRepeatSelectedRows') },
              { value: 'page-layout:print-titles:clear', label: t('appClearPrintTitles') },
            ],
          )}
        </RibbonGroup>
        <RibbonGroup label={t('appGroupScaleToFit')}>
          <div className="row-stack">
            <label className="styles-row" data-tip={t('appWidth')}>
              <ToolSymbol symbol="↔" />
              {t('appWidth')}:{fitOptions('width')}
            </label>
            <label className="styles-row" data-tip={t('appHeight')}>
              <ToolSymbol symbol="↕" />
              {t('appHeight')}:{fitOptions('height')}
            </label>
          </div>
        </RibbonGroup>
        <RibbonGroup label={t('appGroupSheetOptions')}>
          <div className="check-column">
            <span className="check-head">{t('appGridlines')}</span>
            <button
              className="check-item"
              data-tip={t('appGridlines')}
              onClick={() => onCommand('toggle-gridlines')}
            >
              <i className="check-box">{pageLayout.showGridlines ? '✓' : ''}</i>
              {t('appViewCheck')}
            </button>
            <button
              className="check-item"
              data-tip={t('appPrintGridlinesTitle')}
              onClick={() =>
                onCommand(`page-layout:print-gridlines:${pageLayout.printGridlines ? '0' : '1'}`)
              }
            >
              <i className="check-box">{pageLayout.printGridlines ? '✓' : ''}</i>
              {t('appPrintCheck')}
            </button>
          </div>
          <div className="check-column">
            <span className="check-head">{t('appHeadings')}</span>
            <button
              className="check-item"
              data-tip={t('appHeadings')}
              onClick={() => onCommand('toggle-headings')}
            >
              <i className="check-box">{pageLayout.showHeadings ? '✓' : ''}</i>
              {t('appViewCheck')}
            </button>
            <button
              className="check-item"
              data-tip={t('appPrintHeadingsTitle')}
              onClick={() =>
                onCommand(`page-layout:print-headings:${pageLayout.printHeadings ? '0' : '1'}`)
              }
            >
              <i className="check-box">{pageLayout.printHeadings ? '✓' : ''}</i>
              {t('appPrintCheck')}
            </button>
          </div>
        </RibbonGroup>
      </div>
    )
  }

  if (activeTab === 'Formulas') {
    const definedNames = onListNames()
    // Category buttons all open the same catalog dialog; the per-category
    // menus funnel into Insert Function.
    // Each button opens the catalog filtered to its own category; 'All'
    // for the ones the catalog has no counterpart for (Recently Used / More Functions).
    const functionCategory = (label: string, symbol: string, category: string) => (
      <RibbonButton
        large
        menu
        label={label}
        detail={category === 'All' ? t('appBrowseCatalog') : t('appBrowseCatalogFiltered')}
        symbol={symbol}
        onClick={() => onCommand(`insert-function-open:${category}`)}
      />
    )
    return (
      <div className="ribbon" data-ribbon-body="">
        <RibbonGroup label={t('appGroupFunctionLibrary')}>
          <RibbonButton
            large
            label={t('appInsertFunction')}
            detail={t('appBrowseCatalog')}
            symbol="ƒx"
            onClick={() => onCommand('insert-function-open')}
          />
          <div className="ribbon-tool large" data-tip={t('appAutoSumTitle')}>
            <span className="tool-icon-row">
              <ToolSymbol symbol="Σ" />
              <CaretIcon />
            </span>
            <span>
              <strong>{t('appAutoSum')}</strong>
            </span>
            <MenuSelect
              cover
              label="AutoSum"
              options={autoSumOptions(t)}
              onPick={(value) => onCommand(`autofn:${value}`)}
            />
          </div>
          {functionCategory(t('appFnCatRecentlyUsed'), '🕘', 'All')}
          {functionCategory(t('appFnCatFinancial'), '$', 'Financial')}
          {functionCategory(t('appFnCatLogical'), '?', 'Logical')}
          {functionCategory(t('appFnCatText'), 'A', 'Text')}
          {functionCategory(t('appFnCatDateTime'), '🕐', 'Date & Time')}
          {functionCategory(t('appFnCatLookup'), '🔍', 'Lookup')}
          {functionCategory(t('appFnCatMathTrig'), 'θ', 'Math')}
          {functionCategory(t('appFnCatMore'), '⋯', 'All')}
        </RibbonGroup>
        <RibbonGroup label={t('appGroupDefinedNames')}>
          <RibbonButton
            large
            label={t('appNameManager')}
            detail={t('appNameManagerDetail')}
            symbol="🏷"
            onClick={() => onCommand('name-manager-open')}
          />
          <div className="row-stack">
            <button
              className="styles-row as-button"
              data-tip={t('appDefineName')}
              onClick={() => onCommand('name-manager-open')}
            >
              <ToolSymbol symbol="🏷" />
              {t('appDefineName')}
              <CaretIcon />
            </button>
            <span className="styles-row" data-tip={t('appUseInFormulaTitle')}>
              <ToolSymbol symbol="ƒ" />
              {t('appUseInFormula')}
              <CaretIcon />
              <MenuSelect
                cover
                label={t('appUseInFormula')}
                options={
                  definedNames.length > 0
                    ? definedNames.map((name) => ({ value: `use-in-formula:${name}`, label: name }))
                    : [{ value: 'name-manager-open', label: t('appNoNamesYet') }]
                }
                onPick={onCommand}
              />
            </span>
            <span className="styles-row" data-tip={t('appCreateFromSelectionTitle')}>
              <ToolSymbol symbol="⊞" />
              {t('appCreateFromSelection')}
              <CaretIcon />
              <MenuSelect
                cover
                label={t('appCreateFromSelection')}
                options={[
                  { value: 'create-names:top', label: t('appCreateNamesTopRow') },
                  { value: 'create-names:left', label: t('appCreateNamesLeftCol') },
                ]}
                onPick={onCommand}
              />
            </span>
          </div>
        </RibbonGroup>
        <RibbonGroup label={t('appGroupFormulaAuditing')}>
          <div className="row-stack">
            <button
              className="styles-row as-button"
              data-tip={t('appTracePrecedentsTitle')}
              onClick={() => onCommand('trace-precedents')}
            >
              <ToolSymbol symbol="⇢" />
              {t('appTracePrecedents')}
            </button>
            <button
              className="styles-row as-button"
              data-tip={t('appTraceDependentsTitle')}
              onClick={() => onCommand('trace-dependents')}
            >
              <ToolSymbol symbol="⇠" />
              {t('appTraceDependents')}
            </button>
            <button
              className="styles-row as-button"
              data-tip={t('appRemoveArrowsTitle')}
              onClick={() => onCommand('remove-arrows')}
            >
              <ToolSymbol symbol="⌫" />
              {t('appRemoveArrows')}
            </button>
          </div>
          <RibbonButton
            large
            label={t('appShowFormulas')}
            detail={t('appShowFormulasDetail')}
            symbol="ƒ"
            onClick={() => onCommand('toggle-show-formulas')}
          />
          <RibbonButton
            large
            label={t('appErrorChecking')}
            detail={t('appErrorCheckingDetail')}
            symbol="⚠"
            onClick={() => onCommand('error-checking')}
          />
          <RibbonButton
            large
            label={t('appWatchWindow')}
            detail={t('appWatchWindowDetail')}
            symbol="👓"
            onClick={() => onCommand('watch-window')}
          />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupCalculation')}>
          {largeMenu(t('appCalculationOptions'), '🧮', t('appCalculationOptionsTitle'), [
            { value: 'calc-mode:auto', label: t('appCalcAuto') + (calcManual ? '' : ' ✓') },
            { value: 'calc-mode:manual', label: t('appCalcManual') + (calcManual ? ' ✓' : '') },
          ])}
          <div className="row-stack">
            <button
              className="styles-row as-button"
              data-tip={t('appCalculateNowTitle')}
              onClick={() => onCommand('calculate-now')}
            >
              <ToolSymbol symbol="⟳" />
              {t('appCalculateNow')}
            </button>
            <button
              className="styles-row as-button"
              data-tip={t('appCalculateSheetTitle')}
              onClick={() => onCommand('calculate-sheet')}
            >
              <ToolSymbol symbol="▦" />
              {t('appCalculateSheet')}
            </button>
          </div>
        </RibbonGroup>
      </div>
    )
  }

  if (activeTab === 'Data') {
    return (
      <div className="ribbon" data-ribbon-body="">
        <RibbonGroup label={t('appPivotTable')}>
          <RibbonButton
            large
            label={t('appPivotTable')}
            detail={t('appOoxmlPivot')}
            symbol="⊞"
            onClick={() => onCommand('pivot-open')}
          />
          <RibbonButton
            large
            label={t('appRefresh')}
            detail={onIsSelectionInPivot() ? t('appRefreshHintIn') : t('appRefreshHintOut')}
            symbol="⟳"
            onClick={() => {
              const err = onRefreshPivot()
              if (err) onCommand(`error:${err}`)
            }}
          />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupGetData')}>
          <div className="row-stack">
            <button
              className="styles-row as-button"
              data-tip={t('appFromTextCsvTitle')}
              onClick={() => onCommand('import-csv')}
            >
              <ToolSymbol symbol="🗎" />
              {t('appFromTextCsv')}
            </button>
            <button
              className="styles-row as-button"
              data-tip={t('appMergeWorkbooksTip')}
              onClick={() => onCommand('merge-workbooks')}
            >
              <ToolSymbol symbol="⧉" />
              {t('appMergeWorkbooks')}
            </button>
            <button
              className="styles-row as-button"
              data-tip={t('appRefreshAllTitle')}
              onClick={() => onCommand('refresh-all')}
            >
              <ToolSymbol symbol="⟳" />
              {t('appRefreshAll')}
            </button>
          </div>
        </RibbonGroup>
        <RibbonGroup label={t('appGroupSortFilter')}>
          {largeMenu(t('appSort'), '⇅', t('appSortSelectionTitle'), [
            { value: 'sort:asc', label: t('appSortAToZ') },
            { value: 'sort:desc', label: t('appSortZToA') },
            { value: 'sort-custom-open', label: t('appCustomSort') },
          ])}
          <RibbonButton
            large
            label={t('appFilter')}
            detail={platformShortcuts(t('appFilterToggleDetail'))}
            symbol="▽"
            onClick={() => onCommand('filter-toggle')}
          />
          <div className="row-stack">
            <button
              className="styles-row as-button"
              data-tip={t('appClearFilterTitle')}
              onClick={() => onCommand('filter-clear')}
            >
              <ToolSymbol symbol="⊘" />
              {t('appClear')}
            </button>
            <button
              className="styles-row as-button"
              data-tip={t('appReapplyTitle')}
              onClick={() => onCommand('filter-reapply')}
            >
              <ToolSymbol symbol="↻" />
              {t('appReapply')}
            </button>
            <button
              className="styles-row as-button"
              data-tip={t('appAdvancedFilterTitle')}
              onClick={() => onCommand('filter-advanced')}
            >
              <ToolSymbol symbol="▽" />
              {t('appAdvanced')}
            </button>
          </div>
        </RibbonGroup>
        <RibbonGroup label={t('appGroupDataTools')}>
          {largeMenu(t('appTextToColumns'), '⇶', t('appTextToColumnsTitle'), [
            { value: 'text-to-columns:2', label: t('appSplitByComma') },
            { value: 'text-to-columns:4', label: t('appSplitBySemicolon') },
            { value: 'text-to-columns:8', label: t('appSplitBySpace') },
            { value: 'text-to-columns:1', label: t('appSplitByTab') },
          ])}
          <RibbonButton
            large
            label={t('appFlashFill')}
            detail={t('appFlashFillDetail')}
            symbol="⚡"
            onClick={() => onCommand('flash-fill')}
          />
          <RibbonButton
            large
            label={t('appRemoveDuplicates')}
            detail={t('appInSelection')}
            symbol="⧉"
            onClick={() => onCommand('remove-duplicates-open')}
          />
          <RibbonButton
            large
            menu
            label={t('appDataValidation')}
            detail={t('appRulesPanel')}
            symbol="✓"
            onClick={() => onCommand('dv-open')}
          />
          <RibbonButton
            large
            label={t('appConsolidate')}
            detail={t('appCombineRanges')}
            symbol="⊕"
            onClick={() => onCommand('consolidate-open')}
          />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupForecast')}>
          {largeMenu(t('appWhatIfAnalysis'), '❔', t('appWhatIfTitle'), [
            { value: 'goal-seek-open', label: t('appGoalSeek') },
          ])}
        </RibbonGroup>
        <RibbonGroup label={t('appGroupOutline')}>
          {largeMenu(t('appOutlineGroup'), '⊟', t('appOutlineGroupTitle'), [
            { value: 'outline-group:rows', label: t('appGroupRows') },
            { value: 'outline-group:cols', label: t('appGroupCols') },
            { value: 'outline-hide-detail:rows', label: t('appHideDetailRows') },
            { value: 'outline-hide-detail:cols', label: t('appHideDetailCols') },
            {
              value: 'outline-summary-below',
              label: `${outlineSummary.below ? '✓ ' : ''}${t('appOutlineSummaryBelow')}`,
            },
            {
              value: 'outline-summary-right',
              label: `${outlineSummary.right ? '✓ ' : ''}${t('appOutlineSummaryRight')}`,
            },
          ])}
          {largeMenu(t('appOutlineUngroup'), '⊞', t('appOutlineUngroupTitle'), [
            { value: 'outline-ungroup:rows', label: t('appUngroupRows') },
            { value: 'outline-ungroup:cols', label: t('appUngroupCols') },
            { value: 'outline-show-detail:rows', label: t('appShowDetailRows') },
            { value: 'outline-show-detail:cols', label: t('appShowDetailCols') },
            { value: 'outline-clear', label: t('appClearOutline') },
          ])}
          <RibbonButton
            large
            label={t('appSubtotal')}
            detail={t('appSubtotalDetail')}
            symbol="∑"
            onClick={() => onCommand('subtotal-open')}
          />
        </RibbonGroup>
      </div>
    )
  }

  if (activeTab === 'View') {
    return (
      <div className="ribbon" data-ribbon-body="">
        <RibbonGroup label={t('appGroupWorkbookViews')}>
          <RibbonButton
            large
            label={t('appNormalView')}
            detail={t('appCurrentViewTitle')}
            symbol="▦"
            active={pageLayout.pageBreakPreview !== true}
            onClick={() => {
              if (pageLayout.pageBreakPreview === true) onCommand('toggle-page-break-preview')
            }}
          />
          <RibbonButton
            large
            label={t('appPageBreakPreview')}
            detail={t('appPageBreakPreviewTitle')}
            symbol="┆"
            active={pageLayout.pageBreakPreview === true}
            onClick={() => {
              if (pageLayout.pageBreakPreview !== true) onCommand('toggle-page-break-preview')
            }}
          />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupShow')}>
          <div className="check-column check-column-view">
            <button
              className="check-item"
              data-tip={t('appGridlines')}
              onClick={() => onCommand('toggle-gridlines')}
            >
              <i className="check-box">{pageLayout.showGridlines ? '✓' : ''}</i>
              {t('appGridlines')}
            </button>
            <button
              className="check-item"
              data-tip={t('appFormulaBar')}
              onClick={() => onCommand('toggle-formula-bar')}
            >
              <i className="check-box">{formulaBarVisible ? '✓' : ''}</i>
              {t('appFormulaBar')}
            </button>
            <button
              className="check-item"
              data-tip={t('appHeadings')}
              onClick={() => onCommand('toggle-headings')}
            >
              <i className="check-box">{pageLayout.showHeadings ? '✓' : ''}</i>
              {t('appHeadings')}
            </button>
            <button
              className="check-item"
              data-tip={t('appRightToLeft')}
              onClick={() => onCommand('toggle-right-to-left')}
            >
              <i className="check-box">{pageLayout.rightToLeft ? '✓' : ''}</i>
              {t('appRightToLeft')}
            </button>
            <button
              className="check-item"
              data-tip={t('appCrossHighlight')}
              onClick={() => onCommand('toggle-cross-highlight')}
            >
              <i className="check-box">{crossHighlightVisible ? '✓' : ''}</i>
              {t('appCrossHighlight')}
            </button>
          </div>
        </RibbonGroup>
        <RibbonGroup label={t('appZoomLabel')}>
          {largeMenu(t('appZoomLabel'), '🔍', t('appZoomSheetTitle'), [
            { value: 'zoom:50', label: '50%' },
            { value: 'zoom:75', label: '75%' },
            { value: 'zoom:100', label: '100%' },
            { value: 'zoom:125', label: '125%' },
            { value: 'zoom:150', label: '150%' },
            { value: 'zoom:200', label: '200%' },
          ])}
          <RibbonButton
            large
            label="100%"
            detail={t('appResetZoom')}
            symbol="⊙"
            onClick={() => onCommand('zoom-reset')}
          />
          <RibbonButton
            large
            label={t('appZoomToSelection')}
            detail={t('appZoomToSelectionDetail')}
            symbol="⌖"
            onClick={() => onCommand('zoom-to-selection')}
          />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupWindow')}>
          {largeMenu(t('appFreezePanes'), '❄', t('appFreezeTitle'), [
            { value: 'freeze-here', label: t('appFreezeAtSelection') },
            { value: 'freeze-top-row', label: t('appFreezeTopRow') },
            { value: 'freeze-first-col', label: t('appFreezeFirstCol') },
            { value: 'unfreeze', label: t('appUnfreeze') },
          ])}
        </RibbonGroup>
      </div>
    )
  }

  if (activeTab === 'Review') {
    return (
      <div className="ribbon" data-ribbon-body="">
        <RibbonGroup label={t('appGroupProofing')}>
          <RibbonButton
            large
            label={t('appWorkbookStatsLabel')}
            detail={t('appSheetsCellsFormulas')}
            symbol="🧮"
            onClick={() => onCommand('workbook-statistics')}
          />
        </RibbonGroup>
        <RibbonGroup label={t('appGroupLanguage')}>
          <div className="ribbon-tool large" data-tip={t('appTranslateTitle')}>
            <span className="tool-icon-row">
              <ToolSymbol symbol="文" />
              <CaretIcon />
            </span>
            <span>
              <strong>{t('appTranslate')}</strong>
            </span>
            <MenuSelect
              cover
              label={t('appTranslate')}
              options={TRANSLATE_LANGUAGES.map((language) => ({
                value: language,
                label: language,
              }))}
              onPick={(language) => onAiRun(t('appTranslatePrompt', { language }))}
            />
          </div>
        </RibbonGroup>
        <RibbonGroup label={t('appGroupComments')}>
          <RibbonButton
            large
            label={t('appNewComment')}
            detail={t('appNewCommentDetail')}
            symbol="🗨"
            onClick={() => onCommand('comment-new')}
          />
          <RibbonButton
            large
            label={t('appDeleteLabel')}
            detail={t('appCommentAtSelection')}
            symbol="🗑"
            onClick={() => onCommand('comment-delete')}
          />
          <div className="row-stack">
            <button
              className="styles-row as-button"
              data-tip={t('appNotePrevTitle')}
              onClick={() => onCommand('comment-prev')}
            >
              <ToolSymbol symbol="←" />
              {t('appPrevious')}
            </button>
            <button
              className="styles-row as-button"
              data-tip={t('appNoteNextTitle')}
              onClick={() => onCommand('comment-next')}
            >
              <ToolSymbol symbol="→" />
              {t('appNext')}
            </button>
            <button
              className="styles-row as-button"
              data-tip={t('appShowCommentsPaneTitle')}
              onClick={() => onCommand('comments-pane-toggle')}
            >
              <ToolSymbol symbol="🗨" />
              {t('appShowComments')}
            </button>
          </div>
        </RibbonGroup>
        <RibbonGroup label={t('appGroupNotes')}>
          {largeMenu(t('appGroupNotes'), '🗒', t('appNotesDetail'), [
            { value: 'note-open', label: t('appNewNote') },
            { value: 'note-show-toggle', label: t('appShowHideNote') },
            { value: 'note-show-all', label: t('appShowAllNotes') },
            { value: 'notes-convert', label: t('appConvertToNotes') },
            { value: 'note-pictures-pane-toggle', label: t('appNotePicturesToggle') },
          ])}
        </RibbonGroup>
        <RibbonGroup label={t('appGroupProtection')}>
          <RibbonButton
            large
            label={t(sheetProtected ? 'appUnprotectSheet' : 'appProtectSheet')}
            detail={t(sheetProtected === null ? 'appOpenFileFirst' : 'appProtectSheetDetail')}
            symbol={sheetProtected ? '🔓' : '🔒'}
            onClick={() => onCommand('sheet-protect')}
          />
          <RibbonButton
            large
            label={t(workbookProtected ? 'appUnprotectWorkbook' : 'appProtectWorkbook')}
            detail={t(workbookProtected === null ? 'appOpenFileFirst' : 'appProtectWorkbookTitle')}
            symbol={workbookProtected ? '🔓' : '🔐'}
            onClick={() => onCommand('workbook-protect')}
          />
          <RibbonButton
            large
            label={t('appAllowEditRanges')}
            detail={t('appAllowEditRangesTitle')}
            symbol="⬚"
            onClick={() => onCommand('allow-edit-ranges-open')}
          />
        </RibbonGroup>
      </div>
    )
  }

  const fontSizes = [9, 10, 11, 12, 14, 16, 18, 22, 26]
  const echoFamily = selectionFormat?.fontFamily ?? 'Aptos'
  const echoSize = selectionFormat?.fontSize ?? 11
  const fontGroups = fontFamilyGroups(systemFontFamilies, echoFamily)
  const familyOptions = [
    ...fontGroups.common.map((family) => ({ value: family, label: family })),
    ...fontGroups.system.map((family, index) => ({
      value: family,
      label: family,
      sep: index === 0,
    })),
  ]
  const sizeOptions = fontSizes.includes(echoSize)
    ? fontSizes
    : [...fontSizes, echoSize].sort((a, b) => a - b)
  return (
    <div className="ribbon" data-ribbon-body="">
      <RibbonGroup label={t('appGroupAiAssistant')}>
        <button
          className={`ribbon-tool as-button large ai-entry ${aiOpen ? 'active' : ''}`}
          data-tip={t('aiOpenAssistant')}
          onClick={onAiToggle}
        >
          <span className="tool-icon-row">
            <GensparkMark size={26} />
          </span>
          <span>
            <strong>Genspark AI</strong>
          </span>
        </button>
        <button
          className="ribbon-tool as-button large ai-entry"
          disabled={!sheetHasContent}
          data-tip={t('aiCheckBtn')}
          onClick={() => onAiRun(t('aiCheckPrompt'))}
        >
          <span className="tool-icon-row">
            <span className="ai-feature-icon" aria-hidden="true">
              <svg
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="1.5"
                strokeLinecap="round"
                strokeLinejoin="round"
              >
                <path d="M11 3.25C15.2802 3.25 18.75 6.71979 18.75 11C18.75 15.2802 15.2802 18.75 11 18.75C6.71979 18.75 3.25 15.2802 3.25 11C3.25 6.71979 6.71979 3.25 11 3.25Z" />
                <path
                  d="M7.5 10.8235L9.64097 12.9645C9.93755 13.2611 10.4177 13.2634 10.7171 12.9697L14.7647 9"
                  strokeLinecap="round"
                />
                <path d="M20 20.5L16.5 17" strokeLinecap="round" />
              </svg>
            </span>
          </span>
          <span>
            <strong>{t('aiCheckBtn')}</strong>
          </span>
        </button>
        <button
          className="ribbon-tool as-button large ai-entry"
          disabled={!sheetHasContent}
          data-tip={t('aiAnalyzeBtn')}
          onClick={() => onAiRun(t('aiAnalyzePrompt'))}
        >
          <span className="tool-icon-row">
            <span className="ai-feature-icon" aria-hidden="true">
              <svg
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="1.5"
                strokeLinecap="round"
                strokeLinejoin="round"
              >
                <path d="M3.88589 14.2073H8.48682" strokeLinecap="round" />
                <path d="M3.88589 19.0112H8.48682" strokeLinecap="round" />
                <path d="M3.88589 9.40369H11.692" strokeLinecap="round" />
                <path d="M3.88589 4.59998H19.1645" strokeLinecap="round" />
                <path d="M15.1995 10.5445C15.3784 10.0908 16.0206 10.0908 16.1996 10.5445L16.706 11.8286C17.0338 12.6598 17.6918 13.3178 18.523 13.6456L19.8071 14.1521C20.2608 14.331 20.2608 14.9732 19.8071 15.1522L18.523 15.6586C17.6918 15.9864 17.0338 16.6444 16.706 17.4756L16.1996 18.7597C16.0206 19.2134 15.3784 19.2134 15.1995 18.7597L14.693 17.4756C14.3652 16.6444 13.7072 15.9864 12.876 15.6586L11.592 15.1522C11.1382 14.9732 11.1382 14.331 11.592 14.1521L12.876 13.6456C13.7072 13.3178 14.3652 12.6598 14.693 11.8286L15.1995 10.5445Z" />
              </svg>
            </span>
          </span>
          <span>
            <strong>{t('aiAnalyzeBtn')}</strong>
          </span>
        </button>
      </RibbonGroup>
      <RibbonGroup label={t('appGroupClipboard')}>
        <button
          className="ribbon-tool as-button large"
          data-tip={t('appPasteTitle')}
          onClick={() => onCommand('paste')}
        >
          <span className="tool-icon-row">
            <ToolSymbol symbol="📋" />
          </span>
          <span>
            <strong>{t('appPaste')}</strong>
          </span>
        </button>
        {largeMenu(t('appPasteSpecial'), '📑', t('appPasteSpecialTitle'), [
          { value: 'paste-special:value', label: t('appPasteValuesOnly') },
          { value: 'paste-special:formula', label: t('appPasteFormulasOnly') },
          { value: 'paste-special:format', label: t('appPasteFormattingOnly') },
          { value: 'paste-special:col-width', label: t('appPasteColWidths') },
          { value: 'paste-special:besides-border', label: t('appPasteExceptBorders') },
          { value: PASTE_SPECIAL_OPEN_SHELL_COMMAND, label: t('appPasteSpecialDialog') },
        ])}
        <div className="tool-stack">
          <button
            data-tip={t('appCutTitle')}
            aria-label={t('appCutTitle')}
            onClick={() => onCommand('cut')}
          >
            <ToolSymbol symbol="✂" />
          </button>
          <button
            data-tip={t('appCopyTitle')}
            aria-label={t('appCopyTitle')}
            onClick={() => onCommand('copy')}
          >
            <ToolSymbol symbol="⧉" />
          </button>
          <button
            data-tip={t('appFormatPainterTip')}
            aria-label={t('appFormatPainter')}
            aria-pressed={formatPainterActive}
            className={formatPainterActive ? 'is-active' : ''}
            onClick={(event) =>
              onCommand(event.detail >= 2 ? 'format-painter:dblclick' : 'format-painter')
            }
          >
            <ToolSymbol symbol="🖌" />
          </button>
        </div>
      </RibbonGroup>
      <RibbonGroup label={t('appGroupFont')}>
        <div className="ribbon-rows">
          <div className="inline-tools">
            <EditableMenuSelect
              className="select-like font-name"
              label="Font family"
              data-tip={t('appFontFamilyTip')}
              value={echoFamily}
              options={familyOptions}
              onOpen={loadSystemFonts}
              onPick={(value) => onCommand(`font-family:${value}`)}
              commit={(text) => onCommand(`font-family:${text}`)}
            />
            <EditableMenuSelect
              className="select-like font-size"
              label="Font size"
              data-tip={t('appFontSizeTip')}
              value={String(echoSize)}
              options={sizeOptions.map((size) => ({ value: String(size), label: String(size) }))}
              onPick={(value) => onCommand(`font-size:${value}`)}
              commit={(text) => {
                const size = Number(text.replace(',', '.'))
                // Excel's font-size bounds
                if (Number.isFinite(size) && size >= 1 && size <= 409)
                  onCommand(`font-size:${size}`)
              }}
            />
            <button
              data-tip={t('appIncreaseFontSize')}
              aria-label={t('appIncreaseFontSize')}
              onClick={() => onCommand(`font-size:${stepFontSize(echoSize, 1)}`)}
            >
              <ToolSymbol symbol="A↑" />
            </button>
            <button
              data-tip={t('appDecreaseFontSize')}
              aria-label={t('appDecreaseFontSize')}
              onClick={() => onCommand(`font-size:${stepFontSize(echoSize, -1)}`)}
            >
              <ToolSymbol symbol="A↓" />
            </button>
          </div>
          <div className="inline-tools">
            <button
              data-tip={t('appBold')}
              className={selectionFormat?.bold ? 'is-active' : ''}
              onClick={() => onCommand('bold')}
            >
              <b>B</b>
            </button>
            <button
              data-tip={t('appItalic')}
              className={selectionFormat?.italic ? 'is-active' : ''}
              onClick={() => onCommand('italic')}
            >
              <em>I</em>
            </button>
            <button
              data-tip={t('appUnderline')}
              className={selectionFormat?.underline ? 'is-active' : ''}
              onClick={() => onCommand('underline')}
            >
              <u>U</u>
            </button>
            <button
              data-tip={t('appDoubleUnderline')}
              onClick={() => onCommand('underline:double')}
            >
              <u style={{ textDecorationStyle: 'double' }}>D</u>
            </button>
            <button
              data-tip={t('appStrikethrough')}
              className={selectionFormat?.strike ? 'is-active' : ''}
              onClick={() => onCommand('strike')}
            >
              <s>S</s>
            </button>
            <button
              data-tip={t('appSuperscript')}
              aria-label={t('appSuperscript')}
              className={selectionFormat?.vertAlign === 'superscript' ? 'is-active' : ''}
              onClick={() => onCommand('superscript')}
            >
              <ToolSymbol symbol="x²" />
            </button>
            <button
              data-tip={t('appSubscript')}
              aria-label={t('appSubscript')}
              className={selectionFormat?.vertAlign === 'subscript' ? 'is-active' : ''}
              onClick={() => onCommand('subscript')}
            >
              <ToolSymbol symbol="x₂" />
            </button>
            <ColorDropdown
              label="Font color"
              data-tip={t('appFontColor')}
              split
              display={
                <span className="swatch-letter">
                  A<i style={{ background: fontColor }} />
                </span>
              }
              value={fontColor}
              auto={t('appAutomaticColor')}
              onPick={(hex) => {
                setFontColor(hex ?? '#000000')
                onCommand(hex ? `font-color:${hex}` : 'font-color:auto')
              }}
            />
            <ColorDropdown
              label="Fill color"
              data-tip={t('appFillColor')}
              split
              display={
                <span className="swatch-letter">
                  <ToolSymbol symbol="◧" />
                  <i style={{ background: fillColor }} />
                </span>
              }
              value={fillColor}
              auto={t('dlgFcNoFill')}
              onPick={(hex) => {
                if (hex) setFillColor(hex)
                onCommand(hex ? `fill:${hex}` : 'fill:none')
              }}
            />
            <MenuSelect
              className="select-like compact"
              label="Borders"
              data-tip={t('appBorders')}
              display={
                <>
                  <ToolSymbol symbol="⊡" /> {t('appBorders')}
                </>
              }
              options={[
                { value: 'all', label: t('appBorderAll'), icon: <BorderAllIcon /> },
                { value: 'outer', label: t('appBorderOuter'), icon: <BorderOuterIcon /> },
                {
                  value: 'thick-outer',
                  label: t('appBorderThickOuter'),
                  icon: <BorderThickOuterIcon />,
                },
                { value: 'top', label: t('appBorderTop'), icon: <BorderTopIcon /> },
                { value: 'bottom', label: t('appBorderBottom'), icon: <BorderBottomIcon /> },
                { value: 'left', label: t('appBorderLeft'), icon: <BorderLeftIcon /> },
                { value: 'right', label: t('appBorderRight'), icon: <BorderRightIcon /> },
                { value: 'none', label: t('appBorderNone'), icon: <BorderNoneIcon /> },
              ]}
              onPick={(value) => onCommand(`border:${value}:${borderColor}`)}
            />
            <ColorDropdown
              label="Border color"
              data-tip={t('appBorderColor')}
              display={
                <span className="swatch-letter">
                  <ToolSymbol symbol="⊡" />
                  <i style={{ background: borderColor }} />
                </span>
              }
              value={borderColor}
              onPick={(hex) => {
                if (hex) setBorderColor(hex)
              }}
            />
          </div>
        </div>
      </RibbonGroup>
      <RibbonGroup label={t('appGroupAlignment')}>
        <div className="ribbon-rows">
          <div className="inline-tools alignment-tools">
            <button
              data-tip={t('appTopAlign')}
              aria-label={t('appTopAlign')}
              onClick={() => onCommand('valign:top')}
            >
              <ToolSymbol symbol="⤒" />
            </button>
            <button
              data-tip={t('appMiddleAlign')}
              aria-label={t('appMiddleAlign')}
              onClick={() => onCommand('valign:middle')}
            >
              <ToolSymbol symbol="↕" />
            </button>
            <button
              data-tip={t('appBottomAlign')}
              aria-label={t('appBottomAlign')}
              onClick={() => onCommand('valign:bottom')}
            >
              <ToolSymbol symbol="⤓" />
            </button>
            <button
              data-tip={t('dlgFcWrapText')}
              aria-label={t('dlgFcWrapText')}
              className={selectionFormat?.wrap ? 'is-active' : ''}
              onClick={() => onCommand('wrap')}
            >
              <ToolSymbol symbol="↩" />
            </button>
            <MenuSelect
              className="select-like compact"
              label="Orientation"
              data-tip={t('dlgFcOrientation')}
              display={<ToolSymbol symbol="⤴" />}
              options={[
                { value: '45', label: t('appAngleCcw') },
                { value: '-45', label: t('appAngleCw') },
                { value: 'vertical', label: t('appVerticalText') },
                { value: '90', label: t('appRotateUp') },
                { value: '-90', label: t('appRotateDown') },
                { value: '0', label: t('appClearRotation') },
              ]}
              onPick={(value) => onCommand(`rotate:${value}`)}
            />
          </div>
          <div className="inline-tools alignment-tools">
            <button
              data-tip={t('appAlignLeft')}
              aria-label={t('appAlignLeft')}
              onClick={() => onCommand('align:left')}
            >
              <ToolSymbol symbol="≡" />
            </button>
            <button
              data-tip={t('appAlignCenter')}
              aria-label={t('appAlignCenter')}
              onClick={() => onCommand('align:center')}
            >
              <ToolSymbol symbol="≣" />
            </button>
            <button
              data-tip={t('appAlignRight')}
              aria-label={t('appAlignRight')}
              onClick={() => onCommand('align:right')}
            >
              <ToolSymbol symbol="☰" />
            </button>
            <button
              data-tip={t('appDecreaseIndent')}
              aria-label={t('appDecreaseIndent')}
              disabled={!selectionFormat?.indent}
              onClick={() => onCommand('indent-step:-1')}
            >
              <ToolSymbol symbol="⇤" />
            </button>
            <button
              data-tip={t('appIncreaseIndent')}
              aria-label={t('appIncreaseIndent')}
              disabled={(selectionFormat?.indent ?? 0) >= 250}
              onClick={() => onCommand('indent-step:1')}
            >
              <ToolSymbol symbol="⇥" />
            </button>
            <MenuSelect
              className="select-like compact"
              label="Merge cells"
              data-tip={t('appMergeCells')}
              display={
                <>
                  <ToolSymbol symbol="⇔" /> {t('appMerge')}
                </>
              }
              options={[
                { value: 'center', label: t('appMergeCenter') },
                { value: 'across', label: t('appMergeAcross') },
                { value: 'cells', label: t('appMergeCells') },
                { value: 'unmerge', label: t('appUnmergeCells') },
              ]}
              onPick={(value) => onCommand(`merge:${value}`)}
            />
          </div>
        </div>
      </RibbonGroup>
      <RibbonGroup label={t('appGroupNumber')}>
        <div className="ribbon-rows">
          <NumberFormatSelect pattern={selectionFormat?.numberFormat ?? ''} onCommand={onCommand} />
          <div className="inline-tools">
            <button
              data-tip={t('appCurrency')}
              aria-label={t('appCurrency')}
              onClick={() => onCommand(`format:${accountingStyle(RIBBON_CURRENCY_SYMBOL)}`)}
            >
              $
            </button>
            <button
              data-tip={t('appPercentTitle')}
              aria-label={t('appPercentTitle')}
              onClick={() => onCommand(`format:${PERCENT_STYLE}`)}
            >
              %
            </button>
            <button
              data-tip={t('appThousands')}
              aria-label={t('appThousands')}
              onClick={() => onCommand(`format:${COMMA_STYLE}`)}
            >
              ,
            </button>
            <button
              data-tip={t('appIncreaseDecimal')}
              aria-label={t('appIncreaseDecimal')}
              onClick={() => onCommand('decimal-inc')}
            >
              .0+
            </button>
            <button
              data-tip={t('appDecreaseDecimal')}
              aria-label={t('appDecreaseDecimal')}
              onClick={() => onCommand('decimal-dec')}
            >
              .0−
            </button>
          </div>
        </div>
      </RibbonGroup>
      <RibbonGroup label={t('appGroupStyles')}>
        <div className="styles-stack">
          <button
            className="styles-row as-button"
            data-tip={t('appConditionalFormatting')}
            onClick={() => onCommand('cf-open')}
          >
            <ToolSymbol symbol="▤" />
            {t('appConditionalFormatting')}
            <CaretIcon />
          </button>
          <div className="styles-row as-button" data-tip={t('appFormatAsTableTitle')}>
            <ToolSymbol symbol="▦" />
            {t('appFormatAsTable')}
            <CaretIcon />
            <MenuSelect
              cover
              label="Format as Table"
              options={[
                { value: 'TableStyleLight1', label: t('appTableStyleLight1') },
                { value: 'TableStyleLight9', label: t('appTableStyleLight9') },
                { value: 'TableStyleMedium2', label: t('appTableStyleMedium2') },
                { value: 'TableStyleMedium4', label: t('appTableStyleMedium4') },
                { value: 'TableStyleMedium7', label: t('appTableStyleMedium7') },
                { value: 'TableStyleDark2', label: t('appTableStyleDark2') },
              ]}
              onPick={(value) => onCommand(`format-as-table:${value}`)}
            />
          </div>
          <div className="styles-row as-button" data-tip={t('appCellStylesTitle')}>
            <ToolSymbol symbol="🎨" />
            {t('appCellStyles')}
            <CaretIcon />
            <MenuSelect
              cover
              label="Cell Styles"
              options={[
                { value: 'good', label: t('appStyleGood') },
                { value: 'bad', label: t('appStyleBad') },
                { value: 'neutral', label: t('appStyleNeutral') },
                { value: 'input', label: t('appStyleInput') },
                { value: 'output', label: t('appStyleOutput') },
                { value: 'calculation', label: t('appStyleCalculation') },
                { value: 'warning-text', label: t('appStyleWarningText') },
                { value: 'title', label: t('appTitle') },
                { value: 'heading-1', label: t('appStyleHeading1') },
                { value: 'heading-2', label: t('appStyleHeading2') },
                { value: 'total', label: t('appStyleTotal') },
                { value: 'accent1-20', label: t('appStyleAccent120') },
                { value: 'accent1-40', label: t('appStyleAccent140') },
                { value: 'accent1', label: t('appStyleAccent1') },
              ]}
              onPick={(value) => onCommand(`cell-style:${value}`)}
            />
          </div>
        </div>
      </RibbonGroup>
      <RibbonGroup label={t('appGroupCells')}>
        <div className="ribbon-rows">
          <button
            className="styles-row as-button"
            data-tip={t('appFormatCells')}
            data-tip-kbd={platformShortcuts('⌘1')}
            onClick={() => onCommand('format-cells')}
          >
            <ToolSymbol symbol="🎨" />
            {t('appFormatCells')} {platformShortcuts('⌘1')}
            <CaretIcon />
          </button>
          <div className="inline-tools cell-tools">
            <button
              data-tip={t('appInsertRow')}
              aria-label={t('appInsertRow')}
              onClick={() => onCommand('insert-row-here')}
            >
              <ToolSymbol symbol="⤒" />
            </button>
            <button
              data-tip={t('appDeleteRow')}
              aria-label={t('appDeleteRow')}
              onClick={() => onCommand('delete-row-here')}
            >
              <ToolSymbol symbol="⤓" />
            </button>
            <button
              data-tip={t('appInsertCol')}
              aria-label={t('appInsertCol')}
              onClick={() => onCommand('insert-col-here')}
            >
              <ToolSymbol symbol="⇤" />
            </button>
            <button
              data-tip={t('appDeleteCol')}
              aria-label={t('appDeleteCol')}
              onClick={() => onCommand('delete-col-here')}
            >
              <ToolSymbol symbol="⇥" />
            </button>
            <button
              data-tip={t('appInsertCells')}
              data-tip-kbd={platformShortcuts('⇧⌘=')}
              aria-label={t('appInsertCells')}
              onClick={() => onCommand('insert-cells')}
            >
              <ToolSymbol symbol="⊞" />
            </button>
            <button
              data-tip={t('appDeleteCells')}
              data-tip-kbd={platformShortcuts('⌘-')}
              aria-label={t('appDeleteCells')}
              onClick={() => onCommand('delete-cells')}
            >
              <ToolSymbol symbol="⊟" />
            </button>
            <MenuSelect
              className="select-like compact"
              label="Format"
              data-tip={t('appFormatMenu')}
              display={
                <>
                  <ToolSymbol symbol="⇳" /> {t('appFormatMenu')}
                </>
              }
              options={[
                { value: 'row-height-open', label: `${t('appRowHeight')}…` },
                { value: 'autofit-row-height', label: t('appAutoFitRowHeight') },
                { value: 'col-width-open', label: `${t('appColWidth')}…` },
                { value: 'autofit-col-width', label: t('appAutoFitColWidth') },
              ]}
              onPick={(value) => onCommand(value)}
            />
          </div>
        </div>
      </RibbonGroup>
      <RibbonGroup label={t('appGroupEditing')}>
        <div className="ribbon-rows">
          <div className="inline-tools">
            <MenuSelect
              className="select-like compact"
              label="AutoSum"
              data-tip={t('appAutoSumTitle')}
              display={
                <>
                  <ToolSymbol symbol="Σ" /> {t('appAutoSum')}
                </>
              }
              options={autoSumOptions(t)}
              onPick={(value) => onCommand(`autofn:${value}`)}
            />
            <MenuSelect
              className="select-like compact"
              label="Sort & Filter"
              data-tip={t('appGroupSortFilter')}
              display={
                <>
                  <ToolSymbol symbol="⇅" /> {t('appGroupSortFilter')}
                </>
              }
              options={[
                { value: 'sort:asc', label: t('appSortAToZ') },
                { value: 'sort:desc', label: t('appSortZToA') },
                { value: 'sort-custom-open', label: t('appCustomSort') },
                { value: 'filter-toggle', label: t('appFilter') },
                { value: 'filter-clear', label: t('appClearFilterTitle') },
                { value: 'filter-reapply', label: t('appReapplyTitle') },
              ]}
              onPick={(value) => onCommand(value)}
            />
            <MenuSelect
              className="select-like compact"
              label="Fill"
              data-tip={t('appFillMenu')}
              display={
                <>
                  <ToolSymbol symbol="↓" /> {t('appFillMenu')}
                </>
              }
              options={[
                { value: 'fill-down', label: t('appFillDown') },
                { value: 'fill-right', label: t('appFillRight') },
                { value: 'fill-up', label: t('appFillUp') },
                { value: 'fill-left', label: t('appFillLeft') },
                { value: 'fill-series-open', label: t('appFillSeries') },
                { value: 'flash-fill', label: t('appFlashFill') },
              ]}
              onPick={(value) => onCommand(value)}
            />
            <MenuSelect
              className="select-like compact"
              label="Clear"
              data-tip={t('appClear')}
              display={
                <>
                  <ToolSymbol symbol="⌫" /> {t('appClear')}
                </>
              }
              options={[
                { value: 'all', label: t('appClearAll') },
                { value: 'formats', label: t('appClearFormats') },
                { value: 'contents', label: t('appClearContents') },
                { value: 'comments', label: t('appClearComments') },
                { value: 'hyperlinks', label: t('appClearHyperlinks') },
              ]}
              onPick={(value) => onCommand(`clear-${value}`)}
            />
          </div>
          <div className="inline-tools">
            <button
              className="labeled"
              data-tip={t('appFindTitle')}
              onClick={() => onCommand('find')}
            >
              <ToolSymbol symbol="🔍" /> {t('appFind')}
            </button>
            <button
              className="labeled"
              data-tip={t('appReplace')}
              onClick={() => onCommand('replace')}
            >
              <ToolSymbol symbol="⇄" /> {t('appReplace')}
            </button>
            <button
              className="labeled"
              data-tip={`${t('appGoTo')} ${platformShortcuts('⌘G')}`}
              onClick={() => onCommand('goto-open')}
            >
              <ToolSymbol symbol="⌖" /> {t('appGoTo')}
            </button>
          </div>
        </div>
      </RibbonGroup>
    </div>
  )
}

function autoSumOptions(t: (key: StringKey) => string): { value: string; label: string }[] {
  return [
    { value: 'SUM', label: t('appFnSum') },
    { value: 'AVERAGE', label: t('appFnAverage') },
    { value: 'COUNT', label: t('appFnCountNumbers') },
    { value: 'MAX', label: t('appFnMax') },
    { value: 'MIN', label: t('appFnMin') },
  ]
}

/// Escape-to-close for the ribbon dropdowns; outside-press / blur / shell
/// chrome-press dismissal lives in the shared useDismissablePopover.
export function useEscapeClose(open: boolean, close: () => void): void {
  const closeRef = useRef(close)
  closeRef.current = close
  useEffect(() => {
    if (!open) return
    const onKey = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') closeRef.current()
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [open])
}

/// Custom ribbon dropdown replacing the native <select>: macOS pops the native
/// menu over the control (covering ribbon content), while this panel is
/// anchored below its trigger — same pattern as the slides ribbon's .rb-drop.
function MenuSelect({
  label,
  'data-tip': tip,
  className,
  cover = false,
  display,
  value = '',
  options,
  onPick,
}: {
  /// aria-label (mirrors the old select's aria-label)
  readonly label: string
  readonly 'data-tip'?: string
  /// trigger classes for the combobox look (e.g. 'select-like compact')
  readonly className?: string
  /// invisible trigger stretched over the host tool (old .cover-select)
  readonly cover?: boolean
  /// content shown in the trigger (combobox variants)
  readonly display?: React.ReactNode
  /// currently applied value, highlighted in the open panel ('' = none)
  readonly value?: string
  readonly options: readonly { value: string; label: string; icon?: React.ReactNode }[]
  readonly onPick: (value: string) => void
}): React.JSX.Element {
  const [open, setOpen] = useState(false)
  const wrapRef = useRef<HTMLDivElement>(null)
  // outside press / window blur / shell chrome press — the shared hook
  useDismissablePopover(open, () => setOpen(false), { inside: () => [wrapRef.current] })
  useEscapeClose(open, () => setOpen(false))
  return (
    <div ref={wrapRef} className={`menu-select${cover ? ' menu-select-cover' : ''}`}>
      <button
        type="button"
        className={cover ? 'cover-select' : className}
        data-tip={tip}
        aria-label={label}
        aria-haspopup="listbox"
        aria-expanded={open}
        onClick={() => setOpen((v) => !v)}
      >
        {!cover && (
          <>
            <span className="menu-select-value">{display}</span>
            <CaretIcon />
          </>
        )}
      </button>
      {open && (
        <div className="menu-select-drop" role="listbox" aria-label={label}>
          {options.map((option) => (
            <button
              type="button"
              key={option.value}
              role="option"
              aria-selected={value !== '' && option.value === value}
              className={value !== '' && option.value === value ? 'on' : ''}
              onClick={() => {
                setOpen(false)
                onPick(option.value)
              }}
            >
              {option.icon && <span className="menu-option-icon">{option.icon}</span>}
              {option.label}
            </button>
          ))}
        </div>
      )}
    </div>
  )
}

/// Cross-app shared shape gallery (slides parity), minus Lines — the grid
/// renderer draws shapes as filled paths and cannot show stroke-only connectors.
const SHEET_SHAPE_GROUPS = SHAPE_GALLERY_GROUPS.filter(
  (g) => g.groupKey !== 'ribbonShapeGroupLines',
)

/// Shapes dropdown: grouped outline-icon grid (same look/content as docs and
/// slides); the trigger is an invisible cover like MenuSelect's `cover` mode.
function ShapeGallerySelect({
  label,
  onPick,
}: {
  readonly label: string
  readonly onPick: (prst: string) => void
}): React.JSX.Element {
  const { t } = useI18n()
  const [open, setOpen] = useState(false)
  const wrapRef = useRef<HTMLDivElement>(null)
  // outside press / window blur / shell chrome press — the shared hook
  useDismissablePopover(open, () => setOpen(false), { inside: () => [wrapRef.current] })
  useEscapeClose(open, () => setOpen(false))
  return (
    <div ref={wrapRef} className="menu-select menu-select-cover">
      <button
        type="button"
        className="cover-select"
        aria-label={label}
        aria-haspopup="listbox"
        aria-expanded={open}
        onClick={() => setOpen((v) => !v)}
      />
      {open && (
        <div className="menu-select-drop rb-shape-gallery" role="listbox" aria-label={label}>
          {SHEET_SHAPE_GROUPS.map((group) => (
            <div key={group.groupKey}>
              <div className="rb-drop-title">{t(group.groupKey as StringKey)}</div>
              <div className="rb-shape-grid">
                {group.shapes.map((s) => (
                  <button
                    type="button"
                    key={s.prst}
                    className="rb-shape-cell"
                    data-tip={t(s.labelKey as StringKey)}
                    aria-label={t(s.labelKey as StringKey)}
                    onClick={() => {
                      setOpen(false)
                      onPick(s.prst)
                    }}
                  >
                    <ShapePreview prst={s.prst} size={18} />
                  </button>
                ))}
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  )
}

/// MenuSelect variant whose trigger is a free-text input (Excel's font
/// name/size boxes): Enter or clicking away commits the typed value, Esc
/// reverts, the caret opens the preset list.
function EditableMenuSelect({
  label,
  'data-tip': tip,
  className,
  value,
  options,
  onOpen,
  onPick,
  commit,
}: {
  readonly label: string
  readonly 'data-tip'?: string
  readonly className?: string
  readonly value: string
  readonly options: readonly { value: string; label: string; sep?: boolean }[]
  /// fired on the click that opens the list (lazy option loading)
  readonly onOpen?: () => void
  readonly onPick: (value: string) => void
  /// apply typed text (caller validates; invalid input is dropped silently)
  readonly commit: (text: string) => void
}): React.JSX.Element {
  const [open, setOpen] = useState(false)
  const [draft, setDraftState] = useState<string | null>(null)
  // mirrors draft synchronously: Enter/Esc call blur(), whose handler runs
  // before the setState above lands
  const draftRef = useRef<string | null>(null)
  const setDraft = (next: string | null): void => {
    draftRef.current = next
    setDraftState(next)
  }
  const wrapRef = useRef<HTMLDivElement>(null)
  // outside press / window blur / shell chrome press — the shared hook
  useDismissablePopover(open, () => setOpen(false), { inside: () => [wrapRef.current] })
  useEscapeClose(open, () => setOpen(false))
  const commitDraft = (): void => {
    const text = draftRef.current?.trim()
    if (text && text !== value) commit(text)
    setDraft(null)
  }
  return (
    <div ref={wrapRef} className="menu-select">
      <span className={`${className ?? ''} menu-select-edit`} data-tip={tip}>
        <input
          value={draft ?? value}
          aria-label={label}
          onFocus={(event) => event.currentTarget.select()}
          onChange={(event) => setDraft(event.currentTarget.value)}
          onKeyDown={(event) => {
            if (event.key === 'Enter') {
              event.currentTarget.blur() // blur handler commits
            } else if (event.key === 'Escape') {
              setDraft(null)
              event.currentTarget.blur()
            }
          }}
          onBlur={(event) => {
            // focus moved into the dropdown → let the option click win
            if (wrapRef.current?.contains(event.relatedTarget as Node)) return
            commitDraft()
          }}
        />
        <button
          type="button"
          className="menu-select-caret"
          aria-label={label}
          aria-haspopup="listbox"
          aria-expanded={open}
          onClick={() => {
            if (!open) onOpen?.()
            setOpen(!open)
          }}
        >
          <CaretIcon />
        </button>
      </span>
      {open && (
        <div className="menu-select-drop" role="listbox" aria-label={label}>
          {options.map((option) => (
            <button
              type="button"
              key={option.value}
              role="option"
              aria-selected={option.value === value}
              className={`${option.value === value ? 'on' : ''}${option.sep ? ' sep-above' : ''}`}
              onClick={() => {
                setOpen(false)
                setDraft(null)
                onPick(option.value)
              }}
            >
              {option.label}
            </button>
          ))}
        </div>
      )}
    </div>
  )
}

function NumberFormatSelect({
  pattern,
  onCommand,
}: {
  readonly pattern: string
  readonly onCommand: (command: string) => void
}): React.JSX.Element {
  const { t } = useI18n()
  const current = categoryOptionForPattern(pattern)
  const localized = (label: string): string => {
    const labelKey = NUMBER_FORMAT_LABEL[label]
    return labelKey ? t(labelKey) : label
  }
  return (
    <MenuSelect
      className="select-like"
      label="Number format"
      data-tip={pattern || t('dlgFcNumGeneral')}
      value={current}
      display={localized(current)}
      options={numberFormatCategories().map((category) => ({
        value: category.label,
        label: localized(category.label),
      }))}
      onPick={(value) => {
        const category = numberFormatCategories().find((candidate) => candidate.label === value)
        if (category) onCommand(`format:${category.pattern}`)
      }}
    />
  )
}

function RibbonGroup({
  label,
  children,
}: {
  readonly label: string
  readonly children: React.ReactNode
}): React.JSX.Element {
  // No visible group captions — the label stays for assistive tech.
  return (
    <section className="ribbon-group" aria-label={label}>
      <div className="ribbon-group-content">{children}</div>
    </section>
  )
}

function RibbonButton({
  label,
  detail,
  symbol,
  accent = false,
  menu = false,
  compact = false,
  large = false,
  active = false,
  disabled = false,
  onClick,
}: {
  readonly label: string
  readonly detail: string
  readonly symbol: string
  readonly accent?: boolean
  readonly menu?: boolean
  /** Icon-only in narrow windows (e.g. Cut/Copy). */
  readonly compact?: boolean
  /** Icon over label, for the tall Insert-tab buttons. */
  readonly large?: boolean
  /** Pressed-state echo (e.g. highlighting the current chart type). */
  readonly active?: boolean
  readonly disabled?: boolean
  readonly onClick: () => void
}): React.JSX.Element {
  return (
    <button
      className={`ribbon-tool as-button ${accent ? 'accent' : ''} ${compact ? 'compact-icon' : ''} ${large ? 'large' : ''} ${active ? 'active' : ''}`}
      onClick={onClick}
      disabled={disabled}
      data-tip={label}
      data-tip-detail={detail}
    >
      {large ? (
        // The dropdown caret sits beside the icon (top row, inside the
        // hover plate), not after the label
        <span className="tool-icon-row">
          <ToolSymbol symbol={symbol} />
          {menu && <CaretIcon />}
        </span>
      ) : (
        <ToolSymbol symbol={symbol} />
      )}
      <span>
        <strong>
          {label}
          {!large && menu && <CaretIcon />}
        </strong>
        <small>{detail}</small>
      </span>
    </button>
  )
}
