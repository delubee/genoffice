import {
  SheetChartContent,
  sheetChartAreaStyle,
  seriesColor,
  pieSliceColor,
  truncateLabel,
  type ChartMetadata,
  type ChartElementRef,
} from './SheetChart'
export * from './SheetChart'
import { useEffect, useId, useRef, useState, useSyncExternalStore } from 'react'
import { createPortal } from 'react-dom'
import { BooleanNumber } from '@univerjs/core'
import { Dropdown, shapePreviewPath, useDismissablePopover } from '@genoffice/ui/chrome'

import type { createUniver } from './create-univer'

import {
  applyChartStateEdit,
  chartSupportsDataLabels,
  formatCategoryLabel,
  isComboChartTypes,
  scatterAxisBounds,
  splitSheetRef,
} from '@genoffice/xlsx-gateway/domain/chart-visual'
import { parseAddress } from '@genoffice/xlsx-gateway/domain/cell-address'
import { ColorDropdown } from './ColorDropdown'
import { t } from './i18n/locale'
import { oleCaption, oleFrameStyle, oleRenderKind } from './ole-visual'
import { VisualDeleteButton } from './VisualDeleteButton'
import { shouldShowVisualDeleteButton } from './visual-delete-button'
import type { WorkbookChartEdit, WorkbookFile, WorkbookVisualObject } from '../shared/desktop-api'
import { shapeRunFontSize, shapeTextOverflowClass, shapeTextScaleStyle } from './shape-text-scale'
import {
  shapeBulletWidth,
  shapeParagraphIndentStyle,
  shapeParagraphMarkers,
  type ShapeParagraph,
} from './shape-text-bullets'
import { useWorkbookMediaUrl, type WorkbookMediaReader } from './workbook-media'

type UniverRuntime = ReturnType<typeof createUniver>
type ActiveWorkbook = NonNullable<ReturnType<UniverRuntime['univerAPI']['getActiveWorkbook']>>
type UniverWorksheet = NonNullable<ReturnType<ActiveWorkbook['getSheetBySheetId']>>

interface Disposable {
  dispose(): void
}

/// What installWorkbookVisuals actually needs from a workbook: the demo
/// workbook satisfies this with its snapshot visuals and a stable fake
/// sessionId; lazy mode passes the full WorkbookFile.
export interface VisualHost {
  readonly readMedia?: WorkbookMediaReader
  readonly sessionId: WorkbookFile['sessionId']
  readonly visuals: readonly WorkbookVisualObject[]
}

export type ChartEditData = Omit<WorkbookChartEdit, 'chartPath'>

export interface ChartVectorRead {
  readonly vector: readonly (string | number | boolean | null | undefined)[]
  readonly ref: string
}

/// Chart edits key on the chart part path for file charts and on the visual
/// id for charts that only exist in memory (session adds, demo workbook);
/// the app-side handler routes on the key's shape.
export interface ChartEditing {
  readonly edits: ReadonlyMap<string, ChartEditData>
  readonly onEdit: (editKey: string, edit: ChartEditData) => void | Promise<void>
  readonly canEdit?: (editKey: string) => boolean
  readonly contextMenu?: boolean
  readonly dialogs?: ReadonlySet<ChartDialogKind>
  readonly visualCommands?: ReadonlySet<string>
  readonly onError?: (error: unknown) => void
  /// Reads a single-row/column range so series data-range edits can bake
  /// current sheet values into the chart cache. Throws on invalid ranges.
  readonly readVector?: (editKey: string, range: string) => Promise<ChartVectorRead>
}

export interface ShapeEditChanges {
  readonly anchor?: WorkbookVisualObject['anchor']
  /// New a:xfrm ext in EMU — rides along when resizing a rotated shape,
  /// whose anchor stores the rotated AABB rather than the true frame.
  readonly frameSize?: { readonly width: number; readonly height: number }
  readonly text?: string
  /// `#rrggbb` or "none"; rewrites the shape's own spPr paint.
  readonly fillColor?: string
  readonly lineColor?: string
  readonly remove?: true
  readonly rotation?: number
  readonly flipH?: boolean
  readonly flipV?: boolean
  readonly altText?: string
  readonly editAs?: 'twoCell' | 'oneCell' | 'absolute'
  /// Empty string removes the link.
  readonly hyperlink?: string
}

/// Right-click / Arrange commands that need app state (dialogs, clipboard,
/// z-order across visuals) bubble to the app through one listener, keyed
/// by the same `visual:` command vocabulary the ribbon uses.
let visualCommandListener: ((command: string, visualId: string) => void) | null = null

export function setVisualCommandListener(
  listener: ((command: string, visualId: string) => void) | null,
): void {
  visualCommandListener = listener
}

let openShapeEditors = 0
export function isShapeTextEditorOpen() {
  return openShapeEditors > 0
}
export interface ShapeEditing {
  readonly onTextEdit?: ((visualId: string, changes: ShapeEditChanges) => Promise<void>) | undefined
  readonly onEdit: (visualId: string, changes: ShapeEditChanges) => void
}

/// Visuals created this session are fully editable in place: their whole
/// object lives in the journal, so anchor/text changes save for free.
export function isEditableShape(visual: WorkbookVisualObject): boolean {
  return visual.kind === 'shape' && visual.id.startsWith('added-shape-')
}

/// Pictures pasted or inserted this session move/resize/delete the same way
/// (no text editing).
export function isEditableAddedImage(visual: WorkbookVisualObject): boolean {
  return visual.kind === 'image' && visual.id.startsWith('added-image-')
}

/// Images and shapes already in the file move/delete through a surgical
/// anchor edit, located by the sidecar's (drawingPath, drawingIndex) pair.
/// Charts keep their own editor; deletion routes through the same edit.
export function isEditableFileVisual(visual: WorkbookVisualObject): boolean {
  return (
    (visual.kind === 'image' || visual.kind === 'shape') &&
    !visual.id.startsWith('added-') &&
    visual.drawingPath !== undefined &&
    visual.drawingIndex !== undefined
  )
}

/// Charts move through the same anchor edit as images/shapes and delete
/// through a cascading one (the save also drops the chart part, rel, and
/// content-type override); session-added and demo charts live in memory.
/// Inline style of one run; the bullet marker borrows its paragraph's first
/// run without the underline. Caps are display-only so the stored text keeps
/// its casing on save.
function shapeRunStyle(run: ShapeParagraph['runs'][number] | undefined, withDecoration: boolean) {
  if (!run) return {}
  return {
    ...(run.color ? { color: run.color } : {}),
    ...(run.bold ? { fontWeight: 700 } : {}),
    ...(run.italic ? { fontStyle: 'italic' } : {}),
    ...(withDecoration && run.underline ? { textDecoration: 'underline' } : {}),
    ...(run.size ? { fontSize: shapeRunFontSize(run.size) } : {}),
    ...(withDecoration && run.caps === 'all' ? { textTransform: 'uppercase' as const } : {}),
    ...(withDecoration && run.caps === 'small' ? { fontVariantCaps: 'small-caps' as const } : {}),
  }
}

function isEditableChart(visual: WorkbookVisualObject): boolean {
  if (visual.kind !== 'chart') return false
  if (visual.id.startsWith('added-') || visual.id.startsWith('demo-')) return true
  return visual.drawingPath !== undefined && visual.drawingIndex !== undefined
}

/// The frame each installed visual was given, in unzoomed sheet px: what the
/// print layout positions its snapshot by (`data-print-visual` marks the node).
export interface InstalledVisualFrame {
  readonly visual: WorkbookVisualObject
  readonly fromRow: number
  readonly fromColumn: number
  readonly toRow: number
  readonly toColumn: number
  readonly marginX: number
  readonly marginY: number
  readonly width: number
  readonly height: number
}

const installedFrames = new Map<string, InstalledVisualFrame[]>()

export function installedVisualFrames(sheetId: string): readonly InstalledVisualFrame[] {
  return installedFrames.get(sheetId) ?? []
}

export function installWorkbookVisuals(
  runtime: UniverRuntime,
  file: VisualHost,
  sheetId: string,
  chartEditing?: ChartEditing,
  shapeEditing?: ShapeEditing,
): Disposable[] {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  if (!workbook) return []
  const disposables: Disposable[] = []
  const frames: InstalledVisualFrame[] = []
  installedFrames.set(sheetId, frames)
  for (const visual of file.visuals.filter((candidate) => candidate.sheetId === sheetId)) {
    const worksheet = workbook.getSheetBySheetId(visual.sheetId)
    if (!worksheet) continue
    const componentKey = `xlsx-${file.sessionId}-${visual.id}`
    const editable =
      isEditableShape(visual) ||
      isEditableAddedImage(visual) ||
      isEditableFileVisual(visual) ||
      isEditableChart(visual)
    // Lazy grids are sized to the data, but session-added visuals anchor
    // beyond it (default: two columns right of the data) — grow the grid so
    // the float keeps its frame instead of being clamped into a sliver.
    if (visual.anchor.toRow >= worksheet.getMaxRows()) {
      worksheet.setRowCount(visual.anchor.toRow + 1)
    }
    if (visual.anchor.toColumn >= worksheet.getMaxColumns()) {
      worksheet.setColumnCount(visual.anchor.toColumn + 1)
    }
    // An out-of-bounds anchor would throw in getRange and take every other
    // visual down with it — clamp the display range to the sheet.
    const maxRows = worksheet.getMaxRows()
    const maxColumns = worksheet.getMaxColumns()
    let fromRow = Math.min(visual.anchor.fromRow, maxRows - 1)
    let fromColumn = Math.min(visual.anchor.fromColumn, maxColumns - 1)
    let toRow = Math.min(visual.anchor.toRow, maxRows - 1)
    let toColumn = Math.min(visual.anchor.toColumn, maxColumns - 1)
    // Pixel-exact frame (Excel twoCellAnchor semantics): each marker is a
    // cell plus an EMU offset inside it, so the box runs from the `from`
    // marker to the `to` marker — not across the whole cell range.
    const columnWidth = (index: number): number => Math.max(worksheet.getColumnWidth(index), 1)
    const rowHeight = (index: number): number => Math.max(worksheet.getRowHeight(index), 1)
    let marginX = Math.max(0, visual.anchor.fromColumnOffset / EMU_PER_PIXEL)
    let marginY = Math.max(0, visual.anchor.fromRowOffset / EMU_PER_PIXEL)
    const explicitTo = visual.anchor.explicitTo === true
    let width = markerSpan(
      { index: fromColumn, offset: marginX },
      clampExplicitTo(markerFrom(toColumn, visual.anchor.toColumnOffset), explicitTo, columnWidth),
      columnWidth,
    )
    let height = markerSpan(
      { index: fromRow, offset: marginY },
      clampExplicitTo(markerFrom(toRow, visual.anchor.toRowOffset), explicitTo, rowHeight),
      rowHeight,
    )
    // A rotated shape's anchor holds its rotated bounds while xfrm ext keeps
    // the true frame; both share the same center. The float container clips
    // (Univer overflow-hidden), so re-anchor it to the AABB of the rotated
    // true frame — the component centers the ext box inside and rotates it.
    if (visual.rotation && visual.frameWidth && visual.frameHeight && width > 0 && height > 0) {
      const extWidth = visual.frameWidth / EMU_PER_PIXEL
      const extHeight = visual.frameHeight / EMU_PER_PIXEL
      const radians = (visual.rotation * Math.PI) / 180
      const aabbWidth =
        Math.abs(extWidth * Math.cos(radians)) + Math.abs(extHeight * Math.sin(radians))
      const aabbHeight =
        Math.abs(extWidth * Math.sin(radians)) + Math.abs(extHeight * Math.cos(radians))
      const fromX = walkMarker(
        { index: fromColumn, offset: marginX },
        (width - aabbWidth) / 2,
        columnWidth,
        maxColumns - 1,
      )
      const fromY = walkMarker(
        { index: fromRow, offset: marginY },
        (height - aabbHeight) / 2,
        rowHeight,
        maxRows - 1,
      )
      fromColumn = fromX.index
      fromRow = fromY.index
      marginX = fromX.offset
      marginY = fromY.offset
      toColumn = walkMarker(fromX, aabbWidth, columnWidth, maxColumns - 1).index
      toRow = walkMarker(fromY, aabbHeight, rowHeight, maxRows - 1).index
      width = aabbWidth
      height = aabbHeight
    }
    const range = worksheet.getRange(
      fromRow,
      fromColumn,
      Math.max(1, toRow + 1 - fromRow),
      Math.max(1, toColumn + 1 - fromColumn),
    )
    // Degenerate anchors (a zero marker span on BOTH axes): keep the legacy
    // behavior of filling the clamped cell range. Any real span stays a
    // marker frame with a 1px floor — tiny icons and hairlines (a line's
    // thin axis spans 0) must not balloon to their cell range, which turned
    // zero-height rules into cell-tall diagonals and icon parts into
    // cell-sized blobs.
    const framed = width > 0 || height > 0
    const frameWidth = Math.max(width, 1)
    const frameHeight = Math.max(height, 1)
    if (framed) {
      frames.push({
        visual,
        fromRow,
        fromColumn,
        toRow,
        toColumn,
        marginX,
        marginY,
        width: frameWidth,
        height: frameHeight,
      })
    }
    // RTL sheets mirror the float too (Excel keeps logical anchors; the box
    // lands mirrored). Univer positions the DOM from the anchor cell's
    // mirrored left edge, so restate the margin as "mirrored box left minus
    // that edge": colWidth(from) - marginX - width.
    const rtl = worksheet.getSheet().getConfig().rightToLeft === BooleanNumber.TRUE
    const anchoredMarginX = rtl ? columnWidth(fromColumn) - marginX - frameWidth : marginX
    const layout = framed
      ? floatDomLayout({
          width: frameWidth,
          height: frameHeight,
          marginX: anchoredMarginX,
          marginY,
        })
      : {}
    const frame = framed ? { width: frameWidth, height: frameHeight } : undefined
    const component =
      shapeEditing && editable
        ? () => (
            <div className="xlsx-print-visual" data-print-visual={visual.id}>
              <EditableShapeVisual
                file={file}
                visual={visual}
                worksheet={worksheet}
                allowText={isEditableShape(visual)}
                chartEditing={chartEditing}
                onEdit={shapeEditing.onEdit}
                onTextEdit={shapeEditing.onTextEdit}
                frame={frame}
              />
            </div>
          )
        : () => (
            <div className="xlsx-print-visual" data-print-visual={visual.id}>
              <WorkbookVisual
                file={file}
                visual={visual}
                chartEditing={chartEditing}
                frame={frame}
              />
            </div>
          )
    disposables.push(runtime.univerAPI.registerComponent(componentKey, component))
    const floating = worksheet.addFloatDomToRange(
      range,
      {
        componentKey,
        allowTransform: false,
        eventPassThrough: false,
      },
      layout,
      `${file.sessionId}-${visual.id}`,
    )
    if (floating) disposables.push(floating)
  }
  return disposables
}

export interface SparklineGroupState {
  readonly type: 'line' | 'column' | 'stacked'
  readonly color?: string | undefined
  readonly negativeColor?: string | undefined
  readonly cells: readonly { readonly cell: string; readonly sourceRef: string }[]
}

const MAX_SPARKLINES_PER_SHEET = 200

/// In-cell minicharts: one pass-through float DOM per sparkline cell, values
/// read live from the grid (journal edits land there before the save).
export function installSparklines(
  runtime: UniverRuntime,
  groups: readonly SparklineGroupState[],
  sheetId: string,
): Disposable[] {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getSheetBySheetId(sheetId)
  if (!workbook || !worksheet) return []
  const sheetIdByName = new Map(
    workbook.getSheets().map((sheet) => [sheet.getSheetName(), sheet.getSheetId()] as const),
  )
  const disposables: Disposable[] = []
  let budget = MAX_SPARKLINES_PER_SHEET
  for (const [groupIndex, group] of groups.entries()) {
    for (const [cellIndex, entry] of group.cells.entries()) {
      if (budget <= 0) return disposables
      budget -= 1
      const split = splitSheetRef(entry.sourceRef)
      const source = split
        ? workbook.getSheetBySheetId(sheetIdByName.get(split.sheetName) ?? '')
        : worksheet
      if (!source) continue
      let values: number[]
      let host: { row: number; column: number }
      try {
        host = parseAddress(entry.cell.replace(/\$/g, ''))
        const grid = source
          .getRange((split?.range ?? entry.sourceRef).replace(/\$/g, ''))
          .getValues() as (string | number | boolean | null | undefined)[][]
        values = grid.flat().map((value) => {
          const numeric = typeof value === 'number' ? value : Number(String(value ?? '').trim())
          return Number.isFinite(numeric) ? numeric : 0
        })
      } catch {
        continue
      }
      if (values.length === 0) continue
      const componentKey = `sparkline-${sheetId}-${groupIndex}-${cellIndex}`
      disposables.push(
        runtime.univerAPI.registerComponent(componentKey, () => (
          <Sparkline
            values={values}
            type={group.type}
            color={group.color}
            negativeColor={group.negativeColor}
          />
        )),
      )
      const range = worksheet.getRange(host.row, host.column, 1, 1)
      const floating = worksheet.addFloatDomToRange(
        range,
        { componentKey, allowTransform: false, eventPassThrough: true },
        {},
        componentKey,
      )
      if (floating) disposables.push(floating)
    }
  }
  return disposables
}

export interface CellImageState {
  readonly id: string
  readonly row: number
  readonly column: number
  /// Session-embedded bytes (a data URL); file pictures resolve through the
  /// media IPC by id instead.
  readonly dataUrl?: string | undefined
}

/// In-cell rich-value pictures ("place picture in cell"): one pass-through
/// float DOM per host cell, bytes fetched through the same media IPC as
/// floating images.
export function installCellImages(
  runtime: UniverRuntime,
  sessionId: string,
  images: readonly CellImageState[],
  sheetId: string,
): Disposable[] {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getSheetBySheetId(sheetId)
  if (!workbook || !worksheet) return []
  const disposables: Disposable[] = []
  for (const image of images) {
    if (image.row >= worksheet.getMaxRows() || image.column >= worksheet.getMaxColumns()) continue
    const componentKey = `cell-image-${sessionId}-${sheetId}-${image.id}`
    disposables.push(
      runtime.univerAPI.registerComponent(componentKey, () => (
        <CellImage sessionId={sessionId} imageId={image.id} dataUrl={image.dataUrl} />
      )),
    )
    const range =
      worksheet.getCellMergeData(image.row, image.column) ??
      worksheet.getRange(image.row, image.column, 1, 1)
    const floating = worksheet.addFloatDomToRange(
      range,
      { componentKey, allowTransform: false, eventPassThrough: true },
      {},
      componentKey,
    )
    if (floating) disposables.push(floating)
  }
  return disposables
}

/// Excel fits the picture inside the cell box, preserving its aspect ratio;
/// while the bytes load the cell just stays blank.
function CellImage({
  sessionId,
  imageId,
  dataUrl,
}: {
  readonly sessionId: string
  readonly imageId: string
  readonly dataUrl?: string | undefined
}): React.JSX.Element | null {
  // Session embeds carry their bytes inline; only file pictures round-trip
  // through the media IPC.
  const { url, unavailable } = useWorkbookMediaUrl(
    dataUrl === undefined ? sessionId : undefined,
    imageId,
  )
  const source = dataUrl ?? url
  if (source === null || (dataUrl === undefined && unavailable)) return null
  return (
    <div className="xlsx-cell-image">
      <img src={source} alt="" />
    </div>
  )
}

function Sparkline({
  values,
  type,
  color,
  negativeColor,
}: {
  readonly values: readonly number[]
  readonly type: SparklineGroupState['type']
  readonly color?: string | undefined
  readonly negativeColor?: string | undefined
}): React.JSX.Element {
  const stroke = color ?? '#376092'
  const negative = negativeColor ?? '#d00000'
  const width = 100
  const height = 26
  if (type === 'line') {
    const minimum = Math.min(...values)
    const maximum = Math.max(...values)
    const span = maximum - minimum || 1
    const count = Math.max(1, values.length - 1)
    const points = values
      .map(
        (value, index) =>
          `${2 + (index / count) * (width - 4)},${height - 3 - ((value - minimum) / span) * (height - 6)}`,
      )
      .join(' ')
    return (
      <svg className="sparkline-svg" viewBox={`0 0 ${width} ${height}`} preserveAspectRatio="none">
        <polyline points={points} fill="none" stroke={stroke} strokeWidth="1.6" />
      </svg>
    )
  }
  // column / stacked (win-loss): bars around a zero baseline.
  const isWinLoss = type === 'stacked'
  const maximum = Math.max(...values.map((value) => Math.abs(value)), 0) || 1
  const barWidth = Math.max(1.5, (width - 4) / values.length - 1)
  const zeroY = height / 2
  return (
    <svg className="sparkline-svg" viewBox={`0 0 ${width} ${height}`} preserveAspectRatio="none">
      {values.map((value, index) => {
        const magnitude = isWinLoss
          ? value === 0
            ? 0
            : zeroY - 3
          : (Math.abs(value) / maximum) * (zeroY - 2)
        return (
          <rect
            key={index}
            x={2 + index * ((width - 4) / values.length)}
            y={value >= 0 ? zeroY - magnitude : zeroY}
            width={barWidth}
            height={Math.max(magnitude, value === 0 ? 0.5 : 1)}
            fill={value < 0 ? negative : stroke}
          />
        )
      })}
    </svg>
  )
}

/// Anchor-frame pixel size at install time — the shape renderer needs the
/// aspect ratio for preset geometry and Excel's rotated-anchor un-swap.
export interface ShapeFrame {
  readonly width: number
  readonly height: number
}

function WorkbookVisual({
  file,
  visual,
  chartEditing,
  frame,
}: {
  readonly file: VisualHost
  readonly visual: WorkbookVisualObject
  readonly chartEditing?: ChartEditing | undefined
  readonly frame?: ShapeFrame | undefined
}): React.JSX.Element {
  if (visual.kind === 'chart' && visual.chart) {
    return (
      <ChartVisual
        chart={visual.chart}
        visualId={visual.id}
        chartPath={visual.chartPath}
        chartEditing={chartEditing}
        frame={frame}
      />
    )
  }
  if (visual.kind === 'image') {
    return <ImageVisual file={file} visual={visual} frame={frame} />
  }
  if (visual.kind === 'ole') {
    return <OleVisual file={file} visual={visual} />
  }
  if (visual.kind === 'slicer') {
    return <SlicerVisual visual={visual} />
  }
  return <ShapeVisual file={file} visual={visual} frame={frame} />
}

/// A drag commit tears down and re-registers every float DOM, dropping the
/// browser focus that doubles as the selection state — this hands it back
/// to the moved visual when its replacement mounts.
let pendingFocusId: string | null = null

/// Reinstalls must not run mid-drag: disposing the dragged node kills its
/// pointer capture and the drag dies silently. Install queues poll this and
/// wait for the drop.
let activeDragCount = 0

export function isVisualDragActive(): boolean {
  return activeDragCount > 0
}

/// Reinstalls also wait while an inline chart editor is open — disposing its
/// float DOM would silently discard everything typed into it.
let activeChartApplyCount = 0
export function isChartApplyActive(): boolean {
  return activeChartApplyCount > 0
}
function trackChartApply(): () => void {
  activeChartApplyCount++
  let released = false
  return () => {
    if (!released) {
      released = true
      activeChartApplyCount--
    }
  }
}

let openEditorCount = 0

export function isChartEditorOpen(): boolean {
  return openEditorCount > 0
}

export function trackChartEditorOpen(): () => void {
  openEditorCount += 1
  return () => {
    openEditorCount -= 1
  }
}

/// A move commit reinstalls the float ~100ms later; removing the body-level
/// ghost on pointerup would flash the visual back at its old anchor for that
/// window. The ghost is parked here and removed once the replacement mounts;
/// the timer covers rejected edits, where no reinstall ever comes.
let dropGhost: { id: string; ghost: HTMLElement; host: HTMLElement; timer: number } | null = null

function discardDropGhost(): void {
  if (!dropGhost) return
  const { ghost, host, timer } = dropGhost
  dropGhost = null
  clearTimeout(timer)
  ghost.remove()
  host.style.visibility = ''
}

function parkDropGhost(id: string, ghost: HTMLElement, host: HTMLElement): void {
  discardDropGhost()
  dropGhost = { id, ghost, host, timer: window.setTimeout(discardDropGhost, 1200) }
}

function claimDropGhost(id: string, node: HTMLElement): void {
  if (dropGhost?.id !== id) return
  const { ghost, timer } = dropGhost
  dropGhost = null
  clearTimeout(timer)
  // The float container may attach/lay out late (see tryFocus below): keep
  // the preview until the replacement actually paints at the new anchor.
  const settle = (attempts: number): void => {
    if (attempts > 0 && (!node.isConnected || node.getBoundingClientRect().width === 0)) {
      setTimeout(() => settle(attempts - 1), 80)
      return
    }
    requestAnimationFrame(() => requestAnimationFrame(() => ghost.remove()))
  }
  settle(12)
}

export interface VisualSelectionListener {
  readonly select: (visual: WorkbookVisualObject) => void
  readonly deselect: () => void
}

/// Selection lives here, not in DOM focus: reinstalls and Univer's own DOM
/// churn drop focus, but the id survives and the remounted wrapper re-renders
/// selected. A module singleton avoids threading callbacks through every
/// install call site; the app mirrors it for the contextual ribbon tab.
let selectedVisualId: string | null = null
/// Shift-click extends the selection; the primary (last clicked) stays in
/// selectedVisualId, the others ride here for Align / Distribute.
let extraSelectedIds: string[] = []
const selectionSubscribers = new Set<() => void>()
let selectionListener: VisualSelectionListener | null = null

function selectVisual(visual: WorkbookVisualObject | null, extend = false): void {
  const nextId = visual?.id ?? null
  const previousExtras = extraSelectedIds
  if (extend && nextId !== null && selectedVisualId !== null && selectedVisualId !== nextId) {
    extraSelectedIds = [
      ...extraSelectedIds.filter((id) => id !== nextId && id !== selectedVisualId),
      selectedVisualId,
    ]
  } else if (!extend || nextId === null) {
    extraSelectedIds = []
  }
  if (nextId !== selectedVisualId || previousExtras !== extraSelectedIds) {
    selectedVisualId = nextId
    for (const notify of selectionSubscribers) notify()
    // Element selection never outlives its chart's selection.
    if (selectedChartElement && selectedChartElement.visualId !== nextId) {
      selectedChartElement = null
      for (const notify of elementSubscribers) notify()
    }
  }
  if (visual) selectionListener?.select(visual)
  else selectionListener?.deselect()
}

/// Every selected visual id, primary last.
export function selectedVisualIds(): readonly string[] {
  return selectedVisualId === null ? [] : [...extraSelectedIds, selectedVisualId]
}

export function subscribeVisualSelection(notify: () => void): () => void {
  selectionSubscribers.add(notify)
  return () => selectionSubscribers.delete(notify)
}

/// Clears the selection; with `onlyId`, only when that visual is selected
/// (deleting one visual must not deselect another).
export function clearVisualSelection(onlyId?: string): void {
  if (onlyId !== undefined && onlyId !== selectedVisualId) {
    if (extraSelectedIds.includes(onlyId)) {
      extraSelectedIds = extraSelectedIds.filter((id) => id !== onlyId)
      for (const notify of selectionSubscribers) notify()
    }
    return
  }
  selectVisual(null)
}

/// Selection of an element INSIDE a chart (title, legend, one series, one
/// point/slice, an axis); the format pane reads it to scope its controls.
/// Module store for the same reason as the visual selection.
let selectedChartElement: { visualId: string; element: ChartElementRef } | null = null
const elementSubscribers = new Set<() => void>()

export function setChartElementSelection(visualId: string, element: ChartElementRef | null): void {
  selectedChartElement = element === null ? null : { visualId, element }
  for (const notify of elementSubscribers) notify()
}

export function getChartElementSelection(): { visualId: string; element: ChartElementRef } | null {
  return selectedChartElement
}

export function subscribeChartElementSelection(notify: () => void): () => void {
  elementSubscribers.add(notify)
  return () => elementSubscribers.delete(notify)
}

/// Context-menu / dialog requests bubble to the app through one listener
/// (same recipe as the visual-selection listener).
export type ChartDialogKind = 'select-data' | 'format'
let chartDialogListener: ((editKey: string, dialog: ChartDialogKind) => void) | null = null

export function setChartDialogListener(
  listener: ((editKey: string, dialog: ChartDialogKind) => void) | null,
): void {
  chartDialogListener = listener
}

export function setVisualSelectionListener(listener: VisualSelectionListener | null): void {
  selectionListener = listener
}

function useIsSelected(visualId: string): boolean {
  return useSyncExternalStore(
    (notify) => {
      selectionSubscribers.add(notify)
      return () => selectionSubscribers.delete(notify)
    },
    () => selectedVisualId === visualId || extraSelectedIds.includes(visualId),
  )
}

const RESIZE_CORNERS = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'] as const
export type ResizeCorner = (typeof RESIZE_CORNERS)[number]

/// RTL mirrors the x-axis: a screen-space handle grabs the opposite logical
/// edge (pair with negating the screen dx).
export const mirrorCornerX = (corner: ResizeCorner): ResizeCorner =>
  (({ nw: 'ne', n: 'n', ne: 'nw', e: 'w', se: 'sw', s: 's', sw: 'se', w: 'e' }) as const)[corner]

const cornerEast = (corner: ResizeCorner): boolean =>
  corner === 'ne' || corner === 'e' || corner === 'se'
const cornerWest = (corner: ResizeCorner): boolean =>
  corner === 'nw' || corner === 'w' || corner === 'sw'
const cornerNorth = (corner: ResizeCorner): boolean =>
  corner === 'nw' || corner === 'n' || corner === 'ne'
const cornerSouth = (corner: ResizeCorner): boolean =>
  corner === 'sw' || corner === 's' || corner === 'se'

/// xlsx drawing offsets are EMU; 9525 EMU per CSS pixel at 96dpi.
export const EMU_PER_PIXEL = 9525

export interface FloatDomLayout {
  readonly width: number
  readonly height: number
  readonly marginX: number
  readonly marginY: number
}

/// Univer's float DOM shaves 2px off every side of the rect it is given
/// (wrapper = rect - 2, inner box = rect - 4 hugging the far corner), so a
/// frame-sized rect draws its content 4px short and `object-fit: contain`
/// shrinks small pictures on both axes. Grow the rect by that inset so the
/// inner box lands exactly on the anchor frame. The inset is CSS px while
/// the layout is sheet px: exact at 100% zoom, within 4px elsewhere.
const FLOAT_DOM_INSET = 2

export function floatDomLayout(frame: FloatDomLayout): FloatDomLayout {
  return {
    width: frame.width + 2 * FLOAT_DOM_INSET,
    height: frame.height + 2 * FLOAT_DOM_INSET,
    marginX: frame.marginX - FLOAT_DOM_INSET,
    marginY: frame.marginY - FLOAT_DOM_INSET,
  }
}
/// Frames smaller than this collapse resize handles into each other.
const MIN_FRAME_PIXELS = 24
/// xlsx sheet limits: drags may leave the data-sized lazy grid (install
/// grows it to the anchor), but never the real spreadsheet bounds.
const XLSX_MAX_COLUMN = 16383
const XLSX_MAX_ROW = 1048575

/// One edge of a drawing anchor: a cell index plus a pixel offset inside
/// that cell (the px equivalent of xlsx's `<xdr:col>` + `<xdr:colOff>`).
export interface AnchorMarker {
  index: number
  offset: number
}

export const markerFrom = (index: number, offsetEmu: number): AnchorMarker => ({
  index,
  offset: offsetEmu / EMU_PER_PIXEL,
})

/// Move a marker by a pixel delta, carrying across real row/column sizes.
/// Clamps at the sheet start and inside the last row/column.
export function walkMarker(
  marker: AnchorMarker,
  delta: number,
  sizeOf: (index: number) => number,
  maxIndex: number,
): AnchorMarker {
  let index = Math.min(marker.index, maxIndex)
  let offset = marker.offset + delta
  while (offset < 0 && index > 0) {
    index -= 1
    offset += sizeOf(index)
  }
  if (offset < 0) offset = 0
  while (index < maxIndex && offset >= sizeOf(index)) {
    offset -= sizeOf(index)
    index += 1
  }
  if (index >= maxIndex) offset = Math.min(offset, sizeOf(maxIndex))
  return { index, offset }
}

/// A real `<xdr:to>` marker never reaches past its own cell: Excel clamps
/// the offset at the cell edge (broken writers store the picture's full
/// size there, e.g. ClosedXML). Synthesized to markers (oneCellAnchor ext,
/// absoluteAnchor, group children) encode sizes as offsets past the edge
/// and must keep walking, so only explicit-to anchors clamp.
export function clampExplicitTo(
  to: AnchorMarker,
  explicitTo: boolean,
  sizeOf: (index: number) => number,
): AnchorMarker {
  if (!explicitTo || to.offset <= sizeOf(to.index)) return to
  return { index: to.index, offset: sizeOf(to.index) }
}

/// Pixel distance between two markers (negative when `to` sits before `from`).
function markerSpan(
  from: AnchorMarker,
  to: AnchorMarker,
  sizeOf: (index: number) => number,
): number {
  const span = to.offset - from.offset
  const low = Math.min(from.index, to.index)
  const high = Math.max(from.index, to.index)
  let cells = 0
  for (let index = low; index < high; index += 1) cells += sizeOf(index)
  return span + (from.index <= to.index ? cells : -cells)
}

function EditableShapeVisual({
  file,
  visual,
  worksheet,
  allowText,
  chartEditing,
  onEdit,
  onTextEdit,
  frame,
}: {
  readonly file: VisualHost
  readonly visual: WorkbookVisualObject
  readonly worksheet: UniverWorksheet
  readonly allowText: boolean
  readonly chartEditing?: ChartEditing | undefined
  readonly onEdit: (visualId: string, changes: ShapeEditChanges) => void
  readonly frame?: ShapeFrame | undefined
  readonly onTextEdit?: ((visualId: string, changes: ShapeEditChanges) => Promise<void>) | undefined
}): React.JSX.Element {
  const [drag, setDrag] = useState<{
    mode: 'move' | 'resize'
    corner: ResizeCorner
    startX: number
    startY: number
    dx: number
    dy: number
  } | null>(null)
  const [textEditing, setTextEditing] = useState(false)
  const textBusy = useRef(false),
    [textPending, setTextPending] = useState(false),
    [textError, setTextError] = useState('')
  useEffect(() => {
    if (!textEditing) return
    openShapeEditors++
    return () => {
      openShapeEditors--
    }
  }, [textEditing])
  const [chartEditRequest, setChartEditRequest] = useState(0)
  const [menu, setMenu] = useState<{ x: number; y: number } | null>(null)
  const editorRef = useRef<HTMLDivElement | null>(null)
  // Move previews render as a clone on <body>: the Univer float container is
  // overflow-hidden and pinned to the old anchor, so translating the visual
  // inside it clips everything that leaves the original box.
  const ghostRef = useRef<HTMLDivElement | null>(null)
  const moveDeltaRef = useRef({ dx: 0, dy: 0 })
  const chartClickTargetRef = useRef<Element | null>(null)
  const isSelected = useIsSelected(visual.id)
  const isChart = visual.kind === 'chart' && visual.chart !== undefined

  const dragActiveRef = useRef(false)
  const setDragActive = (active: boolean): void => {
    if (dragActiveRef.current === active) return
    dragActiveRef.current = active
    activeDragCount += active ? 1 : -1
  }

  const removeGhost = (host: HTMLElement | null): void => {
    ghostRef.current?.remove()
    ghostRef.current = null
    if (host) host.style.visibility = ''
  }

  const hostNodeRef = useRef<HTMLElement | null>(null)

  const cancelDrag = (): void => {
    chartClickTargetRef.current = null
    if (!dragActiveRef.current) return
    removeGhost(hostNodeRef.current)
    setDrag(null)
    setDragActive(false)
  }
  const cancelDragRef = useRef(cancelDrag)
  cancelDragRef.current = cancelDrag
  useEffect(() => {
    const cancel = () => cancelDragRef.current()
    const onKey = (event: KeyboardEvent) => {
      if (event.key !== 'Escape' || !dragActiveRef.current) return
      event.preventDefault()
      event.stopPropagation()
      cancel()
    }
    window.addEventListener('blur', cancel)
    window.addEventListener('keydown', onKey, true)
    return () => {
      window.removeEventListener('blur', cancel)
      window.removeEventListener('keydown', onKey, true)
    }
  }, [])

  useEffect(() => {
    if (hostNodeRef.current) claimDropGhost(visual.id, hostNodeRef.current)
    return () => {
      ghostRef.current?.remove()
      ghostRef.current = null
      setDragActive(false)
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  const commitDrag = (
    mode: 'move' | 'resize',
    screenCorner: ResizeCorner,
    dxRaw: number,
    dyRaw: number,
  ): boolean => {
    const zoom = worksheet.getZoom() || 1
    const anchor = visual.anchor
    // The lazy grid is data-sized; past its edge walk with the sheet
    // defaults — install grows the grid to the committed anchor, and Univer
    // gives the new rows/columns exactly these default sizes.
    const config = worksheet.getSheet().getConfig()
    const gridColumns = worksheet.getMaxColumns()
    const gridRows = worksheet.getMaxRows()
    const columnWidth = (index: number): number =>
      index < gridColumns
        ? Math.max(worksheet.getColumnWidth(index), 1)
        : Math.max(config.defaultColumnWidth, 1)
    const rowHeight = (index: number): number =>
      index < gridRows
        ? Math.max(worksheet.getRowHeight(index), 1)
        : Math.max(config.defaultRowHeight, 1)
    const maxColumn = XLSX_MAX_COLUMN
    const maxRow = XLSX_MAX_ROW
    let fromX = markerFrom(anchor.fromColumn, anchor.fromColumnOffset)
    let fromY = markerFrom(anchor.fromRow, anchor.fromRowOffset)
    // Drag from the clamped geometry the user sees, not the raw overflowing
    // offsets — otherwise walkMarker re-expands an explicit-to anchor to its
    // walked size on commit. The rebuilt anchor is fully normalized, so it
    // rightly drops the explicitTo flag.
    const explicitTo = anchor.explicitTo === true
    let toX = clampExplicitTo(
      markerFrom(anchor.toColumn, anchor.toColumnOffset),
      explicitTo,
      columnWidth,
    )
    let toY = clampExplicitTo(markerFrom(anchor.toRow, anchor.toRowOffset), explicitTo, rowHeight)
    // RTL sheets render mirrored geometry from logical anchors, so a screen
    // drag maps to the opposite logical x-shift and the opposite logical edge.
    const rtl = config.rightToLeft === BooleanNumber.TRUE
    const corner = rtl ? mirrorCornerX(screenCorner) : screenCorner
    const dx = (rtl ? -dxRaw : dxRaw) / zoom
    const dy = dyRaw / zoom
    if (mode === 'move') {
      // Free placement: keep the frame size. The grid edge caps
      // the shift — walk the leading marker first and apply the distance it
      // actually covered to both markers, so a clamped drag cannot shrink
      // the frame.
      const effectiveShift = (
        delta: number,
        from: AnchorMarker,
        to: AnchorMarker,
        sizeOf: (index: number) => number,
        maxIndex: number,
      ): number => {
        if (delta === 0) return 0
        const leading = delta > 0 ? to : from
        return markerSpan(leading, walkMarker(leading, delta, sizeOf, maxIndex), sizeOf)
      }
      const shiftX = effectiveShift(dx, fromX, toX, columnWidth, maxColumn)
      const shiftY = effectiveShift(dy, fromY, toY, rowHeight, maxRow)
      fromX = walkMarker(fromX, shiftX, columnWidth, maxColumn)
      fromY = walkMarker(fromY, shiftY, rowHeight, maxRow)
      toX = walkMarker(toX, shiftX, columnWidth, maxColumn)
      toY = walkMarker(toY, shiftY, rowHeight, maxRow)
    } else {
      // Resize handles: each edge moves its own anchor marker,
      // never past the opposite edge minus the minimum frame.
      if (cornerEast(corner)) {
        toX = walkMarker(toX, dx, columnWidth, maxColumn)
        if (markerSpan(fromX, toX, columnWidth) < MIN_FRAME_PIXELS) {
          toX = walkMarker(fromX, MIN_FRAME_PIXELS, columnWidth, maxColumn)
        }
      } else if (cornerWest(corner)) {
        fromX = walkMarker(fromX, dx, columnWidth, maxColumn)
        if (markerSpan(fromX, toX, columnWidth) < MIN_FRAME_PIXELS) {
          fromX = walkMarker(toX, -MIN_FRAME_PIXELS, columnWidth, maxColumn)
        }
      }
      if (cornerSouth(corner)) {
        toY = walkMarker(toY, dy, rowHeight, maxRow)
        if (markerSpan(fromY, toY, rowHeight) < MIN_FRAME_PIXELS) {
          toY = walkMarker(fromY, MIN_FRAME_PIXELS, rowHeight, maxRow)
        }
      } else if (cornerNorth(corner)) {
        fromY = walkMarker(fromY, dy, rowHeight, maxRow)
        if (markerSpan(fromY, toY, rowHeight) < MIN_FRAME_PIXELS) {
          fromY = walkMarker(toY, -MIN_FRAME_PIXELS, rowHeight, maxRow)
        }
      }
    }
    let frameSize: { width: number; height: number } | undefined
    if (mode === 'resize' && visual.rotation && visual.frameWidth && visual.frameHeight) {
      // The dragged box is the shape's new rotated AABB; the file's true
      // frame (xfrm ext) must change with it or the next install re-centers
      // the old AABB and the resize snaps back. Solve the ext whose rotated
      // AABB is the box; near 45° the system is singular, so scale the old
      // ext uniformly to the box's half-perimeter instead.
      const boxWidth = markerSpan(fromX, toX, columnWidth)
      const boxHeight = markerSpan(fromY, toY, rowHeight)
      const radians = (visual.rotation * Math.PI) / 180
      const cos = Math.abs(Math.cos(radians))
      const sin = Math.abs(Math.sin(radians))
      const det = cos * cos - sin * sin
      let extWidth = Math.abs(det) > 0.1 ? (cos * boxWidth - sin * boxHeight) / det : 0
      let extHeight = Math.abs(det) > 0.1 ? (cos * boxHeight - sin * boxWidth) / det : 0
      if (extWidth < MIN_FRAME_PIXELS || extHeight < MIN_FRAME_PIXELS) {
        const oldWidth = visual.frameWidth / EMU_PER_PIXEL
        const oldHeight = visual.frameHeight / EMU_PER_PIXEL
        const scale = Math.max(
          (boxWidth + boxHeight) / ((cos + sin) * (oldWidth + oldHeight)),
          MIN_FRAME_PIXELS / Math.min(oldWidth, oldHeight),
        )
        extWidth = oldWidth * scale
        extHeight = oldHeight * scale
      }
      // Keep the file invariant (anchor = the ext's rotated AABB, shared
      // center): re-derive the markers from the box center.
      const aabbWidth = cos * extWidth + sin * extHeight
      const aabbHeight = sin * extWidth + cos * extHeight
      fromX = walkMarker(fromX, (boxWidth - aabbWidth) / 2, columnWidth, maxColumn)
      fromY = walkMarker(fromY, (boxHeight - aabbHeight) / 2, rowHeight, maxRow)
      toX = walkMarker(fromX, aabbWidth, columnWidth, maxColumn)
      toY = walkMarker(fromY, aabbHeight, rowHeight, maxRow)
      frameSize = {
        width: Math.max(1, Math.round(extWidth * EMU_PER_PIXEL)),
        height: Math.max(1, Math.round(extHeight * EMU_PER_PIXEL)),
      }
    }
    const next = {
      fromRow: fromY.index,
      fromColumn: fromX.index,
      fromRowOffset: Math.round(fromY.offset * EMU_PER_PIXEL),
      fromColumnOffset: Math.round(fromX.offset * EMU_PER_PIXEL),
      toRow: toY.index,
      toColumn: toX.index,
      toRowOffset: Math.round(toY.offset * EMU_PER_PIXEL),
      toColumnOffset: Math.round(toX.offset * EMU_PER_PIXEL),
    }
    if (
      next.fromRow === anchor.fromRow &&
      next.fromColumn === anchor.fromColumn &&
      next.fromRowOffset === anchor.fromRowOffset &&
      next.fromColumnOffset === anchor.fromColumnOffset &&
      next.toRow === anchor.toRow &&
      next.toColumn === anchor.toColumn &&
      next.toRowOffset === anchor.toRowOffset &&
      next.toColumnOffset === anchor.toColumnOffset
    )
      return false
    pendingFocusId = visual.id
    onEdit(visual.id, { anchor: next, ...(frameSize ? { frameSize } : {}) })
    return true
  }

  return (
    <div
      className={`shape-editable${isSelected ? ' selected' : ''}`}
      tabIndex={0}
      ref={(node) => {
        hostNodeRef.current = node
        if (!node || pendingFocusId !== visual.id) return
        pendingFocusId = null
        // The float container may not be attached/laid out yet; retry until
        // focus actually lands. preventScroll: a scroll-into-view here would
        // shift the sheet under the just-dropped visual.
        const tryFocus = (attempts: number): void => {
          node.focus({ preventScroll: true })
          if (document.activeElement !== node && attempts > 0) {
            setTimeout(() => tryFocus(attempts - 1), 80)
          }
        }
        tryFocus(12)
      }}
      onPointerDown={(event) => {
        if (textEditing || event.button !== 0) return
        // Chart editor buttons/inputs keep their own interactions.
        if ((event.target as HTMLElement).closest('button, .chart-editor')) return
        if (event.shiftKey) {
          event.preventDefault()
          event.stopPropagation()
          selectVisual(visual, true)
          event.currentTarget.focus({ preventScroll: true })
          return
        }
        // A plain click collapses a multi-selection onto this object.
        if (selectedVisualIds().length > 1) selectVisual(visual)
        const handleCorner = (event.target as HTMLElement).dataset['corner'] as
          ResizeCorner | undefined
        const mode = handleCorner ? ('resize' as const) : ('move' as const)
        chartClickTargetRef.current = isChart && !handleCorner ? (event.target as Element) : null
        event.currentTarget.setPointerCapture(event.pointerId)
        setDragActive(true)
        moveDeltaRef.current = { dx: 0, dy: 0 }
        setDrag({
          mode,
          corner: handleCorner ?? 'se',
          startX: event.clientX,
          startY: event.clientY,
          dx: 0,
          dy: 0,
        })
        event.preventDefault()
        event.stopPropagation()
        // preventScroll: grabbing a half-visible visual must not scroll it
        // into view mid-drag — the sheet would shift under the pointer.
        event.currentTarget.focus({ preventScroll: true })
      }}
      onFocus={() => {
        if (!selectedVisualIds().includes(visual.id)) selectVisual(visual)
      }}
      onContextMenu={
        isChart || textEditing
          ? undefined
          : (event) => {
              event.preventDefault()
              event.stopPropagation()
              if (!selectedVisualIds().includes(visual.id)) selectVisual(visual)
              event.currentTarget.focus({ preventScroll: true })
              setMenu({ x: event.clientX, y: event.clientY })
            }
      }
      onKeyDown={(event) => {
        if (event.target !== event.currentTarget) return
        if (event.key === 'Delete' || event.key === 'Backspace') {
          event.preventDefault()
          event.stopPropagation()
          if (selectedVisualIds().length > 1 && selectedVisualIds().includes(visual.id)) {
            visualCommandListener?.('visual:delete', visual.id)
          } else {
            onEdit(visual.id, { remove: true })
            clearVisualSelection(visual.id)
          }
        } else if (event.key === 'Escape') {
          event.preventDefault()
          cancelDrag()
          event.currentTarget.blur()
          clearVisualSelection(visual.id)
        }
      }}
      onPointerMove={(event) => {
        if (!drag) return
        const dx = event.clientX - drag.startX
        const dy = event.clientY - drag.startY
        if (drag.mode === 'resize') {
          setDrag({ ...drag, dx, dy })
          return
        }
        // Move: drive the body-level ghost directly, no re-render per event.
        moveDeltaRef.current = { dx, dy }
        const host = event.currentTarget
        if (!ghostRef.current && (Math.abs(dx) > 2 || Math.abs(dy) > 2)) {
          const rect = host.getBoundingClientRect()
          const container = document.createElement('div')
          container.className = 'shape-drag-ghost'
          container.style.left = `${rect.left}px`
          container.style.top = `${rect.top}px`
          container.style.width = `${rect.width}px`
          container.style.height = `${rect.height}px`
          container.appendChild(host.cloneNode(true))
          document.body.appendChild(container)
          // Pointer capture keeps delivering events to the hidden host.
          host.style.visibility = 'hidden'
          ghostRef.current = container
        }
        if (ghostRef.current) {
          ghostRef.current.style.transform = `translate(${dx}px, ${dy}px)`
        }
      }}
      onPointerUp={(event) => {
        if (!drag) return
        const { mode, corner } = drag
        const { dx, dy } = mode === 'move' ? moveDeltaRef.current : drag
        setDrag(null)
        setDragActive(false)
        const committed = (Math.abs(dx) > 2 || Math.abs(dy) > 2) && commitDrag(mode, corner, dx, dy)
        // Pointer capture retargets the browser's click to the outer frame.
        // Delegate a genuine non-drag press to its original chart hit element,
        // preserving existing title/axis/series/slice handlers and drag capture.
        const hit = chartClickTargetRef.current
        chartClickTargetRef.current = null
        if (
          !committed &&
          mode === 'move' &&
          Math.abs(dx) <= 2 &&
          Math.abs(dy) <= 2 &&
          hit?.isConnected &&
          hit !== event.currentTarget &&
          event.currentTarget.contains(hit) &&
          event.currentTarget.hasPointerCapture(event.pointerId)
        ) {
          hit.dispatchEvent(
            new MouseEvent('click', {
              bubbles: true,
              cancelable: true,
              view: window,
              detail: 1,
              button: 0,
              clientX: event.clientX,
              clientY: event.clientY,
              ctrlKey: event.ctrlKey,
              metaKey: event.metaKey,
              altKey: event.altKey,
              shiftKey: event.shiftKey,
            }),
          )
        }
        if (committed && ghostRef.current) {
          parkDropGhost(visual.id, ghostRef.current, event.currentTarget)
          ghostRef.current = null
        } else {
          removeGhost(event.currentTarget)
        }
      }}
      onPointerCancel={cancelDrag}
      onLostPointerCapture={cancelDrag}
      onDoubleClick={
        allowText
          ? () => setTextEditing(true)
          : isChart
            ? (event) => {
                if ((event.target as HTMLElement).closest('button, .chart-editor')) return
                setChartEditRequest((count) => count + 1)
              }
            : undefined
      }
      data-tip={
        textEditing
          ? undefined
          : allowText
            ? t('appVisualHintText')
            : isChart
              ? t('appVisualHintEdit')
              : t('appVisualHint')
      }
    >
      {visual.kind === 'chart' && visual.chart ? (
        <ChartVisual
          chart={visual.chart}
          visualId={visual.id}
          chartPath={visual.chartPath}
          chartEditing={chartEditing}
          frame={frame}
          onRemove={() => {
            onEdit(visual.id, { remove: true })
            clearVisualSelection(visual.id)
          }}
          openEditorSignal={chartEditRequest}
        />
      ) : visual.kind === 'image' ? (
        <ImageVisual file={file} visual={visual} frame={frame} />
      ) : (
        <ShapeVisual
          file={file}
          visual={textEditing ? { ...visual, text: '' } : visual}
          frame={frame}
        />
      )}
      {menu && (
        <VisualContextMenu
          x={menu.x}
          y={menu.y}
          onClose={() => setMenu(null)}
          items={visualMenuItems(
            visual.id,
            visual.kind,
            () => {
              onEdit(visual.id, { remove: true })
              clearVisualSelection(visual.id)
            },
            chartEditing?.visualCommands,
          )}
        />
      )}
      {shouldShowVisualDeleteButton({ selected: isSelected, textEditing }) && (
        <VisualDeleteButton
          hostRef={hostNodeRef}
          worksheet={worksheet}
          label={t('appDeleteVisualTitle')}
          onDelete={() => onEdit(visual.id, { remove: true })}
        />
      )}
      {textEditing && (
        <div
          className="shape-text-editor"
          style={shapeTextScaleStyle(frame?.width)}
          contentEditable={!textPending}
          suppressContentEditableWarning
          role="textbox"
          ref={(node) => {
            if (!node || editorRef.current === node) return
            editorRef.current = node
            node.textContent = visual.text ?? ''
            node.focus()
            const selection = window.getSelection()
            const range = document.createRange()
            range.selectNodeContents(node)
            selection?.removeAllRanges()
            selection?.addRange(range)
          }}
          onBlur={async () => {
            if (textBusy.current) return
            const next = editorRef.current?.textContent ?? ''
            if (next === (visual.text ?? '')) {
              setTextEditing(false)
              editorRef.current = null
              return
            }
            textBusy.current = true
            setTextPending(true)
            setTextError('')
            try {
              await (onTextEdit
                ? onTextEdit(visual.id, { text: next })
                : onEdit(visual.id, { text: next }))
              setTextEditing(false)
              editorRef.current = null
            } catch (reason) {
              setTextError(reason instanceof Error ? reason.message : String(reason))
            } finally {
              textBusy.current = false
              setTextPending(false)
            }
          }}
          onKeyDown={(event) => {
            if (textBusy.current) {
              event.preventDefault()
              return
            }
            if (event.key === 'Escape') {
              event.preventDefault()
              if (editorRef.current) editorRef.current.textContent = visual.text ?? ''
              event.currentTarget.blur()
            } else if (event.key === 'Enter' && !event.shiftKey) {
              event.preventDefault()
              event.currentTarget.blur()
            }
          }}
        />
      )}
      {textEditing && textError && (
        <div className="chart-editor-error" role="alert">
          {textError}
        </div>
      )}
      {drag?.mode === 'resize' && (
        <div
          className="shape-resize-outline"
          style={{
            left: cornerWest(drag.corner) ? drag.dx : 0,
            top: cornerNorth(drag.corner) ? drag.dy : 0,
            width: cornerEast(drag.corner)
              ? `calc(100% + ${drag.dx}px)`
              : cornerWest(drag.corner)
                ? `calc(100% - ${drag.dx}px)`
                : '100%',
            height: cornerSouth(drag.corner)
              ? `calc(100% + ${drag.dy}px)`
              : cornerNorth(drag.corner)
                ? `calc(100% - ${drag.dy}px)`
                : '100%',
          }}
        />
      )}
      {!textEditing &&
        RESIZE_CORNERS.map((corner) => (
          <span
            key={corner}
            className={`shape-handle handle-${corner}${corner === 'se' ? ' shape-resize-handle' : ''}`}
            data-handle="resize"
            data-corner={corner}
          />
        ))}
    </div>
  )
}

/// prstDash → dash/gap units in multiples of the line width (DrawingML's
/// preset patterns); "solid" and unknown values fall through to a solid
/// stroke.
const LINE_DASH_PATTERNS: Record<string, readonly number[]> = {
  dash: [4, 3],
  dashDot: [4, 3, 1, 3],
  dot: [1, 3],
  lgDash: [8, 3],
  lgDashDot: [8, 3, 1, 3],
  lgDashDotDot: [8, 3, 1, 3, 1, 3],
  sysDash: [3, 1],
  sysDashDot: [3, 1, 1, 1],
  sysDashDotDot: [3, 1, 1, 1, 1, 1],
  sysDot: [1, 1],
}

/// DrawingML gradient angle (degrees clockwise, 0 = left-to-right) to SVG
/// linearGradient endpoints in objectBoundingBox units. The bounding-box
/// stretch matches lin@scaled gradients closely enough for cell-anchored
/// shapes.
function gradientEndpoints(angle: number): { x1: number; y1: number; x2: number; y2: number } {
  const radians = (angle * Math.PI) / 180
  const dx = Math.cos(radians) / 2
  const dy = Math.sin(radians) / 2
  const round = (value: number): number => Math.round(value * 1000) / 1000
  return { x1: round(0.5 - dx), y1: round(0.5 - dy), x2: round(0.5 + dx), y2: round(0.5 + dy) }
}

function ShapeVisual({
  file,
  visual,
  frame,
}: {
  readonly file?: VisualHost | undefined
  readonly visual: WorkbookVisualObject
  readonly frame?: ShapeFrame | undefined
}): React.JSX.Element {
  // useId's delimiters (":"/"«») are hostile to url(#…) parsing.
  const shapeId = useId().replace(/[^A-Za-z0-9]/g, '')
  const gradientId = `shape-fill-${shapeId}`
  const clipId = `shape-clip-${shapeId}`
  // a:blipFill: the image paints clipped to the geometry, under the stroke.
  const { url: fillImage } = useWorkbookMediaUrl(
    visual.fillMediaPath && file ? file.sessionId : undefined,
    visual.id,
    file?.readMedia,
  )
  const type = visual.shapeType ?? ''
  const gradient = visual.fillGradient
  const customPath = visual.customPath
  // "none" (explicit <a:noFill/>, or nothing to fill from at all) is
  // transparent; inserted shapes always carry an explicit fill.
  const fill = gradient
    ? `url(#${gradientId})`
    : visual.fillColor === 'none' || visual.fillColor === undefined
      ? 'transparent'
      : visual.fillColor
  // A rotated shape's anchor stores its rotated bounds (Excel writes the
  // quadrant-swapped snap rect, LibreOffice the AABB) while xfrm ext keeps
  // the true unrotated frame; both center the anchor on the shape center.
  // Restore the ext-sized box around that center, then rotate it. Without
  // ext, fall back to swapping the anchor box inside the swap quadrants.
  const rotation = (((visual.rotation ?? 0) % 360) + 360) % 360
  const trueFrame =
    rotation !== 0 && frame && visual.frameWidth && visual.frameHeight
      ? { width: visual.frameWidth / EMU_PER_PIXEL, height: visual.frameHeight / EMU_PER_PIXEL }
      : undefined
  const swapFallback =
    rotation !== 0 &&
    !trueFrame &&
    ((rotation >= 45 && rotation < 135) || (rotation >= 225 && rotation < 315))
  // Preset geometry is aspect-dependent (snips/rounds scale with the short
  // side); use the real frame proportions when the install-time layout is
  // known instead of a stretched 100×100 square.
  const boxWidth = trueFrame
    ? Math.max(1, trueFrame.width)
    : frame
      ? Math.max(1, swapFallback ? frame.height : frame.width)
      : 100
  const boxHeight = trueFrame
    ? Math.max(1, trueFrame.height)
    : frame
      ? Math.max(1, swapFallback ? frame.width : frame.height)
      : 100
  // An unflipped OOXML line runs top-left → bottom-right; flips mirror the
  // endpoints (the gallery preview draws ascending, which is the flipV form).
  const isStraightLine = type === 'line' || type === 'straightConnector1'
  const lineX0 = visual.flipH ? boxWidth : 0
  const lineY0 = visual.flipV ? boxHeight : 0
  // Same geometry source as the gallery previews (and the other apps'
  // renderers), so every insertable preset draws its real silhouette.
  // custGeom paths keep their own coordinate space and scale into the box.
  const d = customPath
    ? customPath.d
    : isStraightLine
      ? `M ${lineX0} ${lineY0} L ${boxWidth - lineX0} ${boxHeight - lineY0}`
      : shapePreviewPath(type, boxWidth, boxHeight)
  if (!d) {
    // Unsupported geometry: never leak the internal shape name (Excel shows
    // no frame at all) — an empty frame only when the shape carries text.
    if (!visual.text) return <div aria-hidden="true" />
    return (
      <div className="xlsx-shape">
        <span>{visual.text}</span>
      </div>
    )
  }
  const stroke =
    visual.lineColor === 'none' || visual.lineColor === undefined ? 'transparent' : visual.lineColor
  // Session text edits overwrite `text` but not `paragraphs`; a mismatch
  // means the flat text is the newer truth.
  const paragraphs =
    visual.paragraphs &&
    visual.paragraphs
      .map((paragraph) => paragraph.runs.map((run) => run.text).join(''))
      .join('\n') === (visual.text ?? '')
      ? visual.paragraphs
      : undefined
  const markers = paragraphs ? shapeParagraphMarkers(paragraphs) : []
  const overflowClass = shapeTextOverflowClass(visual.textVertOverflow, visual.textHorzOverflow)
  const transforms: string[] = []
  if (visual.rotation) transforms.push(`rotate(${visual.rotation}deg)`)
  if (!isStraightLine && (visual.flipH || visual.flipV)) {
    transforms.push(`scale(${visual.flipH ? -1 : 1}, ${visual.flipV ? -1 : 1})`)
  }
  const strokeWidth = visual.lineWidth ?? 1
  let dashPattern = visual.lineDash ? LINE_DASH_PATTERNS[visual.lineDash] : undefined
  // Round-capped dot patterns render as true circles via zero-length dashes:
  // the stretched viewBox distorts dash lengths but not the cap circle.
  if (dashPattern && visual.lineCap === 'rnd' && dashPattern[0] === 1 && dashPattern.length === 2) {
    dashPattern = [0, dashPattern[1]! + 1]
  }
  const content = (
    <>
      <svg viewBox={`0 0 ${boxWidth} ${boxHeight}`} preserveAspectRatio="none" aria-hidden="true">
        {gradient && !isStraightLine && (
          <defs>
            <linearGradient id={gradientId} {...gradientEndpoints(gradient.angle)}>
              {gradient.stops.map((stop, index) => (
                <stop key={index} offset={stop.position} stopColor={stop.color} />
              ))}
            </linearGradient>
          </defs>
        )}
        {customPath?.fillD && (
          // Mixed custGeom: fill only the fillable subpaths; the stroke pass
          // below covers every subpath with a transparent fill.
          <path
            d={customPath.fillD}
            fill={fill}
            stroke="none"
            transform={`scale(${boxWidth / customPath.width}, ${boxHeight / customPath.height})`}
          />
        )}
        {fillImage && !isStraightLine && (
          <>
            <clipPath id={clipId}>
              <path
                d={d}
                {...(customPath
                  ? {
                      transform: `scale(${boxWidth / customPath.width}, ${boxHeight / customPath.height})`,
                    }
                  : {})}
              />
            </clipPath>
            <image
              href={fillImage}
              x={0}
              y={0}
              width={boxWidth}
              height={boxHeight}
              preserveAspectRatio="none"
              clipPath={`url(#${clipId})`}
              {...(visual.opacity !== undefined && visual.opacity < 1
                ? { opacity: visual.opacity }
                : {})}
            />
          </>
        )}
        <path
          d={d}
          fill={
            isStraightLine || customPath?.strokeOnly || customPath?.fillD || fillImage
              ? 'transparent'
              : fill
          }
          {...(customPath
            ? {
                transform: `scale(${boxWidth / customPath.width}, ${boxHeight / customPath.height})`,
              }
            : {})}
          stroke={stroke}
          strokeWidth={strokeWidth}
          {...(dashPattern
            ? { strokeDasharray: dashPattern.map((unit) => unit * strokeWidth).join(' ') }
            : {})}
          {...(visual.lineCap === 'rnd'
            ? { strokeLinecap: 'round' as const }
            : visual.lineCap === 'sq'
              ? { strokeLinecap: 'square' as const }
              : {})}
          vectorEffect="non-scaling-stroke"
        />
      </svg>
      {paragraphs ? (
        <span
          className={`shape-text shape-text-rich shape-anchor-${visual.textAnchor ?? 'ctr'}${overflowClass}`}
          style={visual.textColor ? { color: visual.textColor } : undefined}
        >
          {paragraphs.map((paragraph, index) => (
            <span
              key={index}
              className="shape-paragraph"
              style={{
                textAlign:
                  paragraph.align === 'ctr'
                    ? 'center'
                    : paragraph.align === 'r'
                      ? 'right'
                      : paragraph.align === 'just'
                        ? 'justify'
                        : 'left',
                ...shapeParagraphIndentStyle(paragraph),
              }}
            >
              {markers[index] !== undefined && (
                <span
                  className="shape-bullet"
                  style={{
                    ...shapeRunStyle(paragraph.runs[0], false),
                    ...(shapeBulletWidth(paragraph) ? { width: shapeBulletWidth(paragraph) } : {}),
                  }}
                >
                  {shapeBulletWidth(paragraph) ? markers[index] : `${markers[index]}\u00a0`}
                </span>
              )}
              {paragraph.runs.length === 0
                ? '\u00a0'
                : paragraph.runs.map((run, runIndex) => (
                    <span key={runIndex} style={shapeRunStyle(run, true)}>
                      {run.text}
                    </span>
                  ))}
            </span>
          ))}
        </span>
      ) : (
        visual.text && (
          <span
            className={`shape-text${overflowClass}`}
            style={visual.textColor ? { color: visual.textColor } : undefined}
          >
            {visual.text}
          </span>
        )
      )}
    </>
  )
  if (trueFrame || swapFallback) {
    // cq units keep the inner box proportional to the float container so it
    // tracks zoom and live resizes.
    const size =
      trueFrame && frame
        ? {
            width: `${(trueFrame.width / frame.width) * 100}cqw`,
            height: `${(trueFrame.height / frame.height) * 100}cqh`,
          }
        : { width: '100cqh', height: '100cqw' }
    return (
      <div
        className="xlsx-shape-drawn xlsx-shape-rotated"
        style={shapeTextScaleStyle(frame?.width)}
      >
        <div
          className="xlsx-shape-rotated-inner"
          style={{ ...size, transform: ['translate(-50%, -50%)', ...transforms].join(' ') }}
        >
          {content}
        </div>
      </div>
    )
  }
  return (
    <div
      className="xlsx-shape-drawn"
      style={{
        ...shapeTextScaleStyle(frame?.width),
        ...(transforms.length ? { transform: transforms.join(' ') } : {}),
      }}
    >
      {content}
    </div>
  )
}

function ImageVisual({
  file,
  visual,
  frame,
}: {
  readonly file: VisualHost
  readonly visual: WorkbookVisualObject
  readonly frame?: ShapeFrame | undefined
}): React.JSX.Element {
  const picture = <ImagePicture file={file} visual={visual} />
  const rotation = visual.rotation ?? 0
  if (!rotation && !visual.flipH && !visual.flipV) return picture
  // Same box recipe as ShapeVisual: the anchor holds the rotated AABB, the
  // xfrm ext the true frame, both centered on the picture center.
  const trueFrame =
    rotation && frame && visual.frameWidth && visual.frameHeight
      ? { width: visual.frameWidth / EMU_PER_PIXEL, height: visual.frameHeight / EMU_PER_PIXEL }
      : undefined
  const transforms: string[] = []
  if (rotation) transforms.push(`rotate(${rotation}deg)`)
  if (visual.flipH || visual.flipV) {
    transforms.push(`scale(${visual.flipH ? -1 : 1}, ${visual.flipV ? -1 : 1})`)
  }
  // cq units keep the inner box proportional to the float container so it
  // tracks zoom and live resizes (same recipe as ShapeVisual).
  const size =
    trueFrame && frame
      ? {
          width: `${(trueFrame.width / frame.width) * 100}cqw`,
          height: `${(trueFrame.height / frame.height) * 100}cqh`,
        }
      : { width: '100cqw', height: '100cqh' }
  return (
    <div className="xlsx-image-rotated">
      <div
        className="xlsx-image-rotated-inner"
        style={{ ...size, transform: ['translate(-50%, -50%)', ...transforms].join(' ') }}
      >
        {picture}
      </div>
    </div>
  )
}

function ImagePicture({
  file,
  visual,
}: {
  readonly file: VisualHost
  readonly visual: WorkbookVisualObject
}): React.JSX.Element {
  // Session-added images carry their bytes inline; nothing to fetch.
  const { url, unavailable } = useWorkbookMediaUrl(
    visual.mediaDataUrl ? undefined : file.sessionId,
    visual.id,
    file.readMedia,
  )
  const source = visual.mediaDataUrl ?? url

  if (unavailable) return <div className="xlsx-visual-unavailable" />
  if (!source) return <div className="xlsx-visual-loading">{t('appImageLoading')}</div>
  const opacity =
    visual.opacity !== undefined && visual.opacity < 1 ? { opacity: visual.opacity } : undefined
  const crop = visual.crop
  // a:srcRect: stretch the source so the cropped window fills the frame
  // (Excel fits the remaining region exactly, no letterboxing).
  const visibleWidth = crop ? 1 - crop.left - crop.right : 1
  const visibleHeight = crop ? 1 - crop.top - crop.bottom : 1
  if (crop && visibleWidth > 0 && visibleHeight > 0) {
    return (
      <span className="xlsx-image-crop">
        <img
          src={source}
          alt={visual.altText ?? visual.name ?? t('appWorkbookImageAlt')}
          style={{
            width: `${100 / visibleWidth}%`,
            height: `${100 / visibleHeight}%`,
            left: `${(-100 * crop.left) / visibleWidth}%`,
            top: `${(-100 * crop.top) / visibleHeight}%`,
            ...opacity,
          }}
        />
      </span>
    )
  }
  return (
    <img
      className="xlsx-image"
      src={source}
      alt={visual.altText ?? visual.name ?? t('appWorkbookImageAlt')}
      style={opacity}
    />
  )
}

/// Embedded OLE object: Excel's cached preview picture (EMF/WMF rasterized
/// like any other metafile) at the object's anchor, or — with no preview or
/// an undecodable one — Excel's icon-and-caption box. Always read-only: the
/// object lives in <oleObjects>/xl/embeddings, outside the drawing edit
/// pipeline, so it is never moved, deleted or re-saved as a picture.
function OleVisual({
  file,
  visual,
}: {
  readonly file: VisualHost
  readonly visual: WorkbookVisualObject
}): React.JSX.Element {
  const { url, unavailable } = useWorkbookMediaUrl(
    visual.mediaPath === undefined ? undefined : file.sessionId,
    visual.id,
    file.readMedia,
  )
  const caption = oleCaption(visual.progId)
  // The legacy VML shape's stroke/fill are document colours (Excel's
  // window/windowText by default): the frame Excel draws around every
  // embedded object and the backdrop that hides the grid behind it.
  const frame = oleFrameStyle(visual)
  if (oleRenderKind(visual, unavailable) === 'preview') {
    if (!url) return <div className="xlsx-visual-loading">{t('appImageLoading')}</div>
    return (
      <div className="xlsx-ole-frame" style={frame}>
        <img className="xlsx-image xlsx-ole-preview" src={url} alt={caption} draggable={false} />
      </div>
    )
  }
  return (
    <div
      className="xlsx-ole-placeholder"
      style={frame}
      role="img"
      aria-label={caption}
      title={caption}
    >
      <svg className="xlsx-ole-icon" viewBox="0 0 24 24" aria-hidden="true">
        <path
          d="M6 2h8l5 5v15H6z"
          fill="none"
          stroke="currentColor"
          strokeWidth="1.5"
          strokeLinejoin="round"
        />
        <path d="M14 2v5h5" fill="none" stroke="currentColor" strokeWidth="1.5" />
        <path d="M9 12h7M9 15h7M9 18h5" stroke="currentColor" strokeWidth="1.5" />
      </svg>
      <span className="xlsx-ole-caption">{caption}</span>
    </div>
  )
}

/// Slicer controls are not rendered; the frame and caption header keep
/// Excel's footprint so the surrounding layout reads the same.
function SlicerVisual({ visual }: { readonly visual: WorkbookVisualObject }): React.JSX.Element {
  const caption = visual.text ?? visual.name ?? ''
  return (
    <div className="xlsx-slicer-placeholder" role="img" aria-label={caption} title={caption}>
      <div className="xlsx-slicer-caption">{caption}</div>
    </div>
  )
}

const CHART_TYPE_OPTIONS = [
  { value: 'column', labelKey: 'appChartColumn' },
  { value: 'bar', labelKey: 'appChartBar' },
  { value: 'line', labelKey: 'appChartLine' },
  { value: 'area', labelKey: 'appChartTypeArea' },
  { value: 'pie', labelKey: 'appChartPie' },
  { value: 'doughnut', labelKey: 'appChartDoughnut' },
] as const

/// Applies journaled edits over the parsed chart cache so the on-screen
/// chart previews the pending save (shared with the baked-edit path).
function withChartEdit(chart: ChartMetadata, edit: ChartEditData | undefined): ChartMetadata {
  return applyChartStateEdit(chart, edit) as ChartMetadata
}

/// A chart converts when it is one supported plot type, or a bar+line combo
/// ('combo' — converting collapses the whole chart onto the target family).
/// Single-plot column/bar/line/area/pie/doughnut charts convert; everything
/// else (scatter, radar, richer combos, 3D) fails closed at save time —
/// reject up front.
export function convertibleType(chart: ChartMetadata): string | null {
  if (isComboChartTypes(chart.chartTypes)) return 'combo'
  if (chart.chartTypes.length !== 1) return null
  const only = chart.chartTypes[0]
  if (only === 'lineChart') return 'line'
  if (only === 'areaChart') return 'area'
  if (only === 'barChart') return chart.barDirection === 'bar' ? 'bar' : 'column'
  if (only === 'pieChart') return 'pie'
  if (only === 'doughnutChart') return 'doughnut'
  return null
}

function ChartVisual({
  chart: sourceChart,
  visualId,
  chartPath,
  chartEditing,
  frame,
  onRemove,
  openEditorSignal = 0,
}: {
  readonly chart: ChartMetadata
  readonly visualId: string
  readonly chartPath?: string | undefined
  readonly chartEditing?: ChartEditing | undefined
  /// Unzoomed frame size; the rendered figure over it is the sheet zoom.
  readonly frame?: ShapeFrame | undefined
  readonly onRemove?: (() => void) | undefined
  /// Bumped by the wrapper on double-click to open the inline editor.
  readonly openEditorSignal?: number
}): React.JSX.Element {
  const [isEditing, setIsEditing] = useState(false)
  const [menu, setMenu] = useState<{ x: number; y: number } | null>(null)
  useEffect(() => {
    if (openEditorSignal > 0) setIsEditing(true)
  }, [openEditorSignal])
  const selectedEl = useSyncExternalStore(subscribeChartElementSelection, () =>
    selectedChartElement?.visualId === visualId ? selectedChartElement.element : null,
  )
  // In-memory charts (session adds, demo) bake edits into the visual itself,
  // so only file charts carry a pending overlay.
  const editKey = chartPath ?? visualId
  const pendingEdit = chartPath ? chartEditing?.edits.get(chartPath) : undefined
  // Files written without numCache (openpyxl, pandas) have series refs but no
  // cached points; hydrate those series from the referenced cells. Categories
  // hydrate on their own too (scatter xVal numRef without numCache).
  const readVector = chartEditing?.readVector
  const needsHydration = sourceChart.series.some(
    (series) =>
      (series.values.length === 0 && series.valuesRef) ||
      (series.categories.length === 0 && series.categoriesRef) ||
      series.nameRef,
  )
  const [hydration, setHydration] = useState<ReadonlyMap<
    number,
    {
      values: readonly number[] | null
      categories: readonly string[] | null
      name: string | null
    }
  > | null>(null)
  useEffect(() => {
    if (!needsHydration || !readVector) return
    let cancelled = false
    void (async () => {
      // The lazy sheet may still be indexing right after open; retry briefly.
      for (let attempt = 0; attempt < 8 && !cancelled; attempt += 1) {
        const entries = new Map<
          number,
          {
            values: readonly number[] | null
            categories: readonly string[] | null
            name: string | null
          }
        >()
        let retry = false
        for (const [index, series] of sourceChart.series.entries()) {
          const valuesRef = series.values.length === 0 ? series.valuesRef : undefined
          const categoriesRef = series.categories.length === 0 ? series.categoriesRef : undefined
          const nameRef = series.nameRef
          if (valuesRef === undefined && categoriesRef === undefined && nameRef === undefined) {
            continue
          }
          try {
            const values = valuesRef === undefined ? null : await readVector(editKey, valuesRef)
            // A failed categories read only aborts (and retries) when it is
            // the sole target; alongside values it stays best-effort.
            const categories =
              categoriesRef === undefined
                ? null
                : valuesRef === undefined
                  ? await readVector(editKey, categoriesRef)
                  : await readVector(editKey, categoriesRef).catch(() => null)
            // The name (an uncached c:tx cell reference) throws into the
            // retry loop when it is the sole target — a still-indexing sheet
            // must not freeze the SeriesN placeholder. Alongside other reads
            // it stays best-effort: the placeholder already stands in.
            const name =
              nameRef === undefined
                ? null
                : valuesRef === undefined && categoriesRef === undefined
                  ? await readVector(editKey, nameRef)
                  : await readVector(editKey, nameRef).catch(() => null)
            if (values === null && categories === null && name === null) continue
            const nameText = String(name?.vector[0] ?? '').slice(0, 255)
            entries.set(index, {
              values:
                values?.vector.map((value) => {
                  const numeric =
                    typeof value === 'number' ? value : Number(String(value ?? '').trim())
                  return Number.isFinite(numeric) ? numeric : 0
                }) ?? null,
              categories:
                categories?.vector.map((value) => String(value ?? '').slice(0, 255)) ?? null,
              name: nameText === '' ? null : nameText,
            })
          } catch {
            retry = true
          }
        }
        if (!retry) {
          if (!cancelled && entries.size > 0) setHydration(entries)
          return
        }
        await new Promise((resolve) => setTimeout(resolve, 900))
      }
    })()
    return () => {
      cancelled = true
    }
  }, [needsHydration, readVector, editKey, sourceChart])
  const hydratedSource = hydration
    ? {
        ...sourceChart,
        series: sourceChart.series.map((series, index) => {
          const entry = hydration.get(index)
          if (!entry) return series
          return {
            ...series,
            ...(entry.values && series.values.length === 0 ? { values: [...entry.values] } : {}),
            ...(series.categories.length === 0 && entry.categories
              ? { categories: [...entry.categories] }
              : {}),
            ...(entry.name !== null ? { name: entry.name } : {}),
          }
        }),
      }
    : sourceChart
  const chart = withChartEdit(hydratedSource, pendingEdit)
  const areaStyle = sheetChartAreaStyle(chart)
  if (!chart.series.some((series) => series.values.length > 0))
    return <figure className="xlsx-chart" style={areaStyle} />
  const isPie = chart.chartTypes.includes('pieChart') || chart.chartTypes.includes('doughnutChart')
  const canEdit = chartEditing !== undefined && (chartEditing.canEdit?.(editKey) ?? true)
  const selectElement = canEdit
    ? (element: ChartElementRef): void => setChartElementSelection(visualId, element)
    : undefined
  const labelsOn = chart.dataLabels !== undefined && chart.dataLabels !== 'none'
  return (
    <figure
      className="xlsx-chart"
      style={areaStyle}
      onClick={() => {
        if (selectedEl) setChartElementSelection(visualId, null)
      }}
      onContextMenu={
        canEdit && chartEditing?.contextMenu !== false
          ? (event) => {
              event.preventDefault()
              event.stopPropagation()
              // Right-click also selects the chart (menu actions and the
              // format pane both key off the selection).
              ;(event.currentTarget.closest('.shape-editable') as HTMLElement | null)?.focus()
              setMenu({ x: event.clientX, y: event.clientY })
            }
          : undefined
      }
    >
      {canEdit && !isEditing && (
        <button
          className="chart-edit-button"
          data-tip={t('appEditChartTitle')}
          aria-label={t('appEditChartTitle')}
          onClick={() => setIsEditing(true)}
        >
          ✎
        </button>
      )}
      {canEdit && isEditing && chartEditing && (
        <ChartEditor
          chart={chart}
          isPie={isPie}
          convertible={convertibleType(withChartEdit(sourceChart, undefined))}
          pendingType={pendingEdit?.chartType}
          readVector={rangeReader(chartEditing, editKey)}
          onApply={async (edit) => {
            await chartEditing.onEdit(editKey, edit)
            setIsEditing(false)
          }}
          onCancel={() => setIsEditing(false)}
        />
      )}
      <SheetChartContent
        chart={chart}
        frameWidth={frame?.width}
        selectedEl={selectedEl}
        onElement={selectElement}
      />
      {menu && chartEditing && (
        <VisualContextMenu
          x={menu.x}
          y={menu.y}
          onClose={() => setMenu(null)}
          items={[
            {
              label: t('appFormatSelection'),
              disabled: chartEditing.dialogs !== undefined && !chartEditing.dialogs.has('format'),
              action: () => chartDialogListener?.(editKey, 'format'),
            },
            {
              label: t('appSelectData'),
              disabled:
                chartEditing.dialogs !== undefined && !chartEditing.dialogs.has('select-data'),
              action: () => chartDialogListener?.(editKey, 'select-data'),
            },
            ...(chartSupportsDataLabels(chart.chartTypes)
              ? [
                  {
                    label: labelsOn ? t('appDeleteDataLabels') : t('appAddDataLabels'),
                    action: () => {
                      void Promise.resolve()
                        .then(() =>
                          chartEditing.onEdit(editKey, {
                            dataLabels: labelsOn ? 'none' : isPie ? 'category-percent' : 'value',
                          }),
                        )
                        .catch((error) => chartEditing.onError?.(error))
                    },
                  },
                ]
              : []),
            ...(onRemove
              ? [
                  'separator' as const,
                  ...visualMenuItems(visualId, 'chart', onRemove, chartEditing.visualCommands),
                ]
              : []),
          ]}
        />
      )}
    </figure>
  )
}

export type VisualMenuItem =
  | 'separator'
  | {
      readonly label: string
      readonly action?: () => void
      readonly children?: readonly { readonly label: string; readonly action: () => void }[]
      readonly disabled?: boolean
    }

/// Excel for Mac's picture/shape menu, minus entries with no model here
/// (Assign Macro, Crop, Format Picture…).
function visualMenuItems(
  visualId: string,
  kind: WorkbookVisualObject['kind'],
  onDelete: () => void,
  enabled?: ReadonlySet<string>,
): VisualMenuItem[] {
  const command = (name: string) => () => visualCommandListener?.(name, visualId)
  const isChart = kind === 'chart'
  return [
    ...(isChart
      ? []
      : [
          {
            label: t('appVisualCut'),
            disabled: enabled !== undefined && !enabled.has('visual:cut'),
            action: command('visual:cut'),
          },
          {
            label: t('appVisualCopy'),
            disabled: enabled !== undefined && !enabled.has('visual:copy'),
            action: command('visual:copy'),
          },
          {
            label: t('appVisualPaste'),
            disabled: enabled !== undefined && !enabled.has('visual:paste'),
            action: command('visual:paste'),
          },
          'separator' as const,
        ]),
    {
      label: t('appVisualEditAltText'),
      disabled: enabled !== undefined && !enabled.has('visual:alt-text'),
      action: command('visual:alt-text'),
    },
    'separator',
    {
      label: t('appVisualBringToFront'),
      disabled:
        enabled !== undefined && !enabled.has('visual:z:front') && !enabled.has('visual:z:forward'),
      children: [
        { label: t('appVisualBringToFront'), action: command('visual:z:front') },
        { label: t('appVisualBringForward'), action: command('visual:z:forward') },
      ],
    },
    {
      label: t('appVisualSendToBack'),
      disabled:
        enabled !== undefined && !enabled.has('visual:z:back') && !enabled.has('visual:z:backward'),
      children: [
        { label: t('appVisualSendToBack'), action: command('visual:z:back') },
        { label: t('appVisualSendBackward'), action: command('visual:z:backward') },
      ],
    },
    'separator',
    {
      label: t('appVisualSizeProperties'),
      disabled: enabled !== undefined && !enabled.has('visual:size'),
      action: command('visual:size'),
    },
    ...(isChart
      ? []
      : [
          {
            label: t('appVisualHyperlink'),
            disabled: enabled !== undefined && !enabled.has('visual:hyperlink'),
            action: command('visual:hyperlink'),
          },
        ]),
    ...(kind === 'image'
      ? [
          {
            label: t('appVisualEmbedInCell'),
            disabled: enabled !== undefined && !enabled.has('visual:embed-in-cell'),
            action: command('visual:embed-in-cell'),
          },
        ]
      : []),
    'separator',
    { label: t(isChart ? 'appDeleteChart' : 'appVisualDelete'), action: onDelete },
  ]
}

function VisualContextMenu({
  x,
  y,
  items,
  onClose,
}: {
  readonly x: number
  readonly y: number
  readonly items: readonly VisualMenuItem[]
  readonly onClose: () => void
}): React.JSX.Element {
  const menuRef = useRef<HTMLDivElement>(null)
  const [openSubmenu, setOpenSubmenu] = useState<number | null>(null)
  // Outside press / window blur / shell chrome press — the shared dismissal
  // (the click-through backdrop below stays as belt and braces). The menu is
  // portaled to the body, so the inside guard must be its own element.
  useDismissablePopover(true, onClose, { inside: () => [menuRef.current], escape: true })
  // Portal to the body: the float-DOM container clips overflowing children.
  return createPortal(
    <div
      className="chart-menu-backdrop"
      onClick={onClose}
      onContextMenu={(event) => event.preventDefault()}
    >
      <div
        ref={menuRef}
        className="chart-menu"
        style={{
          left: Math.min(x, window.innerWidth - 200),
          top: Math.min(y, window.innerHeight - items.length * 32 - 12),
        }}
      >
        {items.map((item, index) =>
          item === 'separator' ? (
            <hr key={index} className="chart-menu-separator" />
          ) : item.children ? (
            <div
              key={item.label}
              className={`chart-menu-sub${openSubmenu === index ? ' open' : ''}`}
              onMouseEnter={() => {
                if (!item.disabled) setOpenSubmenu(index)
              }}
            >
              <button
                disabled={item.disabled}
                aria-haspopup="menu"
                aria-expanded={openSubmenu === index}
                onClick={(event) => {
                  event.stopPropagation()
                  setOpenSubmenu(index)
                }}
              >
                {item.label}
                <span className="chart-menu-caret" aria-hidden="true">
                  ▸
                </span>
              </button>
              {openSubmenu === index && (
                <div className="chart-menu chart-menu-child" role="menu">
                  {item.children.map((child) => (
                    <button
                      key={child.label}
                      onClick={(event) => {
                        event.stopPropagation()
                        child.action()
                        onClose()
                      }}
                    >
                      {child.label}
                    </button>
                  ))}
                </div>
              )}
            </div>
          ) : (
            <button
              key={item.label}
              disabled={item.disabled}
              onMouseEnter={() => setOpenSubmenu(null)}
              onClick={(event) => {
                event.stopPropagation()
                item.action?.()
                onClose()
              }}
            >
              {item.label}
            </button>
          ),
        )}
      </div>
    </div>,
    document.body,
  )
}

function rangeReader(
  chartEditing: ChartEditing,
  editKey: string,
): ((range: string) => Promise<ChartVectorRead>) | undefined {
  const readVector = chartEditing.readVector
  if (!readVector) return undefined
  return (range) => readVector(editKey, range)
}

/// Inline editor: title, chart type (axis-based family only), series colors.
interface SeriesRangeInput {
  readonly values: string
  readonly categories: string
}

function ChartEditor({
  chart,
  isPie,
  convertible,
  pendingType,
  readVector,
  onApply,
  onCancel,
}: {
  readonly chart: ChartMetadata
  readonly isPie: boolean
  readonly convertible: string | null
  readonly pendingType: ChartEditData['chartType'] | undefined
  readonly readVector?: ((range: string) => Promise<ChartVectorRead>) | undefined
  readonly onApply: (edit: ChartEditData) => void | Promise<void>
  readonly onCancel: () => void
}): React.JSX.Element {
  const [title, setTitle] = useState(chart.title)
  const [chartType, setChartType] = useState(pendingType ?? convertible ?? '')
  const [colors, setColors] = useState<Record<string, string>>({})
  const [sliceColors, setSliceColors] = useState<Record<string, string>>({})
  const [ranges, setRanges] = useState<Record<string, SeriesRangeInput>>({})
  const [rangeError, setRangeError] = useState<string | null>(null)
  const [isReading, setIsReading] = useState(false)
  const [isApplying, setIsApplying] = useState(false)
  const applyingRef = useRef(false)
  const applyReleaseRef = useRef<(() => void) | null>(null)
  useEffect(() => () => applyReleaseRef.current?.(), [])
  // Keep drafts mounted during unrelated updates. An accepted hosted apply
  // replaces the projection before its promise resolves, so release this
  // installation guard while committing (the controls stay disabled).
  useEffect(() => (isApplying ? undefined : trackChartEditorOpen()), [isApplying])

  async function seriesRangeEdits(): Promise<NonNullable<ChartEditData['series']>> {
    if (!readVector) return []
    const entries: NonNullable<ChartEditData['series']>[number][] = []
    for (const [key, input] of Object.entries(ranges)) {
      const valuesRange = input.values.trim()
      const categoriesRange = input.categories.trim()
      if (!valuesRange && !categoriesRange) continue
      entries.push({
        index: Number(key),
        ...(valuesRange
          ? await readVector(valuesRange).then((read) => ({
              values: read.vector.map((value) => {
                const numeric =
                  typeof value === 'number' ? value : Number(String(value ?? '').trim())
                return Number.isFinite(numeric) ? numeric : 0
              }),
              valuesRef: read.ref,
            }))
          : {}),
        ...(categoriesRange
          ? await readVector(categoriesRange).then((read) => ({
              categories: read.vector.map((value) => String(value ?? '').slice(0, 255)),
              categoriesRef: read.ref,
            }))
          : {}),
      })
    }
    return entries
  }

  async function handleApply(): Promise<void> {
    if (applyingRef.current) return
    applyingRef.current = true
    const release = trackChartApply()
    applyReleaseRef.current = release
    try {
      let series: NonNullable<ChartEditData['series']>
      try {
        setIsReading(true)
        setRangeError(null)
        series = await seriesRangeEdits()
      } catch (error: unknown) {
        setRangeError(error instanceof Error ? error.message : t('appReadRangeFailed'))
        applyingRef.current = false
        return
      } finally {
        setIsReading(false)
      }
      const edit: ChartEditData = {
        ...(title !== chart.title ? { title } : {}),
        // Compare against the effective type (pending conversion included) so
        // selecting the original type back is a real edit, not a no-op.
        ...(convertible && chartType && chartType !== (pendingType ?? convertible)
          ? { chartType: chartType as NonNullable<ChartEditData['chartType']> }
          : {}),
        ...(Object.keys(colors).length > 0 ? { seriesColors: colors } : {}),
        ...(Object.keys(sliceColors).length > 0 ? { pointColors: { '0': sliceColors } } : {}),
        ...(series.length > 0 ? { series } : {}),
      }
      if (Object.keys(edit).length === 0) {
        applyingRef.current = false
        onCancel()
        return
      }
      setIsApplying(true)
      try {
        await onApply(edit)
      } catch (error) {
        setRangeError(error instanceof Error ? error.message : t('appReadRangeFailed'))
      } finally {
        applyingRef.current = false
        setIsApplying(false)
      }
    } finally {
      release()
      if (applyReleaseRef.current === release) applyReleaseRef.current = null
    }
  }

  return (
    // stopPropagation: the Univer float container forwards wheel events to
    // the sheet canvas — without this the grid scrolls (dragging the chart
    // and this panel along with it) instead of the panel's own list.
    <div
      className="chart-editor"
      aria-busy={isReading || isApplying}
      onWheel={(event) => event.stopPropagation()}
      onKeyDown={(event) => {
        // Enter/arrow/delete belong to the property inputs, not the sheet.
        event.stopPropagation()
        if (event.key === 'Escape' && !applyingRef.current) onCancel()
      }}
    >
      <fieldset disabled={isReading || isApplying}>
        <label>
          {t('appTitle')}
          <input
            className="chart-editor-title"
            value={title}
            maxLength={255}
            onChange={(event) => setTitle(event.target.value)}
          />
        </label>
        {convertible && (
          <div className="chart-editor-field">
            <span>{t('appTypeLabel')}</span>
            <Dropdown
              className="chart-editor-type"
              ariaLabel={t('appTypeLabel')}
              value={chartType}
              options={[
                // A combo has no single current type: the placeholder labels
                // the mixed state and is not pickable (picking a real family
                // converts the whole chart).
                ...(convertible === 'combo'
                  ? [{ value: 'combo', label: t('appChartCombo'), disabled: true }]
                  : []),
                ...CHART_TYPE_OPTIONS.map((option) => ({
                  value: option.value,
                  label: t(option.labelKey),
                })),
              ]}
              onPick={setChartType}
            />
          </div>
        )}
        {!isPie &&
          chart.series.slice(0, 6).map((series, index) => (
            <label key={index}>
              {truncateLabel(series.name || t('appSeriesN', { n: index + 1 }), 14)}
              <ColorDropdown
                label={series.name || t('appSeriesN', { n: index + 1 })}
                value={colors[String(index)] ?? seriesColor(series, index)}
                portal
                onPick={(hex) => {
                  if (hex) setColors((previous) => ({ ...previous, [String(index)]: hex }))
                }}
              />
            </label>
          ))}
        {isPie &&
          chart.series[0]?.categories.slice(0, 12).map((category, index) => (
            <label key={`slice-${index}`}>
              {truncateLabel(category || t('appSliceN', { n: index + 1 }), 14)}
              <ColorDropdown
                label={category || t('appSliceN', { n: index + 1 })}
                value={sliceColors[String(index)] ?? pieSliceColor(chart.series[0] ?? {}, index)}
                portal
                onPick={(hex) => {
                  if (hex) setSliceColors((previous) => ({ ...previous, [String(index)]: hex }))
                }}
              />
            </label>
          ))}
        {readVector &&
          chart.series.slice(0, 6).map((series, index) => (
            <div className="chart-editor-series" key={`ranges-${index}`}>
              <span className="chart-editor-series-name">
                {t('appSeriesData', {
                  name: truncateLabel(series.name || t('appSeriesN', { n: index + 1 }), 20),
                })}
              </span>
              <input
                className="chart-editor-range"
                placeholder={series.valuesRef ?? t('appValuesPlaceholder')}
                data-tip={t('appValuesRangeTitle', {
                  name: series.name || t('appSeriesN', { n: index + 1 }),
                })}
                value={ranges[String(index)]?.values ?? ''}
                onChange={(event) =>
                  setRanges((previous) => ({
                    ...previous,
                    [String(index)]: {
                      categories: '',
                      ...previous[String(index)],
                      values: event.target.value,
                    },
                  }))
                }
              />
              <input
                className="chart-editor-range"
                placeholder={series.categoriesRef ?? t('appLabelsPlaceholder')}
                data-tip={t('appCategoriesRangeTitle', {
                  name: series.name || t('appSeriesN', { n: index + 1 }),
                })}
                value={ranges[String(index)]?.categories ?? ''}
                onChange={(event) =>
                  setRanges((previous) => ({
                    ...previous,
                    [String(index)]: {
                      values: '',
                      ...previous[String(index)],
                      categories: event.target.value,
                    },
                  }))
                }
              />
            </div>
          ))}
        {rangeError && <div className="chart-editor-error">{rangeError}</div>}
        <div className="chart-editor-actions">
          <button
            className="chart-editor-cancel"
            disabled={isReading || isApplying}
            onClick={onCancel}
          >
            {t('appCancel')}
          </button>
          <button
            className="chart-editor-apply"
            disabled={isReading || isApplying}
            onClick={() => void handleApply()}
          >
            {isReading || isApplying ? t('appReading') : t('appApply')}
          </button>
        </div>
      </fieldset>
    </div>
  )
}
