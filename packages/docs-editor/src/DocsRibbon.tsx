import { useState, type ComponentProps, type ReactNode } from 'react'
import type { DocDefaults, StyleInfo } from '@genoffice/docx-engine'
import { Ribbon } from '../../../apps/docs/src/renderer/components/Ribbon'
import { computeFormatState } from '../../../apps/docs/src/renderer/components/ribbon-format-state'
import { useShallowStable } from '../../../apps/docs/src/renderer/use-stable'
import type { DocumentRuntime } from './runtime'
import {
  allocateListNumId,
  createCustomListDef,
  type NumberingContext,
} from '../../../apps/docs/src/renderer/numbering-actions'
import type { DocState } from '../../../apps/docs/src/renderer/doc-state'

type RibbonProps = ComponentProps<typeof Ribbon>

/**
 * The desktop Docs ribbon bound to the VS Code DOCX runtime. Editor commands
 * run on the shared editor and enter the host session through its validated
 * steps; view toggles stay local to the view, and host capabilities without a
 * VS Code port report explicitly instead of hiding the control.
 */
export function DocsRibbon({
  runtime,
  onParagraphDialog,
  onFind,
  onReplace,
  onHfEdit,
  comments,
  onInsertNote,
  panes,
  quickActions,
}: {
  runtime: DocumentRuntime
  panes: {
    nav: boolean
    styles: boolean
    onNav(show: boolean): void
    onStyles(show: boolean): void
  }
  onInsertNote(kind: 'footnote' | 'endnote'): void
  comments: {
    count: number
    open: number
    resolved: number
    atCaret: boolean
    canComment: boolean
    onNew(): void
    onShow(): void
    onDelete(): void
    onDeleteAll(resolvedOnly: boolean): void
    onGoto(direction: 1 | -1): void
  }
  onParagraphDialog(): void
  onHfEdit(kind: 'header' | 'footer'): void
  onFind(): void
  onReplace(): void
  quickActions?: ReactNode
}) {
  const editor = runtime.editor
  const storage = editor.storage.listNumbering as
    { styles?: Map<string, StyleInfo>; docDefaults?: DocDefaults } | undefined
  const formatState = useShallowStable(
    computeFormatState(editor, storage?.styles, storage?.docDefaults),
  )
  const [zoom, setZoom] = useState(100)
  const [showMarks, setShowMarks] = useState(false)
  const [showRuler, setShowRuler] = useState(false)
  const [darkPage, setDarkPage] = useState(false)
  const [viewMode, setViewMode] = useState<RibbonProps['viewMode']>('print')
  const [readMode, setReadMode] = useState(false)
  const [showGrid, setShowGrid] = useState(false)
  const [splitView, setSplitView] = useState(false)
  const [tableGridlines, setTableGridlines] = useState(true)
  // The desktop numbering actions, over the presented definitions and the view's pending set.
  const numbering = (): NumberingContext => ({
    editor,
    doc: {
      parsed: {
        numbering: (storage as { defs?: Map<string, never> } | undefined)?.defs ?? new Map(),
      },
    } as unknown as DocState,
    pendingNumbering: runtime.pendingNumbering,
    setPendingNumbering: (update) => {
      runtime.pendingNumbering = update(runtime.pendingNumbering) as typeof runtime.pendingNumbering
    },
    numIdFloorRef: runtime.numIdFloor,
    setStatus: () => {},
  })
  const unavailable = (feature: string) => () =>
    runtime.reportError(`${feature} is not available in the VS Code editor yet.`)
  const file = (action: 'open' | 'save' | 'saveAs') => () =>
    void runtime.officeAction({ action }).catch((error) => runtime.reportError(error))
  const props: RibbonProps = {
    editor,
    formatState,
    quickActions,
    hasDoc: true,
    blocks: [],
    styles: storage?.styles,
    docDefaults: storage?.docDefaults,
    onParagraphDialog,
    onFind,
    onReplace,
    onOpen: file('open'),
    onSave: file('save'),
    onSaveAs: file('saveAs'),
    onConvert: unavailable('Export'),
    showAi: false,
    onToggleAi: () =>
      void runtime
        .officeAction({ action: 'chat', scope: 'document', instruction: '' })
        .catch((error) => runtime.reportError(error)),
    section: runtime.settings.section ?? null,
    onSection: (section) => void runtime.applySettings({ section }),
    activeSection: null,
    onInsertSectionBreak: unavailable('Section breaks'),
    onPaperSizeAll: (pageWidth, pageHeight, orientation) => {
      const section = runtime.settings.section
      if (!section) return
      const landscape = orientation === 'landscape'
      void runtime.applySettings({
        section: {
          ...section,
          orientation,
          pageWidth: landscape ? pageHeight : pageWidth,
          pageHeight: landscape ? pageWidth : pageHeight,
        },
      })
    },
    mirrorMargins: runtime.settings.mirrorMargins ?? false,
    onMirrorMargins: (mirrorMargins) => void runtime.applySettings({ mirrorMargins }),
    pageColor: runtime.settings.pageColor ?? null,
    onPageColor: (pageColor) => void runtime.applySettings({ pageColor }),
    watermark: runtime.settings.watermark ?? null,
    onWatermark: (watermark) => void runtime.applySettings({ watermark }),
    themeFonts: null,
    onThemeFonts: unavailable('Theme fonts'),
    themeColors: null,
    onThemeColors: unavailable('Theme colors'),
    inkTool: 'select',
    onInkTool: unavailable('Drawing'),
    inkPen: { color: 'C00000', width: 2 },
    onInkPen: unavailable('Drawing'),
    inkHighlighter: { color: 'FFFF00', width: 10 },
    onInkHighlighter: unavailable('Drawing'),
    inkCount: 0,
    onInkClearAll: unavailable('Drawing'),
    onInsertNote,
    sources: [],
    onAddSource: unavailable('Citations'),
    zoom,
    onZoom: setZoom,
    onZoomFit: () => setZoom(100),
    onZoomDialog: unavailable('Zoom dialog'),
    darkPage,
    onDarkPage: setDarkPage,
    onAiPreset: (instruction) =>
      void runtime
        .officeAction({ action: 'chat', scope: 'document', instruction })
        .catch((error) => runtime.reportError(error)),
    header: runtime.settings.header ?? null,
    onHeader: (header) => void runtime.applySettings({ header }),
    onPageNumFormat: unavailable('Page numbers'),
    onInsertField: unavailable('Fields'),
    footer: runtime.settings.footer ?? null,
    onFooter: (footer) => void runtime.applySettings({ footer }),
    hfEditing: null,
    onHfAction: unavailable('Header & Footer'),
    onHfEdit,
    showMarks,
    onShowMarks: setShowMarks,
    showRuler,
    onShowRuler: setShowRuler,
    showNav: panes.nav,
    onShowNav: panes.onNav,
    showStylesPane: panes.styles,
    onShowStylesPane: panes.onStyles,
    commentCount: comments.count,
    openCommentCount: comments.open,
    resolvedCommentCount: comments.resolved,
    onShowComments: comments.onShow,
    canComment: comments.canComment,
    onNewComment: comments.onNew,
    commentAtCaret: comments.atCaret,
    onDeleteComment: comments.onDelete,
    onDeleteAllComments: comments.onDeleteAll,
    onGotoComment: comments.onGoto,
    trackChanges: false,
    onTrackChanges: unavailable('Track Changes'),
    spellcheck: false,
    onSpellcheck: unavailable('Spelling'),
    revisionDisplay: 'all',
    onRevisionDisplay: unavailable('Revision display'),
    revisionCount: 0,
    onAcceptRevision: unavailable('Revisions'),
    onRejectRevision: unavailable('Revisions'),
    onGotoRevision: unavailable('Revisions'),
    isProtected: false,
    commentsAllowed: true,
    trackChangesForced: false,
    protectActive: false,
    onProtectDoc: unavailable('Document protection'),
    onCompare: unavailable('Compare'),
    mergeSource: null,
    onPickMergeDataSource: unavailable('Mail merge'),
    onMergeMailings: () =>
      Promise.reject(new Error('Mail merge is not available in the VS Code editor yet.')),
    filePath: null,
    viewMode,
    onViewMode: setViewMode,
    readMode,
    onReadMode: setReadMode,
    showGrid,
    onShowGrid: setShowGrid,
    splitView,
    onSplitView: setSplitView,
    onPagePreview: unavailable('Print preview'),
    tableGridlines,
    onToggleTableGridlines: () => setTableGridlines((value) => !value),
    onTableDialog: unavailable('Table dialogs'),
    onListDialog: unavailable('List dialogs'),
    allocateNumId: (kind) => allocateListNumId(numbering(), kind),
    createListDef: (levels) => createCustomListDef(numbering(), levels),
    onPasteDefaults: unavailable('Paste defaults'),
  }
  return (
    <div
      className="office-docx-ribbon"
      data-show-marks={showMarks || undefined}
      style={{ ['--office-docx-zoom' as string]: String(zoom / 100) }}
    >
      <Ribbon {...props} />
    </div>
  )
}
