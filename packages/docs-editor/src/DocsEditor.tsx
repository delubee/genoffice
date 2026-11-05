import { useEffect, useRef, useState } from 'react'
import { OfficeFileActions, OFFICE_ACTION_LABELS } from '@genoffice/ui/chrome'
import { EditorContent } from '@tiptap/react'
import type { DocumentRuntime } from './runtime'
import type { SectionEdit } from './shared/protocol'

type HfKeyName = 'header' | 'footer' | 'headerFirst' | 'footerFirst' | 'headerEven' | 'footerEven'
import { getLang, LocaleScope, t } from '../../../apps/docs/src/renderer/i18n/locale'
import { FontDialog } from '../../../apps/docs/src/renderer/components/FontDialog'
import { ParagraphDialog } from '../../../apps/docs/src/renderer/components/ParagraphDialog'
import type { StyleInfo } from '@genoffice/docx-engine'
import { DocsRibbon, numberingContext } from './DocsRibbon'
import { PageNumberDialog } from './PageNumberDialog'
import { ZoomDialog } from '../../../apps/docs/src/renderer/components/ZoomDialog'
import { InkOverlay } from '../../../apps/docs/src/renderer/components/InkOverlay'
import { rasterizeStroke, type InkTool } from '../../../apps/docs/src/renderer/editor/ink'
import type { InkPenSettings } from '../../../apps/docs/src/renderer/components/ribbon-tabs'
import { ProtectDialog } from '../../../apps/docs/src/renderer/components/ProtectDialog'
import { ComparePanel } from '../../../apps/docs/src/renderer/components/ComparePanel'
import {
  blockTexts,
  compareParagraphs,
  editorBlockTexts,
  type CompareEntry,
} from '../../../apps/docs/src/renderer/editor/compare'
import { parseDocx, applySectionSettings } from '@genoffice/docx-engine'
import {
  CellMarginsDialog,
  DeleteCellsDialog,
  InsertCellsDialog,
  SplitCellsDialog,
  type TableDialogKind,
} from '../../../apps/docs/src/renderer/components/TableDialogs'
import { TablePropertiesDialog } from '../../../apps/docs/src/renderer/components/TablePropertiesDialog'
import { ensureCaretInTable } from '../../../apps/docs/src/renderer/editor/table-ops'
import {
  AdjustListIndentsDialog,
  DefineBulletDialog,
  DefineMultilevelDialog,
  DefineNumberFormatDialog,
  SetNumberingValueDialog,
  type ListDialogKind,
} from '../../../apps/docs/src/renderer/components/ListDialogs'
import {
  addPictureBullet,
  createCustomListDef,
  currentListTarget,
  currentListValue,
  editListLevel,
  hasPreviousList,
  setNumberingValue,
} from '../../../apps/docs/src/renderer/numbering-actions'
import { customLevelFromNumberingLevel } from '@genoffice/docx-engine'
import { rememberListPreset } from '../../../apps/docs/src/renderer/list-presets'
import type { CustomNumberingLevel } from '@genoffice/docx-engine'
import { setDocsHost } from '../../../apps/docs/src/renderer/docs-host'
import { FindPanel } from '../../../apps/docs/src/renderer/components/FindPanel'
import { LinkInsertModal } from '../../../apps/docs/src/renderer/components/ribbon-insert-tab'
import { HeaderFooterArea } from '../../../apps/docs/src/renderer/components/HeaderFooterArea'
import type {
  HfAction,
  HfEditingInfo,
} from '../../../apps/docs/src/renderer/components/ribbon-hf-tab'
import {
  insertHfField,
  insertHfInstrField,
  insertHfPicture,
  type HfEditorHandle,
} from '../../../apps/docs/src/renderer/editor/hf-editor'
import { hfHasPageField, hfWithoutPageMarks } from '../../../apps/docs/src/renderer/editor/hf-dom'
import { PAGE_MARK } from '@genoffice/docx-engine'
import { CommentsPanel } from '../../../apps/docs/src/renderer/components/CommentsPanel'
import {
  cancelNewComment,
  commentThreadAtCaret,
  deleteAllComments,
  deleteComment,
  deleteCommentAtCaret,
  editComment,
  gotoComment,
  replyToComment,
  resolveComment,
  startNewComment,
  submitNewComment,
  type ReviewContext,
} from '../../../apps/docs/src/renderer/review-actions'
import { wordRangeAtCaret } from '../../../apps/docs/src/renderer/editor/comments'
import { PromptModal } from '../../../apps/docs/src/renderer/components/PromptModal'
import { StylesPane } from '../../../apps/docs/src/renderer/components/StylesPane'
import { NavPane } from '../../../apps/docs/src/renderer/components/NavPane'
import { computeFormatState } from '../../../apps/docs/src/renderer/components/ribbon-format-state'
import type { DocDefaults, StyleUpsert } from '@genoffice/docx-engine'
import {
  handleRevision,
  submitNote,
  type NotePrompt,
} from '../../../apps/docs/src/renderer/review-actions'
import { collectRevisions, nextRevision } from '../../../apps/docs/src/renderer/editor/revisions'
import type { RevisionDisplayMode } from '../../../apps/docs/src/renderer/editor/revision-view'
import {
  EditorContextMenu,
  type ContextMenuState,
} from '../../../apps/docs/src/renderer/components/ContextMenu'

export function DocsEditor({ runtime, onReady }: { runtime: DocumentRuntime; onReady?(): void }) {
  const editor = runtime.editor
  const [, update] = useState(0)
  const [fontDialog, setFontDialog] = useState(false)
  const [paragraphDialog, setParagraphDialog] = useState(false)
  const [find, setFind] = useState<{ find: number; replace: number } | null>(null)
  const [menu, setMenu] = useState<ContextMenuState | null>(null)
  const [linkDialog, setLinkDialog] = useState(false)
  const [showComments, setShowComments] = useState(false)
  const [composing, setComposing] = useState(false)
  const [commentFocus, setCommentFocus] = useState<{ id: string; nonce: number } | null>(null)
  const comments = runtime.settings.comments ?? []
  const footnotes = runtime.settings.footnotes ?? []
  const endnotes = runtime.settings.endnotes ?? []
  const [notePrompt, setNotePrompt] = useState<NotePrompt | null>(null)
  const [showNav, setShowNav] = useState(false)
  const [zoom, setZoom] = useState(100)
  const [inkTool, setInkTool] = useState<InkTool>('select')
  const [inkPen, setInkPen] = useState<InkPenSettings>({ color: 'C00000', width: 2 })
  const [inkHighlighter, setInkHighlighter] = useState<InkPenSettings>({
    color: 'FFFF00',
    width: 10,
  })
  const inks = runtime.settings.inks ?? []
  const draw = {
    tool: inkTool,
    onTool: setInkTool,
    pen: inkPen,
    onPen: setInkPen,
    highlighter: inkHighlighter,
    onHighlighter: setInkHighlighter,
    count: inks.length,
    onClearAll: () => void runtime.applySettings({ inks: [] }),
  }
  const [zoomDialog, setZoomDialog] = useState(false)
  const [protectDialog, setProtectDialog] = useState(false)
  const [pageNumberDialog, setPageNumberDialog] = useState(false)
  const [compare, setCompare] = useState<{ otherName: string; entries: CompareEntry[] } | null>(
    null,
  )
  /** Review → Compare: a .docx the user picks in the view, diffed by paragraph as the desktop does. */
  const compareWithFile = () =>
    void pickFile('.docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document')
      .then(async (file) => {
        if (!file) return
        const other = await parseDocx(new Uint8Array(await file.arrayBuffer()))
        setCompare({
          otherName: file.name,
          entries: compareParagraphs(
            editorBlockTexts(editor.getJSON() as never, []),
            blockTexts(other.blocks),
          ),
        })
      })
      .catch((error) => runtime.reportError(error))
  const [spellcheck, setSpellcheck] = useState(false)
  const canvas = useRef<HTMLDivElement>(null)
  /** Word's fit modes against the visible canvas and the page geometry (twips → px). */
  const fitZoom = (mode: 'width' | 'page' | 'multi' | 'text') => {
    const box = canvas.current?.getBoundingClientRect()
    const page = runtime.settings.section
    if (!box || !page) return 100
    const width =
      (mode === 'text' ? page.pageWidth - page.marginLeft - page.marginRight : page.pageWidth) / 15
    const height = page.pageHeight / 15
    const fit =
      mode === 'width' || mode === 'text'
        ? (box.width - 48) / width
        : Math.min((box.width - 48) / width, (box.height - 48) / height)
    return Math.max(10, Math.min(500, Math.round(fit * 100)))
  }
  const view = {
    zoom,
    onZoom: setZoom,
    onZoomFit: (mode: 'width' | 'page' | 'multi') => setZoom(fitZoom(mode)),
    onZoomDialog: () => setZoomDialog(true),
    spellcheck,
    onSpellcheck: (on: boolean) => {
      editor.view.dom.spellcheck = on
      setSpellcheck(on)
    },
  }
  const [revisionDisplay, setRevisionDisplay] = useState<RevisionDisplayMode>('all')
  const recorder = editor.storage.trackChanges as { enabled: boolean } | undefined
  const revisions = {
    tracking: recorder?.enabled ?? false,
    count: collectRevisions(editor.state.doc).length,
    display: revisionDisplay,
    onTrack: (on: boolean) => {
      if (recorder) recorder.enabled = on
      setRevisionDisplay((mode) => mode)
      editor.view.dispatch(editor.state.tr.setMeta('trackChangesToggled', on))
    },
    onDisplay: setRevisionDisplay,
    onAccept: (all: boolean) => handleRevision(review(), 'accept', all),
    onReject: (all: boolean) => handleRevision(review(), 'reject', all),
    onGoto: (direction: 1 | -1) => {
      const target = nextRevision(editor, direction)
      if (target)
        editor
          .chain()
          .focus()
          .setTextSelection({ from: target.from, to: target.to })
          .scrollIntoView()
          .run()
    },
  }
  const [tableDialog, setTableDialog] = useState<TableDialogKind | null>(null)
  const section = runtime.settings.section
  /** Text column width the table dialogs fit grids into (twips → px). */
  const contentWidthPx = section
    ? Math.round((section.pageWidth - section.marginLeft - section.marginRight) / 15)
    : 624
  const [listDialog, setListDialog] = useState<ListDialogKind | null>(null)
  /** Define New Bullet / Number Format / Multilevel List, as the desktop applies them. */
  const applyDefinedList = (
    levels: CustomNumberingLevel[],
    kind: 'bullets' | 'numbers' | 'multi',
  ) => {
    const numId = createCustomListDef(numberingContext(runtime), levels)
    if (!numId) return
    rememberListPreset(kind, levels)
    const ilvl = editor.isActive('docListItem')
      ? Number(editor.getAttributes('docListItem').ilvl) || 0
      : 0
    editor
      .chain()
      .focus()
      .setNode('docListItem', {
        kind: levels[0]?.numFmt === 'bullet' ? 'bullet' : 'ordered',
        numId,
        ilvl,
      })
      .run()
    setListDialog(null)
  }
  const [showStylesPane, setShowStylesPane] = useState(false)
  const [pendingStyles, setPendingStyles] = useState<Record<string, StyleUpsert>>({})
  const docDefaults = (editor.storage.listNumbering as { docDefaults?: DocDefaults } | undefined)
    ?.docDefaults
  /** The desktop review actions over the session's comment list. */
  const review = (): ReviewContext => ({
    editor,
    doc: null,
    dirtyRef: { current: false },
    setStatus: () => {},
    notePrompt,
    setNotePrompt,
    footnotes,
    endnotes,
    setFootnotes: (update) =>
      runtime.setReviewList('footnotes', typeof update === 'function' ? update(footnotes) : update),
    setEndnotes: (update) =>
      runtime.setReviewList('endnotes', typeof update === 'function' ? update(endnotes) : update),
    setNotesDirty: () => {},
    comments,
    setComments: (update) =>
      runtime.setComments(typeof update === 'function' ? update(comments) : update),
    setCommentsDirty: () => {},
    setCommentComposing: setComposing,
    setShowComments,
    setCommentFocus,
    setInkAnnotations: () => {},
    setInksDirty: () => {},
    setCompareResult: () => {},
  })
  const reviewComments = {
    count: comments.length,
    open: comments.filter((comment) => !comment.parentId && comment.done !== true).length,
    resolved: comments.filter((comment) => !comment.parentId && comment.done === true).length,
    atCaret: commentThreadAtCaret(review()) !== null,
    canComment: !editor.state.selection.empty || wordRangeAtCaret(editor) !== null,
    onNew: () => startNewComment(review()),
    onShow: () => setShowComments(true),
    onDelete: () => deleteCommentAtCaret(review()),
    onDeleteAll: (resolvedOnly: boolean) => deleteAllComments(review(), resolvedOnly),
    onGoto: (direction: 1 | -1) => void gotoComment(review(), direction),
  }
  // The section under the cursor, as the desktop's activeSection: Layout and Header & Footer options edit it.
  const sectionList = runtime.settings.sections ?? []
  const activeSectionIndex = (() => {
    if (sectionList.length < 2) return 0
    const { doc, selection } = editor.state
    let docx: number | null = null
    for (let child = Math.min(selection.$head.index(0), doc.childCount - 1); child >= 0; child--) {
      const index = doc.child(child).attrs.docxIndex
      if (typeof index === 'number') {
        docx = index
        break
      }
    }
    const found = sectionList.findIndex((entry) => docx !== null && entry.blockIndex >= docx)
    return found < 0 ? sectionList.length - 1 : found
  })()
  const activeEntry = sectionList[activeSectionIndex]
  // A break inserted in this session has no place in the file yet: the section above it cannot be
  // addressed by the session's section list, so its settings wait for a save.
  const unsavedBreakAhead = (() => {
    const { doc, selection } = editor.state
    for (let i = Math.min(selection.$head.index(0), doc.childCount - 1); i < doc.childCount; i++) {
      const attrs = doc.child(i).attrs
      if (!JSON.stringify(attrs).includes('<w:sectPr')) continue
      return attrs.docxIndex == null
    }
    return false
  })()
  const unsavedSection = () => {
    if (!unsavedBreakAhead) return false
    runtime.reportError('Save the document before editing a section created by a new break.')
    return true
  }
  const activeIsFinal = !activeEntry || activeEntry.final
  // Strips of the section under the cursor: per-section state in a multi-section file, else the single parts.
  const multiSection = sectionList.length > 1
  const variantOf = (key: string): 'default' | 'first' | 'even' =>
    key.endsWith('First') ? 'first' : key.endsWith('Even') ? 'even' : 'default'
  const kindOf = (key: string): 'header' | 'footer' =>
    key.startsWith('header') ? 'header' : 'footer'
  const stripOf = (key: string) =>
    multiSection
      ? runtime.settings.sectionStrips?.find(
          (item) =>
            item.section === activeSectionIndex &&
            item.kind === kindOf(key) &&
            item.variant === variantOf(key),
        )
      : undefined
  const stripValue = (key: HfKeyName) =>
    multiSection ? (stripOf(key)?.value ?? { text: '' }) : (runtime.settings[key] ?? { text: '' })
  const writeStrip = (key: HfKeyName, hf: NonNullable<DocumentRuntime['settings']['header']>) => {
    if (multiSection && unsavedSection()) return
    if (multiSection)
      void runtime.applySettings({
        sectionHfEdit: {
          section: activeSectionIndex,
          kind: kindOf(key),
          variant: variantOf(key),
          hf,
        },
      })
    else void runtime.applySettings({ [key]: hf })
  }
  /** Apply a section edit to the section under the cursor (the final section keeps its top-level settings). */
  const editSection = (edit: SectionEdit, afterBreak = false) => {
    if (!afterBreak && unsavedSection()) return
    if (activeIsFinal) {
      void runtime.applySettings({
        ...(edit.section ? { section: edit.section } : {}),
        ...(edit.startType ? { sectionStart: edit.startType } : {}),
        ...(edit.titlePg !== undefined ? { titlePg: edit.titlePg } : {}),
        ...(edit.pageNumbering ? { pageNumbering: edit.pageNumbering } : {}),
      })
      return
    }
    const key = String(activeEntry.blockIndex)
    const edits = runtime.settings.sectionEdits ?? {}
    void runtime.applySettings({ sectionEdits: { ...edits, [key]: { ...edits[key], ...edit } } })
  }
  type HfKey = HfKeyName
  const baseKind = (key: HfKey): 'header' | 'footer' =>
    key.startsWith('header') ? 'header' : 'footer'
  const [hfEdit, setHfEdit] = useState<Partial<Record<HfKey, number | null>>>({})
  const [hfOpen, setHfOpen] = useState<HfKey | null>(null)
  const hfHandle = useRef<HfEditorHandle | null>(null)
  const hfEditing: HfEditingInfo | null = hfOpen
    ? {
        kind: baseKind(hfOpen),
        section: activeSectionIndex,
        sectionCount: Math.max(1, sectionList.length),
        variant: hfOpen.endsWith('First') ? 'first' : hfOpen.endsWith('Even') ? 'even' : 'default',
        linked: stripOf(hfOpen)?.linked ?? null,
        headerDistTwips: runtime.settings.section?.headerDist ?? 720,
        footerDistTwips: runtime.settings.section?.footerDist ?? 720,
        titlePg: activeEntry?.titlePg ?? runtime.settings.titlePg ?? false,
        evenOddHf: runtime.settings.evenAndOddHeaders ?? false,
      }
    : null
  /** The Header & Footer contextual tab, as the desktop's onHfAction for the default strips. */
  const onHfAction = (action: HfAction) => {
    const sub = hfHandle.current?.editor
    // Actions on the open strip target its variant; otherwise the default part.
    const part = (kind: 'header' | 'footer'): HfKey =>
      hfOpen && baseKind(hfOpen) === kind ? hfOpen : kind
    const value = (kind: 'header' | 'footer') => stripValue(part(kind))
    switch (action.type) {
      case 'close':
        hfHandle.current?.exit()
        break
      case 'goto':
      case 'edit':
        hfHandle.current?.exit()
        setHfEdit((current) => ({ ...current, [action.kind]: Date.now() }))
        break
      case 'remove':
        hfHandle.current?.exit()
        writeStrip(part(action.kind), { text: '' })
        break
      case 'pageNumber':
        hfHandle.current?.exit()
        writeStrip(part(action.kind), {
          text: PAGE_MARK,
          pageNumber: true,
          paras: [{ align: action.align, runs: [{ text: PAGE_MARK }] }],
        })
        break
      case 'pageNumberHere':
        if (sub) insertHfField(sub, 'PAGE')
        break
      case 'numPages':
        if (sub) insertHfField(sub, 'NUMPAGES')
        break
      case 'removePageNumbers':
        hfHandle.current?.exit()
        for (const kind of ['header', 'footer'] as const)
          if (hfHasPageField(value(kind))) writeStrip(part(kind), hfWithoutPageMarks(value(kind)))
        break
      case 'field':
        if (sub)
          insertHfInstrField(
            sub,
            action.instr,
            action.instr === 'DATE'
              ? new Date().toLocaleDateString('zh-CN')
              : new Date().toLocaleTimeString('zh-CN', { hour: '2-digit', minute: '2-digit' }),
          )
        break
      case 'picture':
        if (sub)
          void pickImageFile().then(async (picked) => {
            if (!picked || !hfHandle.current) return
            const dataUrl = `data:${picked.mime};base64,${picked.base64}`
            const image = new Image()
            image.src = dataUrl
            await image.decode()
            insertHfPicture(sub, {
              dataUrl,
              widthPx: image.naturalWidth,
              heightPx: image.naturalHeight,
            })
          }, runtime.reportError.bind(runtime))
        break
      case 'distance': {
        const current = runtime.settings.section
        if (current)
          void runtime.applySettings({
            section: {
              ...current,
              [action.kind === 'header' ? 'headerDist' : 'footerDist']: action.twips,
            },
          })
        break
      }
      case 'pageNumberFormat':
        setPageNumberDialog(true)
        break
      case 'linkToPrevious':
        if (hfOpen && multiSection)
          void runtime.applySettings({
            sectionHfLink: {
              section: activeSectionIndex,
              kind: baseKind(hfOpen),
              variant: variantOf(hfOpen),
              on: action.on,
            },
          })
        break
      case 'titlePg':
        editSection({ titlePg: action.on })
        break
      case 'evenOddHf':
        void runtime.applySettings({ evenAndOddHeaders: action.on })
        break
      default:
        runtime.reportError(
          'This header and footer action is not available in the VS Code editor yet.',
        )
    }
  }
  /**
   * Layout → Breaks: a paragraph closing the section above it with the current page geometry; the
   * following (final) section starts as the chosen type, as the desktop writes it.
   */
  const insertSectionBreak = (type: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage') => {
    const section = activeEntry?.section ?? runtime.settings.section
    if (!section) return runtime.reportError('The document has no page settings to break.')
    const sectPr = applySectionSettings(
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>',
      { ...section, pageBorder: false },
    )
    const { $head } = editor.state.selection
    editor
      .chain()
      .focus()
      .insertContentAt($head.depth > 0 ? $head.after(1) : $head.pos, {
        type: 'docProtected',
        attrs: {
          blockType: 'passthrough',
          label: 'Section break paragraph',
          previewText: '',
          genXml: `<w:p><w:pPr>${sectPr}</w:pPr></w:p>`,
        },
      })
      .run()
    // The section after the break starts as chosen: that is the type of the sectPr closing it.
    editSection({ startType: type }, true)
  }
  const strip = (key: HfKey) => {
    const kind = baseKind(key)
    const value = stripValue(key)
    // Word shows an empty strip only while editing it. A variant strip exists while its option is on.
    if (key === kind && !value.text && !value.paras?.length && hfEdit[key] == null) return null
    return (
      <HeaderFooterArea
        key={key}
        kind={kind}
        value={value}
        readOnly={!editor.isEditable}
        editRequest={hfEdit[key] ?? null}
        onCommit={(next) => writeStrip(key, next)}
        onEditingChange={(editing, handle) => {
          hfHandle.current = editing ? handle : null
          setHfOpen(editing ? key : null)
          if (!editing) setHfEdit((current) => ({ ...current, [key]: null }))
        }}
      />
    )
  }
  /** Edit entries of the variants that are switched on, as Word's first/even page headers. */
  const variantKeys = (kind: 'header' | 'footer'): HfKey[] => [
    ...((activeEntry?.titlePg ?? runtime.settings.titlePg) ? [`${kind}First` as HfKey] : []),
    ...(runtime.settings.evenAndOddHeaders ? [`${kind}Even` as HfKey] : []),
  ]
  const unavailable = (feature: string) => () =>
    runtime.reportError(`${feature} is not available in the VS Code editor yet.`)
  /** Right-click on the page, as the desktop: a click outside the selection moves the caret there. */
  const openMenu = (event: React.MouseEvent) => {
    if (!(event.target as HTMLElement).closest('.doc-page')) return
    event.preventDefault()
    const fromKeyboard = event.clientX === 0 && event.clientY === 0
    let x = event.clientX,
      y = event.clientY,
      pos: number | null = null
    if (fromKeyboard) {
      pos = editor.state.selection.head
      const coords = editor.view.coordsAtPos(pos)
      x = coords.left
      y = coords.bottom
    } else {
      const hit = editor.view.posAtCoords({ left: x, top: y })
      if (hit) {
        pos = hit.pos
        const { from, to } = editor.state.selection
        if (hit.pos < from || hit.pos > to) editor.commands.setTextSelection(hit.pos)
      }
    }
    const anchor = (event.target as HTMLElement).closest('a[href]')
    const image = (event.target as HTMLElement).closest('img')
    setMenu({
      x,
      y,
      pos,
      imageSrc: image?.getAttribute('src') ?? null,
      link: anchor ? { href: anchor.getAttribute('href') ?? '' } : null,
    })
  }
  const openFind = (replace: boolean) =>
    setFind((current) => ({
      find: (current?.find ?? 0) + (replace ? 0 : 1),
      replace: (current?.replace ?? 0) + (replace ? 1 : 0),
    }))
  const styles = editor.storage.listNumbering?.styles as Map<string, StyleInfo> | undefined
  useEffect(() => {
    const changed = () => update((value) => value + 1)
    editor.on('transaction', changed)
    onReady?.()
    return () => {
      editor.off('transaction', changed)
    }
  }, [editor, onReady])
  useEffect(() => {
    // Report in the view rather than reject: the shared ribbon does not catch these calls.
    const unavailable = (feature: string) => () => {
      runtime.reportError(`${feature} is not available in the VS Code editor yet.`)
      return Promise.resolve(undefined as never)
    }
    setDocsHost({
      pickImage: pickImageFile,
      openNewTab: unavailable('New Window'),
      listDocsTabs: unavailable('Switch Windows'),
      focusDocsTab: unavailable('Switch Windows'),
      zoteroCommand: unavailable('Zotero'),
      print: unavailable('Printing'),
    })
    return () => setDocsHost(null)
  }, [runtime])
  useEffect(() => {
    // Word dialog shortcuts at the window, as the desktop: Font Ctrl+D, Paragraph Ctrl+Alt+M.
    const onKeyDown = (event: KeyboardEvent) => {
      if (!(event.metaKey || event.ctrlKey) || event.shiftKey) return
      if (!editor.isEditable && !/^[fh]$/i.test(event.key)) return
      if (document.querySelector('[role="dialog"][aria-modal="true"]')) return
      if (!event.altKey && (event.key.toLowerCase() === 'f' || event.key.toLowerCase() === 'h')) {
        // Find / Replace open the desktop panel instead of VS Code's webview find.
        event.preventDefault()
        event.stopPropagation()
        openFind(event.key.toLowerCase() === 'h')
      } else if (!event.altKey && event.key.toLowerCase() === 'k') {
        event.preventDefault()
        event.stopPropagation()
        setLinkDialog(true)
      } else if (!event.altKey && event.key.toLowerCase() === 'd') {
        event.preventDefault()
        setFontDialog(true)
      } else if (event.altKey && event.code === 'KeyM') {
        event.preventDefault()
        setParagraphDialog(true)
      }
    }
    window.addEventListener('keydown', onKeyDown, true)
    return () => window.removeEventListener('keydown', onKeyDown, true)
  }, [editor])
  return (
    <LocaleScope lang={getLang()}>
      <main className="office-docx">
        <DocsRibbon
          runtime={runtime}
          onParagraphDialog={() => setParagraphDialog(true)}
          onHfEdit={(kind) => setHfEdit((current) => ({ ...current, [kind]: Date.now() }))}
          comments={reviewComments}
          revisions={revisions}
          view={view}
          draw={draw}
          hfEditing={hfEditing}
          onHfAction={onHfAction}
          onProtect={() => setProtectDialog(true)}
          onPageNumberFormat={() => setPageNumberDialog(true)}
          onSectionBreak={insertSectionBreak}
          sectionTarget={{
            index: sectionList.length > 1 ? activeSectionIndex : null,
            section: activeEntry?.section ?? runtime.settings.section ?? null,
            onSection: (section) => editSection({ section }),
          }}
          onCompare={compareWithFile}
          panes={{
            nav: showNav,
            styles: showStylesPane,
            onNav: setShowNav,
            onStyles: setShowStylesPane,
          }}
          onTableDialog={(kind) => {
            if (editor.isEditable && ensureCaretInTable(editor.state, editor.view.dispatch))
              setTableDialog(kind)
          }}
          onListDialog={setListDialog}

          onInsertNote={(kind) => editor.isEditable && setNotePrompt({ kind })}
          onFind={() => openFind(false)}
          onReplace={() => openFind(true)}
          quickActions={
            <div className="office-docx-quick" role="toolbar" aria-label={t('appScGroupText')}>
              {runtime.hasOfficeActions && (
                <OfficeFileActions
                  labels={OFFICE_ACTION_LABELS[getLang()]}
                  onOpen={() => runtime.officeAction({ action: 'open' })}
                  onSave={() => runtime.officeAction({ action: 'save' })}
                  onSaveAs={() => runtime.officeAction({ action: 'saveAs' })}
                  onChat={(scope) =>
                    runtime.officeAction({ action: 'chat', scope, instruction: '' })
                  }
                  onError={(error) => runtime.reportError(error)}
                />
              )}
              <button title={t('appUndo')} onClick={() => void runtime.history('undo')}>
                ↶
              </button>
              <button title={t('appRedo')} onClick={() => void runtime.history('redo')}>
                ↷
              </button>
            </div>
          }
        />
        {find && (
          <FindPanel
            editor={editor}
            onClose={() => setFind(null)}
            focusFindNonce={find.find}
            focusReplaceNonce={find.replace}
          />
        )}
        {tableDialog === 'properties' && (
          <TablePropertiesDialog
            editor={editor}
            sectionContentWidthPx={contentWidthPx}
            onClose={() => setTableDialog(null)}
          />
        )}
        {tableDialog === 'splitCells' && (
          <SplitCellsDialog editor={editor} onClose={() => setTableDialog(null)} />
        )}
        {tableDialog === 'insertCells' && (
          <InsertCellsDialog
            editor={editor}
            sectionContentWidthPx={contentWidthPx}
            onClose={() => setTableDialog(null)}
          />
        )}
        {tableDialog === 'deleteCells' && (
          <DeleteCellsDialog editor={editor} onClose={() => setTableDialog(null)} />
        )}
        {tableDialog === 'cellMargins' && (
          <CellMarginsDialog
            editor={editor}
            sectionContentWidthPx={contentWidthPx}
            onClose={() => setTableDialog(null)}
          />
        )}
        {listDialog === 'value' &&
          (() => {
            const target = currentListTarget(numberingContext(runtime))
            if (!target) return null
            return (
              <SetNumberingValueDialog
                numFmt={target.level?.numFmt ?? 'decimal'}
                lvlText={target.level?.lvlText ?? '%1.'}
                currentValue={currentListValue(editor, target.numId, target.ilvl)}
                canContinue={hasPreviousList(editor, target.numId)}
                onApply={(mode, value) => {
                  setNumberingValue(numberingContext(runtime), { mode, value })
                  setListDialog(null)
                  editor.commands.focus()
                }}
                onClose={() => setListDialog(null)}
              />
            )
          })()}
        {listDialog === 'indents' &&
          (() => {
            const target = currentListTarget(numberingContext(runtime))
            if (!target) return null
            return (
              <AdjustListIndentsDialog
                level={
                  target.level
                    ? customLevelFromNumberingLevel(target.level)
                    : { numFmt: 'decimal', lvlText: '%1.', indentLeft: 720, hanging: 360 }
                }
                onApply={(patch) => {
                  editListLevel(numberingContext(runtime), patch)
                  runtime.flushNumbering()
                  setListDialog(null)
                  editor.commands.focus()
                }}
                onClose={() => setListDialog(null)}
              />
            )
          })()}
        {listDialog === 'bullet' && (
          <DefineBulletDialog
            ilvl={Number(editor.getAttributes('docListItem').ilvl) || 0}
            onApply={(levels, picture) => {
              if (picture) {
                const id = addPictureBullet(numberingContext(runtime), picture)
                const at = Number(editor.getAttributes('docListItem').ilvl) || 0
                levels = levels.map((level, index) =>
                  index === at ? { ...level, picBulletId: id } : level,
                )
              }
              applyDefinedList(levels, 'bullets')
            }}
            onClose={() => setListDialog(null)}
          />
        )}
        {listDialog === 'number' && (
          <DefineNumberFormatDialog
            ilvl={Number(editor.getAttributes('docListItem').ilvl) || 0}
            onApply={(levels) => applyDefinedList(levels, 'numbers')}
            onClose={() => setListDialog(null)}
          />
        )}
        {listDialog === 'multi' && (
          <DefineMultilevelDialog
            styles={[...(styles?.values() ?? [])]}
            onApply={(levels) => applyDefinedList(levels, 'multi')}
            onClose={() => setListDialog(null)}
          />
        )}
        {showNav && (
          <NavPane
            editor={editor}
            doc={editor.state.doc}
            zoom={1}
            pageInfo={{ current: 1, total: 1 }}
            onClose={() => setShowNav(false)}
          />
        )}
        {showStylesPane && (
          <StylesPane
            editor={editor}
            styles={styles}
            {...(docDefaults ? { docDefaults } : {})}
            fs={computeFormatState(editor, styles, docDefaults)}
            canEdit={editor.isEditable}
            pending={pendingStyles}
            onUpsert={(style) => {
              runtime.upsertStyle(style)
              setPendingStyles((current) => ({ ...current, [style.styleId]: style }))
              return null
            }}
            onNotice={() => {}}
            onClose={() => setShowStylesPane(false)}
          />
        )}
        {notePrompt && (
          <PromptModal
            title={t(
              notePrompt.id !== undefined
                ? notePrompt.kind === 'footnote'
                  ? 'appEditFootnote'
                  : 'appEditEndnote'
                : notePrompt.kind === 'footnote'
                  ? 'appInsertFootnote'
                  : 'appInsertEndnote',
            )}
            placeholder={t(
              notePrompt.kind === 'footnote' ? 'appFootnotePlaceholder' : 'appEndnotePlaceholder',
            )}
            initial={
              notePrompt.id !== undefined
                ? ((notePrompt.kind === 'footnote' ? footnotes : endnotes).find(
                    (note) => note.id === notePrompt.id,
                  )?.text ?? '')
                : ''
            }
            multiline
            onSubmit={(text) => submitNote(review(), text)}
            onClose={() => setNotePrompt(null)}
          />
        )}
        {showComments && (
          <CommentsPanel
            comments={comments}
            docNode={editor.state.doc}
            composing={composing}
            onSubmitNew={(text) => submitNewComment(review(), text)}
            onReply={(parentId, text) => replyToComment(review(), parentId, text)}
            onEdit={(id, text) => editComment(review(), id, text)}
            onResolve={(id, done) => resolveComment(review(), id, done)}
            onCancelNew={() => cancelNewComment(review())}
            onDelete={(id) => deleteComment(review(), id)}
            onClose={() => {
              setShowComments(false)
              cancelNewComment(review())
            }}
            focus={commentFocus}
          />
        )}
        {linkDialog && <LinkInsertModal editor={editor} onClose={() => setLinkDialog(false)} />}
        {menu && (
          <EditorContextMenu
            editor={editor}
            menu={menu}
            onClose={() => setMenu(null)}
            onFontDialog={() => setFontDialog(true)}
            onParagraphDialog={() => setParagraphDialog(true)}
            onLink={() => setLinkDialog(true)}
            onNewComment={() => startNewComment(review())}
            onViewImage={unavailable('Picture viewer')}
            onSaveImageAs={unavailable('Save picture')}
            onAiPreset={(instruction) =>
              void runtime
                .officeAction({ action: 'chat', scope: 'selection', instruction })
                .catch((error) => runtime.reportError(error))
            }
          />
        )}
        {compare && (
          <ComparePanel
            otherName={compare.otherName}
            entries={compare.entries}
            onClose={() => setCompare(null)}
          />
        )}
        {protectDialog && (
          <ProtectDialog
            encrypted={false}
            writeProtection={runtime.settings.writeProtection ?? null}
            protection={runtime.settings.protection ?? null}
            removePersonalInfo={runtime.settings.removePersonalInfo ?? false}
            onCancel={() => setProtectDialog(false)}
            onApply={(result) => {
              setProtectDialog(false)
              if (result.openPassword !== undefined)
                runtime.reportError(
                  'Password-to-open encryption is not available in the VS Code editor yet.',
                )
              const { openPassword: _open, ...settings } = result
              if (Object.keys(settings).length) void runtime.applySettings(settings)
            }}
          />
        )}
        {pageNumberDialog && (
          <PageNumberDialog
            value={activeEntry?.pageNumbering ?? runtime.settings.pageNumbering ?? {}}
            onClose={() => setPageNumberDialog(false)}
            onApply={(pageNumbering) => {
              setPageNumberDialog(false)
              editSection({ pageNumbering })
            }}
          />
        )}
        {zoomDialog && (
          <ZoomDialog
            zoom={zoom}
            fitValue={fitZoom}
            onApply={(target) => setZoom(typeof target === 'number' ? target : fitZoom(target))}
            onClose={() => setZoomDialog(false)}
          />
        )}
        <div className="office-docx-canvas" ref={canvas} onContextMenu={openMenu}>
          <div
            className="office-docx-page doc-page"
            style={{
              ...pageStyle(runtime.settings.section, runtime.settings.pageColor),
              zoom: zoom / 100,
            }}
          >
            {inkTool !== 'select' || inks.length ? (
              <InkOverlay
                tool={editor.isEditable ? inkTool : 'select'}
                color={inkTool === 'highlighter' ? inkHighlighter.color : inkPen.color}
                width={inkTool === 'highlighter' ? inkHighlighter.width : inkPen.width}
                zoom={zoom}
                annotations={inks}
                onAdd={(annotation) =>
                  void runtime.applySettings({
                    inks: [
                      ...inks,
                      annotation.stroke
                        ? { ...annotation, raster: rasterizeStroke(annotation.stroke) }
                        : annotation,
                    ],
                  })
                }
                onRemove={(ids) =>
                  void runtime.applySettings({ inks: inks.filter((ink) => !ids.includes(ink.id)) })
                }
              />
            ) : null}
            {strip('header')}
            {variantKeys('header').map(strip)}
            <EditorContent editor={editor} />
            {strip('footer')}
            {variantKeys('footer').map(strip)}
            {(['footnote', 'endnote'] as const).map((kind) => {
              const list = kind === 'footnote' ? footnotes : endnotes
              if (!list.length) return null
              return (
                <ol key={kind} className={`office-docx-notes office-docx-${kind}s`}>
                  {list.map((note) => (
                    <li
                      key={note.id}
                      onDoubleClick={() =>
                        editor.isEditable && setNotePrompt({ kind, id: note.id })
                      }
                    >
                      {note.text}
                    </li>
                  ))}
                </ol>
              )
            })}
          </div>
        </div>
        {fontDialog && <FontDialog editor={editor} onClose={() => setFontDialog(false)} />}
        {paragraphDialog && (
          <ParagraphDialog
            editor={editor}
            {...(styles ? { styles } : {})}
            onClose={() => setParagraphDialog(false)}
          />
        )}
      </main>
    </LocaleScope>
  )
}

/** Insert → Pictures in the view: a local image file read in the webview, as the desktop picker returns it. */
function pickImageFile(): Promise<{
  base64: string
  mime: 'image/png' | 'image/jpeg' | 'image/gif'
  name: string
} | null> {
  return new Promise((resolve, reject) => {
    const input = document.createElement('input')
    input.type = 'file'
    input.accept = 'image/png,image/jpeg,image/gif'
    input.onchange = () => {
      const file = input.files?.[0]
      if (!file) return resolve(null)
      if (!['image/png', 'image/jpeg', 'image/gif'].includes(file.type))
        return reject(new Error('Choose a PNG, JPEG or GIF picture.'))
      if (file.size > 20 * 1024 ** 2) return reject(new Error('The picture is larger than 20 MB.'))
      const reader = new FileReader()
      reader.onload = () => {
        const url = String(reader.result)
        resolve({
          base64: url.slice(url.indexOf(',') + 1),
          mime: file.type as 'image/png' | 'image/jpeg' | 'image/gif',
          name: file.name,
        })
      }
      reader.onerror = () => reject(reader.error ?? new Error('The picture could not be read.'))
      reader.readAsDataURL(file)
    }
    input.oncancel = () => resolve(null)
    input.click()
  })
}

/** Page geometry from the final section (twips → px at 96 dpi); the page color is document content. */
function pageStyle(
  section: DocumentRuntime['settings']['section'],
  pageColor: string | null | undefined,
): React.CSSProperties | undefined {
  if (!section) return undefined
  const px = (twips: number) => `${Math.round(Math.abs(twips) / 15)}px`
  return {
    width: px(section.pageWidth),
    minHeight: px(section.pageHeight),
    padding: `${px(section.marginTop)} ${px(section.marginRight)} ${px(section.marginBottom)} ${px(section.marginLeft)}`,
    boxSizing: 'border-box',
    ...(pageColor ? { background: `#${pageColor}` } : {}),
  }
}

/** A local file chosen in the view (null when cancelled). */
function pickFile(accept: string): Promise<File | null> {
  return new Promise((resolve) => {
    const input = document.createElement('input')
    input.type = 'file'
    input.accept = accept
    input.onchange = () => resolve(input.files?.[0] ?? null)
    input.oncancel = () => resolve(null)
    input.click()
  })
}
