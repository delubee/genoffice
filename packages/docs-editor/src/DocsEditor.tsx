import { useEffect, useState } from 'react'
import { OfficeFileActions, OFFICE_ACTION_LABELS } from '@genoffice/ui/chrome'
import { EditorContent } from '@tiptap/react'
import type { DocumentRuntime } from './runtime'
import { getLang, LocaleScope, t } from '../../../apps/docs/src/renderer/i18n/locale'
import { FontDialog } from '../../../apps/docs/src/renderer/components/FontDialog'
import { ParagraphDialog } from '../../../apps/docs/src/renderer/components/ParagraphDialog'
import type { StyleInfo } from '@genoffice/docx-engine'
import { DocsRibbon } from './DocsRibbon'
import { setDocsHost } from '../../../apps/docs/src/renderer/docs-host'
import { FindPanel } from '../../../apps/docs/src/renderer/components/FindPanel'
import { LinkInsertModal } from '../../../apps/docs/src/renderer/components/ribbon-insert-tab'
import { HeaderFooterArea } from '../../../apps/docs/src/renderer/components/HeaderFooterArea'
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
import { submitNote, type NotePrompt } from '../../../apps/docs/src/renderer/review-actions'
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
  const [hfEdit, setHfEdit] = useState<{ header: number | null; footer: number | null }>({
    header: null,
    footer: null,
  })
  const strip = (kind: 'header' | 'footer') => {
    const value = runtime.settings[kind] ?? { text: '' }
    // Word shows an empty strip only while editing it.
    if (!value.text && !value.paras?.length && hfEdit[kind] == null) return null
    return (
      <HeaderFooterArea
        kind={kind}
        value={value}
        readOnly={!editor.isEditable}
        editRequest={hfEdit[kind]}
        onCommit={(next) => void runtime.applySettings({ [kind]: next })}
        onEditingChange={(editing) => {
          if (!editing) setHfEdit((current) => ({ ...current, [kind]: null }))
        }}
      />
    )
  }
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
    const unavailable = (feature: string) => () =>
      Promise.reject(new Error(`${feature} is not available in the VS Code editor yet.`))
    setDocsHost({
      pickImage: pickImageFile,
      openNewTab: unavailable('New Window'),
      listDocsTabs: unavailable('Switch Windows'),
      focusDocsTab: unavailable('Switch Windows'),
      zoteroCommand: unavailable('Zotero'),
      print: unavailable('Printing'),
    })
    return () => setDocsHost(null)
  }, [])
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
          panes={{
            nav: showNav,
            styles: showStylesPane,
            onNav: setShowNav,
            onStyles: setShowStylesPane,
          }}
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
        <div className="office-docx-canvas" onContextMenu={openMenu}>
          <div
            className="office-docx-page doc-page"
            style={pageStyle(runtime.settings.section, runtime.settings.pageColor)}
          >
            {strip('header')}
            <EditorContent editor={editor} />
            {strip('footer')}
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
