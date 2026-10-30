import { useEffect, useState } from 'react'
import { OfficeFileActions, OFFICE_ACTION_LABELS } from '@genoffice/ui/chrome'
import { EditorContent } from '@tiptap/react'
import type { DocumentRuntime } from './runtime'
import { getLang, LocaleScope, t } from '../../../apps/docs/src/renderer/i18n/locale'
import { FontDialog } from '../../../apps/docs/src/renderer/components/FontDialog'
import { ParagraphDialog } from '../../../apps/docs/src/renderer/components/ParagraphDialog'
import type { StyleInfo } from '@genoffice/docx-engine'
import { DocsRibbon } from './DocsRibbon'

export function DocsEditor({ runtime, onReady }: { runtime: DocumentRuntime; onReady?(): void }) {
  const editor = runtime.editor
  const [, update] = useState(0)
  const [fontDialog, setFontDialog] = useState(false)
  const [paragraphDialog, setParagraphDialog] = useState(false)
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
    // Word dialog shortcuts at the window, as the desktop: Font Ctrl+D, Paragraph Ctrl+Alt+M.
    const onKeyDown = (event: KeyboardEvent) => {
      if (!editor.isEditable || !(event.metaKey || event.ctrlKey) || event.shiftKey) return
      if (document.querySelector('[role="dialog"]')) return
      if (!event.altKey && event.key.toLowerCase() === 'd') {
        event.preventDefault()
        setFontDialog(true)
      } else if (event.altKey && event.code === 'KeyM') {
        event.preventDefault()
        setParagraphDialog(true)
      }
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [editor])
  return (
    <LocaleScope lang={getLang()}>
      <main className="office-docx">
        <DocsRibbon
          runtime={runtime}
          onParagraphDialog={() => setParagraphDialog(true)}
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
        <div className="office-docx-canvas">
          <div className="office-docx-page doc-page">
            <EditorContent editor={editor} />
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
