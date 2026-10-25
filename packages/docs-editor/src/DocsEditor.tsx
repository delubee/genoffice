import { useEffect, useState } from 'react'
import { EditorContent } from '@tiptap/react'
import type { DocumentRuntime } from './runtime'
import { t } from '../../../apps/docs/src/renderer/i18n/locale'
import { addRowAfter, addColumnAfter, deleteRow, deleteColumn } from '@tiptap/pm/tables'

export function DocsEditor({ runtime, onReady }: { runtime: DocumentRuntime; onReady?(): void }) {
  const editor = runtime.editor
  const textStyle = editor.getAttributes('docTextStyle')
  const selectionKey = `${editor.state.selection.from}:${editor.state.selection.to}`
  const [, update] = useState(0)
  useEffect(() => {
    const changed = () => update((value) => value + 1)
    editor.on('transaction', changed)
    onReady?.()
    return () => {
      editor.off('transaction', changed)
    }
  }, [editor, onReady])
  const format = (mark: string) => editor.chain().focus().toggleMark(mark).run()
  const align = (value: string) => {
    const chain = editor.chain().focus()
    for (const type of ['docParagraph', 'docHeading', 'docListItem'])
      chain.updateAttributes(type, { align: value })
    chain.run()
  }
  return (
    <main className="office-docx">
      <div className="office-docx-title">GenOffice Word</div>
      <div className="office-docx-toolbar" role="toolbar" aria-label={t('appScGroupText')}>
        <button title={t('appUndo')} onClick={() => void runtime.history('undo')}>
          ↶
        </button>
        <button title={t('appRedo')} onClick={() => void runtime.history('redo')}>
          ↷
        </button>
        <select
          aria-label={t('appParagraph')}
          defaultValue="0"
          onChange={(event) => {
            const level = Number(event.target.value)
            void runtime.applyToSelection((blockId) => ({ type: 'setHeading', blockId, level }))
          }}
        >
          {(
            [
              'ribbonStyleNormal',
              'ribbonStyleHeading1',
              'ribbonStyleHeading2',
              'ribbonStyleHeading3',
            ] as const
          ).map((key, level) => (
            <option key={key} value={level}>
              {t(key)}
            </option>
          ))}
        </select>
        <input
          key={`font:${selectionKey}:${textStyle.font ?? ''}`}
          aria-label={t('ribbonFontFamilyTip')}
          placeholder={t('ribbonFontFamilyTip')}
          list="office-fonts"
          defaultValue={textStyle.font ?? ''}
          onBlur={(event) =>
            event.target.value &&
            editor.chain().setMark('docTextStyle', { font: event.target.value }).run()
          }
          onKeyDown={(event) => {
            if (event.key === 'Enter') {
              event.preventDefault()
              event.currentTarget.blur()
              editor.commands.focus()
            }
          }}
        />
        <datalist id="office-fonts">
          {['Calibri', 'Arial', 'Times New Roman', 'Microsoft YaHei', 'SimSun'].map((font) => (
            <option key={font}>{font}</option>
          ))}
        </datalist>
        <input
          key={`size:${selectionKey}:${textStyle.sizeHalfPoints ?? ''}`}
          aria-label={t('ribbonFontSizeTip')}
          type="number"
          min="1"
          max="400"
          placeholder={t('ribbonFontSizeTip')}
          defaultValue={textStyle.sizeHalfPoints ? Number(textStyle.sizeHalfPoints) / 2 : ''}
          onBlur={(event) => {
            const points = Number(event.target.value)
            if (points >= 1 && points <= 400)
              editor
                .chain()
                .setMark('docTextStyle', { sizeHalfPoints: points * 2 })
                .run()
          }}
          onKeyDown={(event) => {
            if (event.key === 'Enter') {
              event.preventDefault()
              event.currentTarget.blur()
              editor.commands.focus()
            }
          }}
        />
        {(['bold', 'italic', 'underline', 'strike'] as const).map((mark, index) => (
          <button
            key={mark}
            aria-pressed={editor.isActive(mark)}
            title={t(
              (['appFontBold', 'appFontItalic', 'appUnderline', 'ribbonStrikethrough'] as const)[
                index
              ]!,
            )}
            onClick={() => format(mark)}
          >
            {['B', 'I', 'U', 'S'][index]}
          </button>
        ))}
        {(['left', 'center', 'right', 'justify'] as const).map((value, index) => (
          <button
            key={value}
            title={t(
              (['appScAlignLeft', 'appScAlignCenter', 'appScAlignRight', 'appScJustify'] as const)[
                index
              ]!,
            )}
            onClick={() => align(value)}
          >
            {['≡', '☰', '≡', '☷'][index]}
          </button>
        ))}
        <button
          title={t('ribbonBullets')}
          onClick={() =>
            void runtime.applyToSelection((blockId) => ({
              type: 'setList',
              blockId,
              kind: 'bullet',
            }))
          }
        >
          •
        </button>
        <button
          title={t('ribbonNumbering')}
          onClick={() =>
            void runtime.applyToSelection((blockId) => ({
              type: 'setList',
              blockId,
              kind: 'ordered',
            }))
          }
        >
          1.
        </button>
        <button
          title={t('ribbonTableTip')}
          onClick={() =>
            editor
              .chain()
              .focus()
              .insertContent({
                type: 'docTable',
                content: Array.from({ length: 2 }, () => ({
                  type: 'docTableRow',
                  content: Array.from({ length: 2 }, () => ({
                    type: 'docTableCell',
                    content: [{ type: 'docParagraph' }],
                  })),
                })),
              })
              .run()
          }
        >
          {t('ribbonTable')}
        </button>
        {(
          [
            ['ribbonInsertBelow', addRowAfter],
            ['ribbonInsertRight', addColumnAfter],
            ['ribbonDeleteRow', deleteRow],
            ['ribbonDeleteColumn', deleteColumn],
          ] as const
        ).map(([label, command]) => (
          <button
            key={label}
            title={t(label)}
            onClick={() => {
              editor.commands.focus()
              command(editor.state, editor.view.dispatch)
            }}
          >
            {t(label)}
          </button>
        ))}
      </div>
      <div className="office-docx-canvas">
        <div className="office-docx-page doc-page">
          <EditorContent editor={editor} />
        </div>
      </div>
    </main>
  )
}
