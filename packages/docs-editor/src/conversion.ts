/** Thin adapter: keep the desktop's rich OOXML bridge instead of rebuilding HTML. */
import { readSections, type ParsedDocFull } from '@genoffice/docx-engine'
import { blocksToPmDoc, pmDocOptions } from '../../../apps/docs/src/renderer/editor/convert'

export function documentToEditor(parsed: ParsedDocFull) {
  return blocksToPmDoc(parsed.blocks, readSections(parsed), pmDocOptions(parsed))
}
export {
  pmDocToSavePlan as editorToSavePlan,
  pmDocOptions,
  type PmNode as EditorJSON,
} from '../../../apps/docs/src/renderer/editor/convert'
export {
  docStyleCss,
  docThemeCss,
  docBodyFont,
  docLineFactor,
  docHasCjk,
} from '../../../apps/docs/src/renderer/doc-style-css'
