import { zh } from './zh'
import en from './en'
export function editorText(locale = navigator.language) {
  return locale.toLowerCase().startsWith('zh') ? zh : en
}
