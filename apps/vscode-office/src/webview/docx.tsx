import { createRoot } from 'react-dom/client'
import { DocsEditor } from '@genoffice/docs-editor/DocsEditor'
import { DocumentRuntime } from '@genoffice/docs-editor/runtime'
import type {
  Commit,
  DocumentProjection,
  DocumentPresentation,
} from '@genoffice/docs-editor/shared/protocol'
import { JsonBridge } from '@genoffice/xlsx-gateway/shared/session-bridge'
import { setModuleLang } from '../../../docs/src/renderer/i18n/locale'
import { normalizeLang } from '@genoffice/i18n'
import './docx-theme.css'

declare const acquireVsCodeApi: () => { postMessage(message: unknown): void }
const api = acquireVsCodeApi()
const sessionId = document.documentElement.dataset.sessionId
if (!sessionId) throw new Error('Missing DOCX session identity')
const bridge = new JsonBridge(sessionId, {
  post: (message) => api.postMessage(message),
  subscribe(listener) {
    const handler = (event: MessageEvent) => listener(event.data)
    window.addEventListener('message', handler)
    return () => window.removeEventListener('message', handler)
  },
})
const root = createRoot(document.getElementById('root')!)
const present = (presentation: DocumentPresentation) => {
  let style = document.getElementById('office-docx-styles')
  if (!style) {
    style = document.createElement('style')
    style.id = 'office-docx-styles'
    document.head.append(style)
  }
  style.textContent = presentation.css
}
let runtime: DocumentRuntime | undefined
const showError = (message: string) => {
  void bridge.request('view.presentation', { error: message.slice(0, 4096) }).catch(() => {})
  let notice = document.querySelector<HTMLDivElement>('.office-docx-error')
  if (!notice) {
    notice = document.createElement('div')
    notice.className = 'office-docx-error'
    document.body.append(notice)
  }
  notice.textContent = message
}
bridge.handle('view.flush', async (raw) => {
  const ready = (await runtime?.flush()) ?? false
  if (ready && (raw as { freeze?: boolean })?.freeze) runtime!.freeze(true)
  return { ready }
})
bridge.handle('state.committed', (raw) => {
  runtime?.accept(raw as Commit)
  return { ready: true }
})
bridge.handle('view.selection', () => runtime?.selection() ?? null)
bridge.handle('view.resync', (raw) => {
  runtime?.reset(raw as DocumentProjection)
  return { ready: true }
})
bridge.handle('appearance.changed', (raw) => {
  document.documentElement.dataset.docxAppearance =
    (raw as { appearance: string }).appearance === 'vscode' ? 'vscode' : 'word'
  return { ready: true }
})
void bridge
  .request('view.ready', {})
  .then((raw) => {
    const projection = raw as DocumentProjection
    setModuleLang(normalizeLang(document.documentElement.lang))
    if (projection.presentation) present(projection.presentation)
    runtime = new DocumentRuntime(projection, {
      submit: async (request) => (await bridge.request('view.intent', request)) as Commit,
      history: async (direction) => {
        await bridge.request('history.request', { direction })
      },
      resync: async () => (await bridge.request('view.resync', {})) as DocumentProjection,
      error: showError,
      operations: async (request) => (await bridge.request('docx.operations', request)) as Commit,
      presentation: present,
    })
    root.render(
      <DocsEditor
        runtime={runtime}
        onReady={() => {
          void bridge
            .request('view.presentation', { ready: true })
            .catch((error) => showError(String(error)))
        }}
      />,
    )
  })
  .catch((error) => showError(String(error)))
window.addEventListener('unload', () => {
  runtime?.dispose()
  bridge.dispose()
})
