import { createRoot } from 'react-dom/client'
import { WorkbookEditor } from '@genoffice/sheets-editor/WorkbookEditor'
import { attachCandidateHandler, JsonBridge } from '@genoffice/sheets-editor/bridge'
import type { WorkbookRuntime } from '@genoffice/sheets-editor/runtime'
import type { EditorHost } from '@genoffice/sheets-editor/ports'
import { formatAddress } from '@genoffice/xlsx-gateway/domain/cell-address'
import './theme.css'
declare const acquireVsCodeApi: () => { postMessage(message: unknown): void }

const sessionId = document.documentElement.dataset.sessionId
if (!sessionId) throw new Error('The workbook host did not supply a session.')
const api = acquireVsCodeApi()
const bridge = new JsonBridge(sessionId, {
  post: (message) => api.postMessage(message),
  subscribe: (listener) => {
    const handler = (event: MessageEvent) => listener(event.data)
    window.addEventListener('message', handler)
    return () => window.removeEventListener('message', handler)
  },
})
const releaseCandidate = attachCandidateHandler(bridge)
let ready: (runtime: WorkbookRuntime) => void
let rejectReady: (error: unknown) => void
const projection = new Promise<WorkbookRuntime>((resolve, reject) => {
  ready = resolve
  rejectReady = reject
})
void projection.catch(() => {})
const host: EditorHost = {
  read: (logicalSheetId, range, stateId) =>
    bridge.request('source.read', {
      logicalSheetId,
      stateId,
      range: `${formatAddress(range.startRow, range.startColumn)}:${formatAddress(range.endRow, range.endColumn)}`,
    }),
  submit: (intent) =>
    bridge.request('view.intent', {
      schemaVersion: 1,
      ...intent,
      sessionId,
      origin: 'manual',
      summary: 'Manual workbook edit',
    }),
  selection: () => {}, // Binding explicitly requests the flushed selection.
  history: async (direction) => {
    await bridge.request('history.request', { direction })
  },
}
bridge.handle('state.committed', async (raw) => {
  const { descriptor, state } = raw as {
    descriptor: { stateId: string; revision: number }
    state: unknown
  }
  await (await projection).install(descriptor.stateId, descriptor.revision, state)
  return { projected: descriptor.stateId }
})
bridge.handle('view.resync', async (raw) => {
  const { descriptor, state } = raw as {
    descriptor: { stateId: string; revision: number }
    state: unknown
  }
  await (await projection).install(descriptor.stateId, descriptor.revision, state)
  return { projected: descriptor.stateId }
})
bridge.handle('view.flush', async () => ({ ready: await (await projection).flushEditor() }))
bridge.handle('view.selection', async () => (await projection).selectionContext)
bridge.handle('view.presentation', async (raw) => {
  const request = raw as { stateId?: unknown; range?: unknown }
  if (
    typeof request.stateId !== 'string' ||
    typeof request.range !== 'string' ||
    !/^[A-Z]{1,3}[1-9][0-9]{0,6}(:[A-Z]{1,3}[1-9][0-9]{0,6})?$/.test(request.range)
  )
    throw new Error('invalid_presentation_request')
  const runtime = await projection
  if (runtime.selectionContext.stateId !== request.stateId) throw new Error('stale_projection')
  await new Promise<void>((resolve) =>
    requestAnimationFrame(() => requestAnimationFrame(() => resolve())),
  )
  return {
    stateId: runtime.selectionContext.stateId,
    values: runtime.projectedValues(request.range),
    visible: !document.hidden,
  }
})
const root = createRoot(document.getElementById('root')!)
void bridge
  .request('view.ready', {})
  .then((raw) => {
    const { descriptor, state } = raw as {
      descriptor: { stateId: string; revision: number }
      state: unknown
    }
    root.render(
      <WorkbookEditor
        host={host}
        locale={document.documentElement.lang || navigator.language}
        onReady={(runtime) => {
          void runtime
            .install(descriptor.stateId, descriptor.revision, state)
            .then(() => {
              document.documentElement.dataset.editorReady = 'true'
              ready(runtime)
            })
            .catch((error) => {
              rejectReady(error)
              document.getElementById('root')!.textContent = String(error)
            })
        }}
      />,
    )
  })
  .catch((error) => {
    document.getElementById('root')!.textContent = String(error)
  })
window.addEventListener(
  'unload',
  () => {
    releaseCandidate()
    bridge.dispose()
    root.unmount()
  },
  { once: true },
)
