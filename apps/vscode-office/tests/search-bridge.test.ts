import { expect, it } from 'vitest'
import type { Webview } from 'vscode'
import { WebviewCandidateBridge } from '../src/extension/editor-bridge'
import { JsonBridge } from '@genoffice/xlsx-gateway/shared/session-bridge'
import { sessionSearchPageSchema } from '@genoffice/xlsx-gateway/shared/session-search'
import { initialState } from '@genoffice/workbook-session/state'
import type { WorkbookSource } from '@genoffice/workbook-session/ports'

it('routes a revision-bound indexed query over the actual JSON protocol without exposing source handles', async () => {
  let state = initialState({ sheets: [{ id: 'base', name: 'Data', rowCount: 80, columnCount: 1 }] })
  let nativeReceive!: (message: unknown) => void
  let viewReceive!: (message: unknown) => void
  let queries = 0
  const source: WorkbookSource = {
    async openSnapshot() {
      throw new Error('unexpected open')
    },
    async readBaseRange() {
      throw new Error('unexpected grid read')
    },
    async readFeatures() {
      return {}
    },
    async archiveManifest() {
      return []
    },
    async close() {},
    async findBaseCells() {
      queries++
      return {
        complete: true,
        indexingComplete: true,
        matches: [{ sheetId: 'base', row: 79, column: 0, value: 'needle', valueText: 'needle' }],
      }
    },
  }
  const sessionId = crypto.randomUUID()
  const webview = {
    postMessage(message: unknown) {
      queueMicrotask(() => viewReceive(message))
      return Promise.resolve(true)
    },
    onDidReceiveMessage(listener: (message: unknown) => void) {
      nativeReceive = listener
      return { dispose() {} }
    },
  } as unknown as Webview
  const native = new WebviewCandidateBridge({
    webview,
    sessionId,
    source,
    original: {
      path: '/not-exposed/original.xlsx',
      sha256: '0'.repeat(64),
      sourceUri: 'file:///not-exposed.xlsx',
      sidecarSessionId: crypto.randomUUID(),
    },
    getState: () => state,
    getDescriptor() {
      throw new Error('unexpected descriptor')
    },
    async submit() {
      throw new Error('search must not submit an edit')
    },
  })
  const view = new JsonBridge(sessionId, {
    post(message) {
      queueMicrotask(() => nativeReceive(message))
    },
    subscribe(listener) {
      viewReceive = listener
      return () => {}
    },
  })
  try {
    const result = sessionSearchPageSchema.parse(
      await view.request('source.find', {
        action: 'start',
        queryId: crypto.randomUUID(),
        stateId: state.stateId,
        query: 'needle',
        sheetIds: [state.sheets[0]!.id],
      }),
    )
    expect(result).toMatchObject({ total: 1, hits: [{ address: 'A80', text: 'needle' }] })
    expect(JSON.stringify(result)).not.toContain('/not-exposed')
    expect(state.sheets[0]!.cells).toEqual({})
    await expect(
      view.request('source.find', {
        action: 'start',
        queryId: crypto.randomUUID(),
        stateId: state.stateId,
        query: 'needle',
        sourceSessionId: 'attacker',
      }),
    ).rejects.toThrow()
    expect(queries).toBe(1)
    state = { ...state, stateId: crypto.randomUUID() }
    await expect(
      view.request('source.find', {
        action: 'page',
        stateId: result.stateId,
        searchId: result.searchId,
        offset: 0,
      }),
    ).rejects.toMatchObject({ code: 'revision_conflict' })
  } finally {
    native.dispose()
    view.dispose()
  }
})
