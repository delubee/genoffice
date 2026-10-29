import { expect, it } from 'vitest'
import type { Webview } from 'vscode'
import { WebviewCandidateBridge } from '../src/extension/editor-bridge'
import { initialState } from '@genoffice/workbook-session/state'
import type { WorkbookSource } from '@genoffice/workbook-session/ports'
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'

it('executes picture roots on the host without duplicating inline bytes through candidate transfers', async () => {
  let state = initialState({ sheets: [{ id: 'base', name: 'Data', rowCount: 3, columnCount: 2 }] })
  const source = {
    async readFeatures() {
      return { indexingComplete: true }
    },
  } as unknown as WorkbookSource
  const bridge = new WebviewCandidateBridge({
    webview: {
      postMessage() {
        throw new Error('Picture candidates must execute on the host')
      },
      onDidReceiveMessage() {
        return { dispose() {} }
      },
    } as unknown as Webview,
    sessionId: 'picture-host',
    source,
    original: {
      path: '/original.xlsx',
      sha256: '0'.repeat(64),
      sourceUri: 'file:///original.xlsx',
      sidecarSessionId: 'sidecar',
    },
    getState: () => state,
    getDescriptor() {
      throw new Error('unexpected descriptor')
    },
    async submit() {
      throw new Error('unexpected nested submission')
    },
  })
  const sheetId = state.sheets[0]!.id
  const bytes = Buffer.alloc(20 * 1024 * 1024)
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
    'base64',
  ).copy(bytes)
  const operation: WorkbookOperation = {
    op: 'add_picture',
    sheetId,
    name: 'maximum.png',
    mediaType: 'image/png',
    base64: bytes.toString('base64'),
    anchor: {
      fromRow: 0,
      fromColumn: 0,
      fromRowOffset: 0,
      fromColumnOffset: 0,
      toRow: 2,
      toColumn: 2,
      toRowOffset: 0,
      toColumnOffset: 0,
    },
  }
  const apply = async (operations: WorkbookOperation[]) => {
    const before = state
    const prepared = await bridge.prepare(state, {
      schemaVersion: 1,
      sessionId: 'picture-host',
      transactionId: crypto.randomUUID(),
      expectedRevision: 0,
      origin: 'manual',
      summary: 'Picture',
      operations,
    })
    expect((prepared.inverse as any).state).toEqual(before)
    state = structuredClone(prepared.nextState) as typeof state
  }
  try {
    await apply([operation])
    const image = (state.sheets[0]!.features!.visuals as any[])[0]
    expect(image.mediaDataUrl.length).toBe(
      operation.base64.length + 'data:image/png;base64,'.length,
    )
    await apply([{ op: 'edit_visual', visualId: image.id, rotation: 90, flipH: true }])
    expect((state.sheets[0]!.features!.visuals as any[])[0]).toMatchObject({
      rotation: 90,
      flipH: true,
    })
    await apply([{ op: 'delete_visual', visualId: image.id }])
    expect(state.sheets[0]!.features!.visuals).toEqual([])
  } finally {
    bridge.dispose()
  }
})
