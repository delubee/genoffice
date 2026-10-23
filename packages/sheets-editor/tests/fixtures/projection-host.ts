import { writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import type { Webview } from 'vscode'
import { WebviewCandidateBridge } from '../../../../apps/vscode-office/src/extension/editor-bridge'
import { WorkbookTransactions } from '../../../workbook-session/src/transactions'
import { FileResourceStore } from '../../../workbook-session/src/store'
import { XlsxSidecarClient } from '../../../workbook-session/src/sidecar/client'
import { SidecarWorkbookSource } from '../../../workbook-session/src/sidecar/source'
import { openWorkbookSnapshot } from '../../../workbook-session/src/open'
import { initialState } from '../../../workbook-session/src/state'
import { buildCompatibilityFixture } from '../../../../apps/sheets/tests/fixture-builder'

export async function createProjectionSession(directory: string, binary: string, webview: Webview) {
  const client = new XlsxSidecarClient(binary)
  let opened: Awaited<ReturnType<typeof openWorkbookSnapshot>> | undefined
  try {
    const path = join(directory, 'fixture.xlsx')
    await writeFile(path, await buildCompatibilityFixture())
    const source = new SidecarWorkbookSource(client)
    opened = await openWorkbookSnapshot({
      uri: new URL(`file://${path}`).href,
      storageDirectory: directory,
      source,
    })
    const state = initialState(opened.metadata),
      sessionId = 'projection-session'
    let engine: WorkbookTransactions
    const interpreter = new WebviewCandidateBridge({
      webview,
      sessionId,
      original: opened.original,
      source,
      getState: () => engine.state,
      getDescriptor: () => engine.descriptor,
      submit: (intent) => engine.submit(intent),
      history: async (direction) => {
        historyRequests.push(direction)
        await (direction === 'undo' ? engine.undo() : engine.redo())
      },
    })
    let edits = 0
    const historyRequests: string[] = []
    engine = new WorkbookTransactions({
      sessionId,
      documentUri: opened.original.sourceUri,
      state,
      store: await FileResourceStore.create(join(directory, 'store')),
      interpreter,
      host: {
        committed() {
          edits++
        },
        async projectionChanged() {
          await interpreter.bridge.request('state.committed', {
            descriptor: engine.descriptor,
            state: engine.state,
          })
        },
        conflict() {},
      },
    })
    return {
      engine,
      interpreter,
      state,
      edits: () => edits,
      historyRequests: () => [...historyRequests],
      async dispose() {
        await engine.dispose()
        await opened?.dispose()
        client.stop()
      },
    }
  } catch (error) {
    await opened?.dispose()
    client.stop()
    throw error
  }
}
