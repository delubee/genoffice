import { fileURLToPath, pathToFileURL } from 'node:url'
import { join } from 'node:path'
import { randomUUID } from 'node:crypto'
import { XlsxSidecarClient } from '@genoffice/workbook-session/sidecar/client'
import { SidecarWorkbookSource } from '@genoffice/workbook-session/sidecar/source'
import { openWorkbookSnapshot } from '@genoffice/workbook-session/open'
import { initialState } from '@genoffice/workbook-session/state'
import { FileResourceStore } from '@genoffice/workbook-session/store'
import { WorkbookTransactions } from '@genoffice/workbook-session/transactions'
import { prepareCandidateMaterial } from '@genoffice/workbook-session/candidate-material'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import { prepareImportCandidate } from '@genoffice/workbook-session/exchange/import-candidate'
import { WorkbookSnapshots } from '@genoffice/workbook-session/exchange/snapshots'
import { WorkbookImports } from '@genoffice/workbook-session/exchange/import-plan'
import { readCurrentRange } from '@genoffice/workbook-session/read'
import { captureFileBaseline } from '@genoffice/workbook-session/publish'
import { saveWorkbookDocument, type WorkbookSaveDocument } from '../../src/extension/save'
import type * as vscode from 'vscode'

/** Actual Node authority and Rust source; no UI latency is measured by this fixture. */
export async function nativeSession(path: string, storageDirectory: string) {
  const client = new XlsxSidecarClient(
      fileURLToPath(
        new URL('../../../sheets/native/xlsx-engine/target/release/xlsx-sidecar', import.meta.url),
      ),
    ),
    source = new SidecarWorkbookSource(client)
  const opened = await openWorkbookSnapshot({
    uri: pathToFileURL(path).href,
    storageDirectory,
    source,
  })
  const engine = new WorkbookTransactions({
    sessionId: randomUUID(),
    documentUri: pathToFileURL(path).href,
    state: initialState(opened.metadata),
    store: await FileResourceStore.create(join(storageDirectory, 'session')),
    interpreter: {
      prepare: async (state, intent, signal) =>
        computeCandidateState(
          state,
          intent,
          await prepareCandidateMaterial({
            state,
            operations: intent.operations,
            source,
            sourceSessionId: opened.original.sidecarSessionId,
            signal,
          }),
          { signal },
        ),
      prepareImport: (state, intent, batch, signal) =>
        prepareImportCandidate({
          state,
          intent,
          batch,
          source,
          sourceSessionId: opened.original.sidecarSessionId,
          signal,
        }),
      dispose() {},
    },
    host: { committed() {}, projectionChanged() {}, conflict() {} },
  })
  const snapshots = new WorkbookSnapshots(),
    imports = new WorkbookImports(snapshots)
  const view = () => {
    const state = engine.state
    return {
      state,
      descriptor: engine.descriptor,
      read: (sheetId: string, range: string, signal?: AbortSignal) =>
        readCurrentRange({
          state,
          logicalSheetId: sheetId,
          range,
          source,
          sessionId: opened.original.sidecarSessionId,
          signal,
        }),
    }
  }
  const document: WorkbookSaveDocument = {
    engine,
    sessionId: engine.descriptor.sessionId,
    isClosed: false,
    original: opened.original,
    metadata: opened.metadata,
    targetPath: path,
    baselineHash: opened.original.sha256,
    fileBaseline: await captureFileBaseline(path),
    bridge: undefined,
    pauseFileChecks: () => () => {},
  }
  return {
    engine,
    source,
    client,
    opened,
    snapshots,
    imports,
    view,
    async save() {
      await saveWorkbookDocument(document, { scheme: 'file', path, fsPath: path }, client, {
        isCancellationRequested: false,
        onCancellationRequested() {
          return { dispose() {} }
        },
      } as unknown as vscode.CancellationToken)
    },
    async waitForIndex(
      sessionId = opened.original.sidecarSessionId,
      originalSheetIds = engine.state.sheets.flatMap((sheet) =>
        sheet.originalSheetId ? [sheet.originalSheetId] : [],
      ),
    ) {
      const deadline = Date.now() + 30_000
      for (const id of originalSheetIds)
        while (
          !((await source.readFeatures(sessionId, id)) as unknown as { indexingComplete: boolean })
            .indexingComplete
        ) {
          if (Date.now() > deadline) throw new Error('Native fixture indexing timed out')
          await new Promise((resolve) => setTimeout(resolve, 10))
        }
    },
    async dispose() {
      await imports.dispose()
      await engine.dispose()
      await opened.dispose()
      client.stop()
    },
  }
}
