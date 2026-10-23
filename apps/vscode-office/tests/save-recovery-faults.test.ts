import { expect, it } from 'vitest'
import { readFile, writeFile } from 'node:fs/promises'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { join } from 'node:path'
import type * as vscode from 'vscode'
import { manualWorkbook } from './fixtures/manual-workbook'
import { withTestWorkspace } from './helpers'
import { XlsxSidecarClient } from '@genoffice/workbook-session/sidecar/client'
import { SidecarWorkbookSource } from '@genoffice/workbook-session/sidecar/source'
import { openWorkbookSnapshot } from '@genoffice/workbook-session/open'
import { initialState } from '@genoffice/workbook-session/state'
import { FileResourceStore } from '@genoffice/workbook-session/store'
import { WorkbookTransactions } from '@genoffice/workbook-session/transactions'
import { prepareCandidateMaterial } from '@genoffice/workbook-session/candidate-material'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import { captureFileBaseline } from '@genoffice/workbook-session/publish'
import { saveWorkbookDocument, type WorkbookSaveDocument } from '../src/extension/save'

const token = (cancelled = false) =>
  ({
    isCancellationRequested: cancelled,
    onCancellationRequested() {
      return { dispose() {} }
    },
  }) as unknown as vscode.CancellationToken

it('retains content/history on save cancellation, pending input and a failed saved-marker publication after an actual XLSX write', async () => {
  await withTestWorkspace(async (directory) => {
    const path = join(directory, 'fault.xlsx'),
      original = await manualWorkbook(),
      destination = { scheme: 'file', path, fsPath: path },
      client = new XlsxSidecarClient(
        fileURLToPath(
          new URL('../../sheets/native/xlsx-engine/target/release/xlsx-sidecar', import.meta.url),
        ),
      ),
      source = new SidecarWorkbookSource(client)
    await writeFile(path, original)
    const opened = await openWorkbookSnapshot({
      uri: pathToFileURL(path).href,
      storageDirectory: directory,
      source,
    })
    let engine: WorkbookTransactions | undefined
    let failRoot = false,
      diskFull = false,
      publishing = ''
    try {
      const store = await FileResourceStore.create(join(directory, 'store'), (point) => {
        if (failRoot && publishing === 'current' && point === 'beforeManifestRename')
          throw new Error('Injected marker publication failure')
      })
      const publish = store.publishManifest.bind(store)
      store.publishManifest = async (name, value) => {
        publishing = name
        await publish(name, value)
      }
      const writeChunk = store.writeChunk.bind(store)
      store.writeChunk = async (bytes) => {
        if (diskFull)
          throw Object.assign(new Error('Injected full backup disk'), { code: 'ENOSPC' })
        return writeChunk(bytes)
      }
      const state = initialState(opened.metadata)
      engine = new WorkbookTransactions({
        sessionId: 'save-fault',
        documentUri: pathToFileURL(path).href,
        state,
        store,
        interpreter: {
          async prepare(before, intent) {
            return computeCandidateState(
              before,
              intent,
              await prepareCandidateMaterial({
                state: before,
                operations: intent.operations,
                source,
                sourceSessionId: opened.original.sidecarSessionId,
              }),
            )
          },
          dispose() {},
        },
        host: { committed() {}, projectionChanged() {}, conflict() {} },
      })
      await engine.submit({
        schemaVersion: 1,
        sessionId: 'save-fault',
        transactionId: 'edit',
        expectedRevision: 0,
        origin: 'manual',
        summary: 'Unsaved content',
        operations: [{ op: 'set_cell', sheetId: state.sheets[0]!.id, address: 'A2', value: 99 }],
      })
      let ready = true
      const document: WorkbookSaveDocument = {
        engine,
        sessionId: 'save-fault',
        isClosed: false,
        original: opened.original,
        metadata: opened.metadata,
        targetPath: path,
        baselineHash: opened.original.sha256,
        fileBaseline: await captureFileBaseline(path),
        bridge: {
          async request() {
            return { ready }
          },
        },
        pauseFileChecks: () => () => {},
      }
      const current = engine.state,
        history = engine.history
      await expect(
        saveWorkbookDocument(document, destination, client, token(true)),
      ).rejects.toMatchObject({ code: 'operation_cancelled' })
      ready = false
      await expect(
        saveWorkbookDocument(document, destination, client, token()),
      ).rejects.toMatchObject({ code: 'editor_busy' })
      expect(await readFile(path)).toEqual(original)
      expect(engine.history).toEqual(history)
      ready = true
      const mode = document.fileBaseline.mode
      document.fileBaseline = { ...document.fileBaseline, mode: 0o444 }
      await expect(
        saveWorkbookDocument(document, destination, client, token()),
      ).rejects.toMatchObject({ code: 'save_failed' })
      document.fileBaseline = { ...document.fileBaseline, mode }
      diskFull = true
      await expect(
        saveWorkbookDocument(document, destination, client, token()),
      ).rejects.toMatchObject({ code: 'ENOSPC' })
      diskFull = false
      expect(await readFile(path)).toEqual(original)
      expect(engine.state).toBe(current)
      expect(engine.history).toEqual(history)
      expect(engine.descriptor.dirty).toBe(true)
      ready = true
      failRoot = true
      await expect(
        saveWorkbookDocument(document, destination, client, token()),
      ).rejects.toMatchObject({ code: 'save_failed', details: { fileWritten: true } })
      expect(await readFile(path)).not.toEqual(original)
      expect(engine.state).toBe(current)
      expect(engine.history).toEqual(history)
      expect(engine.descriptor.dirty).toBe(true)
      failRoot = false
      await engine.undo()
      await saveWorkbookDocument({ ...document, bridge: undefined }, destination, client, token())
      expect(await readFile(path)).toEqual(original)
      expect(engine.descriptor.dirty).toBe(false)
      expect(engine.history).toEqual(history)
    } finally {
      await engine?.dispose()
      await opened.dispose()
      client.stop()
    }
  })
})
