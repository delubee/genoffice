import { writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { initialState } from '../src/state'
import { FileResourceStore } from '../src/store'
import { WorkbookTransactions } from '../src/transactions'
import { readCurrentRange } from '../src/read'
import { prepareCandidateMaterial } from '../src/candidate-material'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import { WorkbookSnapshots } from '../src/exchange/snapshots'
import { WorkbookImports } from '../src/exchange/import-plan'
import { prepareImportCandidate } from '../src/exchange/import-candidate'
import { writeTypedNdjson } from '../src/exchange/ndjson'
import type { TypedCell } from '../src/exchange/schemas'
import type { WorkbookSource } from '../src/ports'

export async function importFixture(
  directory: string,
  options: {
    cellCount?: number
    columns?: number
    startCell?: string
    mode?: 'values' | 'formulas'
    cell?(index: number): TypedCell
    beforeCommit?(): void
    committed?(): void
  } = {},
) {
  const source: WorkbookSource = {
    async openSnapshot() {
      throw new Error('Not used')
    },
    async readBaseRange() {
      return { cells: [] }
    },
    async readFeatures() {
      return { indexingComplete: true, sheetProtection: null, styles: [] }
    },
    async archiveManifest() {
      return {}
    },
    async close() {},
  }
  const state = structuredClone(
      initialState({ sheets: [{ id: 'base', name: 'Data', rowCount: 1, columnCount: 1 }] }),
    ),
    sheetId = state.sheets[0]!.id
  state.sheets[0]!.cells.A1 = { kind: 'number', rawValue: 120 }
  const store = await FileResourceStore.create(join(directory, 'private', 'session'))
  const publish = store.publishManifest.bind(store)
  store.publishManifest = async (name, data) => {
    if (name === 'current') options.beforeCommit?.()
    await publish(name, data)
  }
  const engine = new WorkbookTransactions({
    sessionId: 'import-session',
    documentUri: 'file:///fixture.xlsx',
    state,
    store,
    interpreter: {
      prepare: async (before, intent, signal) =>
        computeCandidateState(
          before,
          intent,
          await prepareCandidateMaterial({
            state: before,
            operations: intent.operations,
            source,
            sourceSessionId: 'original',
            signal,
          }),
          { signal },
        ),
      prepareImport: (before, intent, batch, signal) =>
        prepareImportCandidate({
          state: before,
          intent,
          batch,
          source,
          sourceSessionId: 'original',
          signal,
        }),
      dispose() {},
    },
    host: {
      committed() {
        options.committed?.()
      },
      projectionChanged() {},
      conflict() {},
    },
  })
  const view = () => {
    const captured = engine.state
    return {
      state: captured,
      descriptor: engine.descriptor,
      read: (id: string, range: string, signal?: AbortSignal) =>
        readCurrentRange({
          state: captured,
          logicalSheetId: id,
          range,
          source,
          sessionId: 'original',
          signal,
        }),
    }
  }
  const snapshots = new WorkbookSnapshots(),
    imports = new WorkbookImports(snapshots)
  const snapshot = await snapshots.export({
    view: view(),
    sheetId,
    range: 'A1',
    workspaceRoot: directory,
  })
  const cellCount = options.cellCount ?? 2502,
    columns = options.columns ?? 6,
    rows = cellCount / columns
  async function* records() {
    for (let row = 0; row < rows; row++)
      yield {
        rowOffset: row,
        columnOffset: 0,
        cells: Array.from(
          { length: columns },
          (_, column) =>
            options.cell?.(row * columns + column) ?? {
              kind: 'number' as const,
              value: row * columns + column,
            },
        ),
      }
  }
  const shape = { rows, columns, cellCount },
    output = await writeTypedNdjson({
      path: join(snapshot.directory, 'result.ndjson'),
      shape,
      records: records(),
    })
  const manifest = {
    schemaVersion: 1,
    kind: 'workbook-result',
    source: {
      snapshotId: snapshot.manifest.snapshotId,
      sourceManifestSha256: snapshot.manifestSha256,
      sourceDataSha256: snapshot.manifest.data.sha256,
      sourceRevision: 0,
    },
    shape,
    output,
    target: { newSheetName: 'Summary', startCell: options.startCell ?? 'B2' },
    mode: options.mode ?? 'values',
    complete: true,
  }
  const resultManifestPath = join(snapshot.directory, 'result.manifest.json')
  await writeFile(resultManifestPath, JSON.stringify(manifest))
  return {
    engine,
    sheetId,
    imports,
    snapshots,
    snapshot,
    resultManifestPath,
    prepare: () =>
      imports.prepare({
        snapshotId: snapshot.manifest.snapshotId,
        bindingId: 'task-binding',
        resultManifestPath,
        privateStorage: join(directory, 'private'),
        view: view(),
        expectedRevision: 0,
        currentDescriptor: () => engine.descriptor,
      }),
    async dispose() {
      await imports.dispose()
      await engine.dispose()
    },
  }
}
