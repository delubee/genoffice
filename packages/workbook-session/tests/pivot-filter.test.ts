import { expect, it } from 'vitest'
import { readFile, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import JSZip from 'jszip'
import { blankXlsxBuffer } from '@genoffice/xlsx-gateway/gateway/csv-import'
import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { XlsxSidecarClient } from '../src/sidecar/client'
import { SidecarWorkbookSource } from '../src/sidecar/source'
import { openWorkbookSnapshot } from '../src/open'
import { initialState } from '../src/state'
import { prepareCandidateMaterial } from '../src/candidate-material'
import { SessionPackageWriter } from '../src/package-writer'
import { useSessionTestResources } from './helpers'
const resources = useSessionTestResources()

const rows: [string, string, number][] = [
  ['North', 'Apples', 12.5],
  ['South', 'Pears', 7],
  ['North', 'Pears', 3],
  ['East', 'Apples', 10],
]

async function salesWorkbook() {
  const header = ['Region', 'Fruit', 'Amount'].map((value, column) => ({
    sheetName: 'Data',
    row: 0,
    column,
    writeValue: true as const,
    cell: { value },
  }))
  const body = rows.flatMap((line, row) =>
    line.map((value, column) => ({
      sheetName: 'Data',
      row: row + 1,
      column,
      writeValue: true as const,
      cell: { value },
    })),
  )
  const mutation = await applyCellEditsToXlsx(await blankXlsxBuffer('Data'), [...header, ...body])
  const zip = await JSZip.loadAsync(mutation.buffer)
  zip.file('customXml/pivot-sentinel.xml', '<preserve>unchanged</preserve>')
  return zip.generateAsync({ type: 'nodebuffer' })
}

function intent(operations: WorkbookOperation[]) {
  return {
    schemaVersion: 1 as const,
    sessionId: 'pivot',
    transactionId: crypto.randomUUID(),
    expectedRevision: 0,
    origin: 'manual' as const,
    summary: 'Create PivotTable',
    operations,
  }
}

async function material(
  state: Parameters<typeof prepareCandidateMaterial>[0]['state'],
  operations: WorkbookOperation[],
  source: SidecarWorkbookSource,
  sessionId: string,
) {
  return prepareCandidateMaterial({ state, operations, source, sourceSessionId: sessionId })
}

it('filters a file PivotTable through a slicer selection, shrinks the output and clears it again', async () => {
  const directory = await resources.directory(),
    path = join(directory, 'sales.xlsx')
  await writeFile(path, await salesWorkbook())
  const client = new XlsxSidecarClient(
    fileURLToPath(
      new URL(
        '../../../apps/sheets/native/xlsx-engine/target/release/xlsx-sidecar' +
          (process.platform === 'win32' ? '.exe' : ''),
        import.meta.url,
      ),
    ),
  )
  try {
    const source = new SidecarWorkbookSource(client)
    const opened = await openWorkbookSnapshot({
      uri: pathToFileURL(path).href,
      storageDirectory: directory,
      source,
    })
    const pivotFile = join(directory, 'with-pivot.xlsx')
    try {
      const state = initialState(opened.metadata)
      const create: WorkbookOperation[] = [
        {
          op: 'add_pivot',
          sheetId: state.sheets[0]!.id,
          sourceRange: 'A1:C5',
          targetCell: 'E2',
          rowFields: 'Region',
          values: [{ field: 'Amount', agg: 'sum' }],
        },
      ]
      const created = await computeCandidateState(
        state,
        intent(create),
        await material(state, create, source, opened.original.sidecarSessionId),
      )
      const staged = await new SessionPackageWriter(client, opened.metadata).stage(
        opened.original.path,
        created.nextState,
        join(directory, 'stage-1.xlsx'),
      )
      await writeFile(pivotFile, await readFile(staged.stagedPath))
    } finally {
      await opened.dispose()
    }
    const reopened = await openWorkbookSnapshot({
      uri: pathToFileURL(pivotFile).href,
      storageDirectory: directory,
      source,
    })
    try {
      const state = initialState(reopened.metadata),
        sheetId = state.sheets[0]!.id
      const meta = state.sheets[0]!.features?.originalMetadata as {
        pivotTables: { path: string }[]
        pivotDefinitions: {
          definition: {
            fields: { sharedItems: unknown[] }[]
            fieldItems: { x: number | null }[][]
          }
        }[]
      }
      const pivot = meta.pivotTables[0]!.path
      const definition = meta.pivotDefinitions[0]!.definition
      const north = definition.fieldItems[0]!.findIndex(
        (item) => item.x !== null && definition.fields[0]!.sharedItems[item.x] === 'North',
      )
      const only: WorkbookOperation[] = [
        { op: 'filter_pivot', sheetId, pivot, field: 0, selected: [north] },
      ]
      const filtered = await computeCandidateState(
        state,
        intent(only),
        await material(state, only, source, reopened.original.sidecarSessionId),
      )
      const cells = filtered.nextState.sheets[0]!.cells
      expect(cells.E3!.rawValue).toBe('North')
      expect(cells.F3!.rawValue).toBe(15.5)
      expect(cells.E4!.rawValue).toBe('Grand Total')
      expect(cells.F4!.rawValue).toBe(15.5)
      expect(cells.E5?.rawValue ?? null).toBe(null)
      expect(cells.F6?.rawValue ?? null).toBe(null)
      const features = filtered.nextState.sheets[0]!.features as {
        pivotRefresh: { outputRefs: Record<string, string> }
        pivotDefinitionEdits: Record<string, { outputRef: string }>
      }
      expect(Object.values(features.pivotRefresh.outputRefs)).toEqual(['E2:F4'])
      expect(features.pivotDefinitionEdits[pivot]!.outputRef).toBe('E2:F4')
      // An unchanged selection is a no-op.
      const same = await computeCandidateState(
        filtered.nextState,
        intent(only),
        await material(filtered.nextState, only, source, reopened.original.sidecarSessionId),
      )
      expect(same.nextState.sheets[0]!.cells).toEqual(cells)
      const staged = await new SessionPackageWriter(client, reopened.metadata).stage(
        reopened.original.path,
        filtered.nextState,
        join(directory, 'stage-2.xlsx'),
      )
      const saved = await JSZip.loadAsync(await readFile(staged.stagedPath))
      expect(await saved.file(pivot)!.async('string')).toMatch(/<location\b[^>]*ref="E2:F4"/)
      // Clearing the slicer restores every member.
      const clear: WorkbookOperation[] = [
        { op: 'filter_pivot', sheetId, pivot, field: 0, selected: null },
      ]
      const cleared = await computeCandidateState(
        filtered.nextState,
        intent(clear),
        await material(filtered.nextState, clear, source, reopened.original.sidecarSessionId),
      )
      const restored = cleared.nextState.sheets[0]!.cells
      expect(restored.E6!.rawValue).toBe('Grand Total')
      expect(restored.F6!.rawValue).toBe(32.5)
      await expect(
        computeCandidateState(
          state,
          intent([{ op: 'filter_pivot', sheetId, pivot, field: 2, selected: [0] }]),
          await material(
            state,
            [{ op: 'filter_pivot', sheetId, pivot, field: 2, selected: [0] }],
            source,
            reopened.original.sidecarSessionId,
          ),
        ),
      ).rejects.toThrow('slicer')
    } finally {
      await reopened.dispose()
    }
  } finally {
    client.stop()
  }
})
