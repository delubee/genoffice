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

it('refreshes a file PivotTable in place and grows it for new members, saving refreshOnLoad and the new location', async () => {
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
    // A file pivot: create in one session and save.
    const opened = await openWorkbookSnapshot({
      uri: pathToFileURL(path).href,
      storageDirectory: directory,
      source,
    })
    const pivotPath = join(directory, 'with-pivot.xlsx')
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
      await writeFile(pivotPath, await readFile(staged.stagedPath))
    } finally {
      await opened.dispose()
    }
    const reopened = await openWorkbookSnapshot({
      uri: pathToFileURL(pivotPath).href,
      storageDirectory: directory,
      source,
    })
    try {
      const state = initialState(reopened.metadata),
        sheetId = state.sheets[0]!.id
      const meta = state.sheets[0]!.features?.originalMetadata as {
        pivotDefinitions?: { definition?: { outputRef: string } }[]
      }
      expect(meta.pivotDefinitions?.[0]?.definition?.outputRef).toBe('E2:F6')
      // Edit a source value: the data area recomputes in place.
      const edit: WorkbookOperation[] = [{ op: 'set_cell', sheetId, address: 'C2', value: 20 }]
      const edited = await computeCandidateState(
        state,
        intent(edit),
        await material(state, edit, source, reopened.original.sidecarSessionId),
      )
      const refresh: WorkbookOperation[] = [{ op: 'refresh_pivot', sheetId }]
      const first = await computeCandidateState(
        edited.nextState,
        intent(refresh),
        await material(edited.nextState, refresh, source, reopened.original.sidecarSessionId),
      )
      const cells = first.nextState.sheets[0]!.cells
      const north = ['E3', 'E4', 'E5'].find((address) => cells[address]?.rawValue === 'North')!
      expect(cells[`F${north.slice(1)}`]!.rawValue).toBe(23)
      expect(cells.F6!.rawValue).toBe(40)
      // A new member inside the source range grows the layout by one row (East stays an item).
      const grow: WorkbookOperation[] = [{ op: 'set_cell', sheetId, address: 'A5', value: 'West' }]
      const added = await computeCandidateState(
        first.nextState,
        intent(grow),
        await material(first.nextState, grow, source, reopened.original.sidecarSessionId),
      )
      // Growth is refused while the new row would overwrite content.
      const blocked: WorkbookOperation[] = [
        { op: 'set_cell', sheetId, address: 'E7', value: 'note' },
      ]
      const occupied = await computeCandidateState(
        added.nextState,
        intent(blocked),
        await material(added.nextState, blocked, source, reopened.original.sidecarSessionId),
      )
      await expect(
        computeCandidateState(
          occupied.nextState,
          intent(refresh),
          await material(occupied.nextState, refresh, source, reopened.original.sidecarSessionId),
        ),
      ).rejects.toThrow('pivot_layout:growConflict')
      const second = await computeCandidateState(
        added.nextState,
        intent(refresh),
        await material(added.nextState, refresh, source, reopened.original.sidecarSessionId),
      )
      const grown = second.nextState.sheets[0]!.cells
      expect(grown.E7!.rawValue).toBe('Grand Total')
      expect(grown.F7!.rawValue).toBe(40)
      expect(['E3', 'E4', 'E5', 'E6'].map((address) => grown[address]!.rawValue)).toContain('West')
      const refreshed = second.nextState.sheets[0]!.features?.pivotRefresh as {
        cachePaths: string[]
        outputRefs: Record<string, string>
      }
      expect(refreshed.cachePaths).toHaveLength(1)
      expect(Object.values(refreshed.outputRefs)).toEqual(['E2:F7'])
      const staged = await new SessionPackageWriter(client, reopened.metadata).stage(
        reopened.original.path,
        second.nextState,
        join(directory, 'stage-2.xlsx'),
      )
      const saved = await JSZip.loadAsync(await readFile(staged.stagedPath))
      const cache = Object.keys(saved.files).find((name) =>
        /^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/.test(name),
      )!
      expect(await saved.file(cache)!.async('string')).toContain('refreshOnLoad="1"')
      const table = Object.keys(saved.files).find((name) =>
        /^xl\/pivotTables\/pivotTable\d+\.xml$/.test(name),
      )!
      expect(await saved.file(table)!.async('string')).toMatch(/<location\b[^>]*ref="E2:F7"/)
      const final = await source.openSnapshot(staged.stagedPath)
      try {
        const sheet = (final.metadata as { sheets: { id: string; pivotRanges?: unknown[] }[] })
          .sheets[0]!
        expect(sheet.pivotRanges).toEqual([
          { startRow: 1, startColumn: 4, endRow: 6, endColumn: 5 },
        ])
        const read = (await source.readBaseRange(final.sessionId, sheet.id, 'F7:F7')) as {
          cells: { value: unknown }[]
        }
        expect(read.cells[0]!.value).toBe(40)
      } finally {
        await source.close(final.sessionId)
      }
    } finally {
      await reopened.dispose()
    }
  } finally {
    client.stop()
  }
})
