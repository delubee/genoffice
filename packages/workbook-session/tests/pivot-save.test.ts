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

it('bakes a session PivotTable, saves native pivot parts and reopens it as a file pivot', async () => {
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
    try {
      const state = initialState(opened.metadata),
        sheetId = state.sheets[0]!.id
      const operations: WorkbookOperation[] = [
        {
          op: 'add_pivot',
          sheetId,
          sourceRange: 'A1:C5',
          targetCell: 'E2',
          rowFields: 'Region',
          values: [{ field: 'Amount', agg: 'sum' }],
          subtotalLabel: '小计',
        },
      ]
      const material = await prepareCandidateMaterial({
        state,
        operations,
        source,
        sourceSessionId: opened.original.sidecarSessionId,
      })
      const result = await computeCandidateState(state, intent(operations), material)
      expect(result.inverse.state).toEqual(state)
      const cells = result.nextState.sheets[0]!.cells
      expect(cells.E2!.rawValue).toBe('Region')
      expect(cells.F2!.rawValue).toBe('Sum of Amount')
      const labels = [cells.E3, cells.E4, cells.E5].map((cell) => cell!.rawValue)
      expect(labels.sort()).toEqual(['East', 'North', 'South'])
      expect(cells.E6!.rawValue).toBe('Grand Total')
      expect(cells.F6!.rawValue).toBe(32.5)
      expect(cells.A2!.rawValue).toBe('North')
      const [pending] = (result.nextState.sheets[0]!.features?.pivotAdditions ?? []) as {
        name: string
        location: object
      }[]
      expect(pending).toMatchObject({
        name: 'Pivot1',
        location: { startRow: 1, startColumn: 4, endRow: 5, endColumn: 5 },
      })

      // Coordinates are pinned until the pivot is saved.
      const structural: WorkbookOperation[] = [{ op: 'insert_rows', sheetId, row: 1, count: 1 }]
      await expect(
        computeCandidateState(
          result.nextState,
          intent(structural),
          await prepareCandidateMaterial({
            state: result.nextState,
            operations: structural,
            source,
            sourceSessionId: opened.original.sidecarSessionId,
          }),
        ),
      ).rejects.toThrow('session_pivot')
      // Output overlapping the pending pivot is refused.
      const overlap: WorkbookOperation[] = [
        { ...operations[0]!, targetCell: 'F3' } as WorkbookOperation,
      ]
      await expect(
        computeCandidateState(
          result.nextState,
          intent(overlap),
          await prepareCandidateMaterial({
            state: result.nextState,
            operations: overlap,
            source,
            sourceSessionId: opened.original.sidecarSessionId,
          }),
        ),
      ).rejects.toThrow('pivot_layout:overlapSession')

      // Edit PivotTable: Fruit on rows shrinks the block; the old last row clears.
      const edit: WorkbookOperation[] = [
        {
          op: 'edit_pivot',
          sheetId,
          pivot: { session: 'Pivot1' },
          rowFields: 'Fruit',
          values: [{ field: 'Amount', agg: 'sum' }],
        },
      ]
      const edited = await computeCandidateState(
        result.nextState,
        intent(edit),
        await prepareCandidateMaterial({
          state: result.nextState,
          operations: edit,
          source,
          sourceSessionId: opened.original.sidecarSessionId,
        }),
      )
      const editedCells = edited.nextState.sheets[0]!.cells
      expect(editedCells.E2!.rawValue).toBe('Fruit')
      expect(editedCells.E5!.rawValue).toBe('Grand Total')
      expect(editedCells.F5!.rawValue).toBe(32.5)
      expect(editedCells.E6?.rawValue ?? null).toBe(null)
      expect(editedCells.F6?.rawValue ?? null).toBe(null)
      expect(edited.nextState.sheets[0]!.features?.pivotAdditions).toEqual([
        expect.objectContaining({
          name: 'Pivot1',
          location: { startRow: 1, startColumn: 4, endRow: 4, endColumn: 5 },
        }),
      ])
      const staged = await new SessionPackageWriter(client, opened.metadata).stage(
        opened.original.path,
        edited.nextState,
        join(directory, 'saved.xlsx'),
      )
      const saved = await JSZip.loadAsync(await readFile(staged.stagedPath))
      const parts = Object.keys(saved.files)
      expect(parts.some((name) => /^xl\/pivotTables\/pivotTable\d+\.xml$/.test(name))).toBe(true)
      expect(
        parts.some((name) => /^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/.test(name)),
      ).toBe(true)
      expect(await saved.file('customXml/pivot-sentinel.xml')!.async('string')).toBe(
        '<preserve>unchanged</preserve>',
      )
      const reopened = await source.openSnapshot(staged.stagedPath)
      try {
        const sheet = (reopened.metadata as { sheets: { pivotRanges?: unknown[] }[] }).sheets[0]!
        expect(sheet.pivotRanges).toEqual([
          { startRow: 1, startColumn: 4, endRow: 4, endColumn: 5 },
        ])
        const read = (await source.readBaseRange(
          reopened.sessionId,
          (reopened.metadata as { sheets: { id: string }[] }).sheets[0]!.id,
          'F5:F5',
        )) as { cells: { value: unknown }[] }
        expect(read.cells[0]!.value).toBe(32.5)
      } finally {
        await source.close(reopened.sessionId)
      }
    } finally {
      await opened.dispose()
    }
  } finally {
    client.stop()
  }
})
