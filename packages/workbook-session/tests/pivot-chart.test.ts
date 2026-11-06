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
import { readCurrentRange } from '../src/read'
import { buildPivotChartData } from '@genoffice/xlsx-gateway/domain/pivot-chart'
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

it('inserts a PivotChart over a file PivotTable as add_chart plus a series edit', async () => {
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
        sheet = state.sheets[0]!
      const definition = (
        sheet.features?.originalMetadata as unknown as {
          pivotDefinitions: { definition: Parameters<typeof buildPivotChartData>[0] }[]
        }
      ).pivotDefinitions[0]!.definition
      const read = await readCurrentRange({
        state,
        source,
        sessionId: reopened.original.sidecarSessionId,
        logicalSheetId: sheet.id,
        range: 'E2:F6',
      })
      const cells = new Map(read.cells.map((cell) => [cell.address, cell.rawValue]))
      const values = [2, 3, 4, 5, 6].map((row) =>
        ['E', 'F'].map((column) => cells.get(`${column}${row}`)),
      )
      const chart = buildPivotChartData(
        definition,
        sheet.name,
        { startRow: 1, startColumn: 4, endRow: 5, endColumn: 5 },
        values as never,
      )
      const ops: WorkbookOperation[] = [
        {
          op: 'add_chart',
          sheetId: sheet.id,
          chartType: 'column',
          dataRange: 'E2:F6',
          title: chart.title,
          anchorCell: 'H2',
        },
        {
          op: 'edit_chart',
          chartPath: `demo-chart-${sheet.id}-1`,
          seriesSelection: {
            mode: 'rows',
            rows: chart.series.map((series) => ({
              name: series.name,
              ...(series.valuesRef ? { valuesRef: series.valuesRef } : {}),
              ...(series.categoriesRef ? { categoriesRef: series.categoriesRef } : {}),
            })),
          },
        },
      ]
      const result = await computeCandidateState(
        state,
        intent(ops),
        await material(state, ops, source, reopened.original.sidecarSessionId),
      )
      const visuals = result.nextState.sheets[0]!.features?.visuals as {
        id: string
        chart: { series: { valuesRef?: string; categoriesRef?: string }[] }
      }[]
      expect(visuals).toHaveLength(1)
      expect(visuals[0]!.chart.series[0]!.valuesRef).toMatch(/\$F\$3:\$F\$5$/)
      expect(visuals[0]!.chart.series[0]!.categoriesRef).toMatch(/\$E\$3:\$E\$5$/)
    } finally {
      await reopened.dispose()
    }
  } finally {
    client.stop()
  }
})
