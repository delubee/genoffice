import { expect, it } from 'vitest'
import { readFile, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { blankXlsxBuffer } from '@genoffice/xlsx-gateway/gateway/csv-import'
import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { XlsxSidecarClient } from '../src/sidecar/client'
import { SidecarWorkbookSource } from '../src/sidecar/source'
import { openWorkbookSnapshot } from '../src/open'
import { initialState } from '../src/state'
import { prepareCandidateMaterial } from '../src/candidate-material'
import { SessionPackageWriter } from '../src/package-writer'
import type { WorkbookState } from '../src/domain'
import { useSessionTestResources } from './helpers'
const resources = useSessionTestResources()

type Table = { name?: string; range: object; styleName?: string; showColumnStripes: boolean }

it('creates, restyles, renames and converts a table through the session and the native writer', async () => {
  const directory = await resources.directory()
  const rows = [
    ['Region', 'Amount'],
    ['North', 12.5],
    ['South', 7],
  ]
  const fixture = await applyCellEditsToXlsx(
    await blankXlsxBuffer('Data'),
    rows.flatMap((line, row) =>
      line.map((value, column) => ({
        sheetName: 'Data',
        row: row + 1,
        column: column + 1,
        writeValue: true as const,
        cell: { value },
      })),
    ),
  )
  let path = join(directory, 'tables.xlsx')
  await writeFile(path, fixture.buffer)
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
    const cycle = async (operations: (sheetId: string) => WorkbookOperation[]) => {
      const opened = await openWorkbookSnapshot({
        uri: pathToFileURL(path).href,
        storageDirectory: directory,
        source,
      })
      try {
        const state: WorkbookState = initialState(opened.metadata)
        const ops = operations(state.sheets[0]!.id)
        const result = await computeCandidateState(
          state,
          {
            schemaVersion: 1,
            sessionId: 'tables',
            transactionId: crypto.randomUUID(),
            expectedRevision: 0,
            origin: 'manual',
            summary: 'Tables',
            operations: ops,
          },
          await prepareCandidateMaterial({
            state,
            operations: ops,
            source,
            sourceSessionId: opened.original.sidecarSessionId,
          }),
        )
        const staged = await new SessionPackageWriter(client, opened.metadata).stage(
          opened.original.path,
          result.nextState,
          join(directory, `stage-${crypto.randomUUID()}.xlsx`),
        )
        path = join(directory, `saved-${crypto.randomUUID()}.xlsx`)
        await writeFile(path, await readFile(staged.stagedPath))
        const reopened = await source.openSnapshot(path)
        try {
          return (reopened.metadata as { sheets: { tables: Table[] }[] }).sheets[0]!.tables
        } finally {
          await source.close(reopened.sessionId)
        }
      } finally {
        await opened.dispose()
      }
    }
    let tables = await cycle((sheetId) => [
      { op: 'add_table', sheetId, range: 'B2:C4', style: 'TableStyleMedium9' },
    ])
    expect(tables).toEqual([
      expect.objectContaining({
        name: 'Table1',
        range: { startRow: 1, startColumn: 1, endRow: 3, endColumn: 2 },
        styleName: 'TableStyleMedium9',
      }),
    ])
    tables = await cycle((sheetId) => [
      {
        op: 'edit_table',
        sheetId,
        table: 'Table1',
        name: 'Sales',
        style: 'TableStyleLight1',
        bandedColumns: true,
      },
    ])
    expect(tables[0]).toMatchObject({
      name: 'Sales',
      styleName: 'TableStyleLight1',
      showColumnStripes: true,
    })
    tables = await cycle((sheetId) => [{ op: 'edit_table', sheetId, table: 'Sales', remove: true }])
    expect(tables).toEqual([])
  } finally {
    client.stop()
  }
})
