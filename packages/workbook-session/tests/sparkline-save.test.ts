import { expect, it } from 'vitest'
import { writeFile } from 'node:fs/promises'
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

it('saves session sparklines through the native writer and reopens them', async () => {
  const directory = await resources.directory()
  const fixture = await applyCellEditsToXlsx(
    await blankXlsxBuffer('Data'),
    [1, 2].flatMap((row) =>
      [3, 5, 4].map((value, column) => ({
        sheetName: 'Data',
        row,
        column,
        writeValue: true as const,
        cell: { value: value * row },
      })),
    ),
  )
  const path = join(directory, 'sparks.xlsx')
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
    const opened = await openWorkbookSnapshot({
      uri: pathToFileURL(path).href,
      storageDirectory: directory,
      source,
    })
    try {
      const state: WorkbookState = initialState(opened.metadata)
      const ops: WorkbookOperation[] = [
        { op: 'add_sparkline', sheetId: state.sheets[0]!.id, type: 'column', dataRange: 'A2:C3' },
      ]
      const result = await computeCandidateState(
        state,
        {
          schemaVersion: 1,
          sessionId: 'sparks',
          transactionId: crypto.randomUUID(),
          expectedRevision: 0,
          origin: 'manual',
          summary: 'Sparklines',
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
        join(directory, 'stage.xlsx'),
      )
      const reopened = await source.openSnapshot(staged.stagedPath)
      try {
        const sheet = (
          reopened.metadata as {
            sheets: {
              sparklines: { type: string; cells: { cell: string; sourceRef: string }[] }[]
            }[]
          }
        ).sheets[0]!
        expect(sheet.sparklines).toEqual([
          expect.objectContaining({
            type: 'column',
            cells: [
              { cell: 'D2', sourceRef: expect.stringMatching(/Data'?!\$A\$2:\$C\$2$/) },
              { cell: 'D3', sourceRef: expect.stringMatching(/Data'?!\$A\$3:\$C\$3$/) },
            ],
          }),
        ])
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
