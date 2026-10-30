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
import { useSessionTestResources } from './helpers'
const resources = useSessionTestResources()

const rows: [string, string, number][] = [
  ['North', 'Apples', 12.5],
  ['South', 'Pears', 7],
  ['North', 'Pears', 3],
  ['East', 'Apples', 10],
]

async function salesWorkbook() {
  const cells = [
    ['Region', 'Amount'],
    ['North', 120],
    ['South', -30],
  ].flatMap((line, row) =>
    line.map((value, column) => ({
      sheetName: 'Data',
      row,
      column,
      writeValue: true as const,
      cell: { value },
    })),
  )
  return (await applyCellEditsToXlsx(await blankXlsxBuffer('Data'), cells)).buffer
}

function intent(operations: WorkbookOperation[]) {
  return {
    schemaVersion: 1 as const,
    sessionId: 'cf',
    transactionId: crypto.randomUUID(),
    expectedRevision: 0,
    origin: 'manual' as const,
    summary: 'Conditional formatting',
    operations,
  }
}

it('saves a session conditional-format rule set as native rules and reads it back as file rules', async () => {
  const directory = await resources.directory(),
    path = join(directory, 'cf.xlsx')
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
    const saved = join(directory, 'saved.xlsx')
    try {
      const state = initialState(opened.metadata),
        sheetId = state.sheets[0]!.id
      const rules = [
        {
          ranges: [{ startRow: 1, endRow: 2, startColumn: 1, endColumn: 1 }],
          stopIfTrue: false,
          rule: {
            type: 'highlightCell',
            subType: 'number',
            operator: 'greaterThan',
            value: 100,
            style: { bg: { rgb: '#FFC7CE' } },
          },
        },
      ]
      const operations: WorkbookOperation[] = [{ op: 'set_conditional_formats', sheetId, rules }]
      const next = await computeCandidateState(
        state,
        intent(operations),
        await prepareCandidateMaterial({
          state,
          operations,
          source,
          sourceSessionId: opened.original.sidecarSessionId,
        }),
      )
      expect(next.nextState.sheets[0]!.features?.cfEdits).toEqual({ rules })
      // Coordinates stay fixed while session rules are pending.
      const insert: WorkbookOperation[] = [{ op: 'insert_rows', sheetId, row: 1, count: 1 }]
      await expect(
        computeCandidateState(
          next.nextState,
          intent(insert),
          await prepareCandidateMaterial({
            state: next.nextState,
            operations: insert,
            source,
            sourceSessionId: opened.original.sidecarSessionId,
          }),
        ),
      ).rejects.toThrow('session_conditional_format')
      // Unsaveable shapes are refused at edit time, not at save.
      const bad: WorkbookOperation[] = [
        {
          op: 'set_conditional_formats',
          sheetId,
          rules: [{ ...rules[0]!, rule: { type: 'mystery' } }],
        },
      ]
      await expect(
        computeCandidateState(
          state,
          intent(bad),
          await prepareCandidateMaterial({
            state,
            operations: bad,
            source,
            sourceSessionId: opened.original.sidecarSessionId,
          }),
        ),
      ).rejects.toThrow('unsupported_conditional_format')
      const staged = await new SessionPackageWriter(client, opened.metadata).stage(
        opened.original.path,
        next.nextState,
        join(directory, 'stage.xlsx'),
      )
      const zip = await JSZip.loadAsync(await readFile(staged.stagedPath))
      const sheetXml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
      expect(sheetXml).toMatch(/<conditionalFormatting sqref="B2:B3">/)
      expect(sheetXml).toMatch(/operator="greaterThan"/)
      await writeFile(saved, await readFile(staged.stagedPath))
    } finally {
      await opened.dispose()
    }
    const reopened = await openWorkbookSnapshot({
      uri: pathToFileURL(saved).href,
      storageDirectory: directory,
      source,
    })
    try {
      const state = initialState(reopened.metadata)
      const read = (await readCurrentRange({
        state,
        source,
        sessionId: reopened.original.sidecarSessionId,
        logicalSheetId: state.sheets[0]!.id,
        range: 'A1:B3',
        includeMerges: true,
      })) as { conditionalRules: { ruleType: string; ranges: unknown[] }[] }
      expect(read.conditionalRules).toHaveLength(1)
      expect(read.conditionalRules[0]!.ranges).toEqual([
        { startRow: 1, endRow: 2, startColumn: 1, endColumn: 1 },
      ])
    } finally {
      await reopened.dispose()
    }
  } finally {
    client.stop()
  }
})
