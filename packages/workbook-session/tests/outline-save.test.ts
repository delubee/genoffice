import { readCurrentOutline } from '../src/read-outline'
import { expect, it } from 'vitest'
import { readFile, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import JSZip from 'jszip'
import { manualWorkbook } from '../../../apps/vscode-office/tests/fixtures/manual-workbook'
import { XlsxSidecarClient } from '../src/sidecar/client'
import { SidecarWorkbookSource } from '../src/sidecar/source'
import { openWorkbookSnapshot } from '../src/open'
import { initialState } from '../src/state'
import { prepareCandidateMaterial } from '../src/candidate-material'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import {
  workbookOperationSchema,
  type WorkbookOperation,
} from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { SessionPackageWriter } from '../src/package-writer'
import { useSessionTestResources } from './helpers'
const resources = useSessionTestResources()
it('saves outline levels, collapsed summary and placement through the canonical journal and reopens with source cells intact', async () => {
  const directory = await resources.directory(),
    path = join(directory, 'outline.xlsx')
  await writeFile(path, await manualWorkbook())
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
        { op: 'set_outline', sheetId, axis: 'rows', start: 2, count: 2, level: 1 },
        { op: 'set_outline', sheetId, axis: 'rows', start: 4, count: 1, level: 0, collapsed: true },
        {
          op: 'set_outline',
          sheetId,
          axis: 'cols',
          start: 2,
          count: 2,
          level: 2,
          collapsed: false,
        },
        { op: 'set_outline_settings', sheetId, summaryBelow: false, summaryRight: true },
        { op: 'set_rows_hidden', sheetId, row: 2, count: 2, hidden: true },
        { op: 'set_rows_hidden', sheetId, row: 2, count: 1, hidden: false },
        { op: 'set_cols_hidden', sheetId, column: 'B', count: 2, hidden: true },
      ]
      const material = await prepareCandidateMaterial({
        state,
        operations,
        source,
        sourceSessionId: opened.original.sidecarSessionId,
      })
      const result = await computeCandidateState(
        state,
        {
          schemaVersion: 1,
          sessionId: 'outline',
          transactionId: crypto.randomUUID(),
          expectedRevision: 0,
          origin: 'manual',
          summary: 'Group rows and columns',
          operations,
        },
        material,
      )
      expect(result.inverse.state).toEqual(state)
      expect(result.nextState.sheets[0]!.cells).toEqual(state.sheets[0]!.cells)
      const staged = await new SessionPackageWriter(client, opened.metadata).stage(
        opened.original.path,
        result.nextState,
        join(directory, 'saved.xlsx'),
      )
      const saved = await JSZip.loadAsync(await readFile(staged.stagedPath))
      const xml = await saved.file('xl/worksheets/sheet1.xml')!.async('string')
      expect(xml).toMatch(/<row\b[^>]*r="2"[^>]*outlineLevel="1"/)
      expect(xml).toMatch(/<row\b[^>]*r="4"[^>]*collapsed="1"/)
      expect(xml).toContain('summaryBelow="0"')
      expect(xml).toContain('outlineLevelCol="2"')
      expect(await saved.file('customXml/manual-sentinel.xml')!.async('string')).toBe(
        '<preserve id="manual-fixture">unchanged</preserve>',
      )
      const reopened = await source.openSnapshot(staged.stagedPath)
      try {
        expect(reopened.metadata.sheets[0]).toMatchObject({
          outlineLevelRow: 1,
          outlineLevelCol: 2,
          outlineSummaryBelow: false,
        })
        const reopenedState = initialState(reopened.metadata)
        const outline = await readCurrentOutline({
          state: reopenedState,
          source,
          sessionId: reopened.sessionId,
          sheetId: reopenedState.sheets[0]!.id,
        })
        expect(outline.rows.find((row) => row.index === 1)).toMatchObject({
          level: 1,
          hidden: false,
        })
        expect(outline.rows.find((row) => row.index === 2)).toMatchObject({
          level: 1,
          hidden: true,
        })
        expect(outline.rows.find((row) => row.index === 3)).toMatchObject({
          level: 0,
          collapsed: true,
        })
        expect(outline.cols.find((col) => col.index === 1)).toMatchObject({
          level: 2,
          hidden: true,
        })
        expect(outline.summaryBelow).toBe(false)
        const read = (await source.readBaseRange(
          reopened.sessionId,
          reopened.metadata.sheets[0]!.id,
          'A2:A2',
        )) as { cells: { value: unknown }[] }
        expect(read.cells[0]!.value).toBe(12.5)
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
it('validates outline grid spans and level bounds before authoring', () => {
  const operation = {
    op: 'set_outline',
    sheetId: 'sheet',
    axis: 'cols',
    start: 16384,
    count: 1,
    level: 7,
  }
  expect(workbookOperationSchema.safeParse(operation).success).toBe(true)
  expect(workbookOperationSchema.safeParse({ ...operation, count: 2 }).success).toBe(false)
  expect(workbookOperationSchema.safeParse({ ...operation, level: 8 }).success).toBe(false)
})
