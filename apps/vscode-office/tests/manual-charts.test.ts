import { expect, it } from 'vitest'
import { writeFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import { join } from 'node:path'
import { manualWorkbook } from './fixtures/manual-workbook'
import { withTestWorkspace } from './helpers'
import { XlsxSidecarClient } from '@genoffice/workbook-session/sidecar/client'
import { SidecarWorkbookSource } from '@genoffice/workbook-session/sidecar/source'
import { openWorkbookSnapshot } from '@genoffice/workbook-session/open'
import { initialState } from '@genoffice/workbook-session/state'
import { prepareCandidateMaterial } from '@genoffice/workbook-session/candidate-material'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import { SessionPackageWriter } from '@genoffice/workbook-session/package-writer'
import type { SheetVisual } from '@genoffice/xlsx-gateway/domain/chart-visual'

it('creates all four chart types from current unsaved data and saves each exactly once', async () => {
  await withTestWorkspace(async (directory) => {
    const client = new XlsxSidecarClient(
        fileURLToPath(
          new URL('../../sheets/native/xlsx-engine/target/release/xlsx-sidecar', import.meta.url),
        ),
      ),
      source = new SidecarWorkbookSource(client),
      path = join(directory, 'charts.xlsx')
    await writeFile(path, await manualWorkbook())
    const opened = await openWorkbookSnapshot({
      uri: new URL(`file://${path}`).href,
      storageDirectory: directory,
      source,
    })
    try {
      const state = structuredClone(initialState(opened.metadata)),
        sheetId = state.sheets[0]!.id
      state.sheets[0]!.cells = {
        A1: { kind: 'string', rawValue: 'Category' },
        B1: { kind: 'string', rawValue: 'Sales' },
        A2: { kind: 'string', rawValue: 'North' },
        B2: { kind: 'number', rawValue: 120 },
        A3: { kind: 'string', rawValue: 'South' },
        B3: { kind: 'number', rawValue: -30 },
      }
      state.sheets[0]!.columns = 2
      const operations = (['column', 'bar', 'line', 'pie'] as const).map((chartType) => ({
        op: 'add_chart' as const,
        sheetId,
        dataRange: 'A1:B3',
        chartType,
        title: chartType,
      }))
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
          sessionId: 'charts',
          transactionId: 'charts',
          expectedRevision: 0,
          origin: 'manual',
          summary: 'Four basic charts',
          operations,
        },
        material,
      )
      const charts = result.nextState.sheets[0]!.features!.charts as unknown as SheetVisual[]
      expect(charts).toHaveLength(4)
      expect(charts.map((chart) => chart.chart.series[0]!.values)).toEqual(
        Array(4).fill([120, -30]),
      )
      expect(result.inverse.state).toEqual(state)
      const writer = new SessionPackageWriter(client, opened.metadata)
      for (const filename of ['first.xlsx', 'again.xlsx']) {
        const target = join(directory, filename),
          staged = await writer.stage(opened.original.path, result.nextState, target)
        await writer.publish(staged.stagedPath, target, null)
        const reopened = await source.openSnapshot(target)
        try {
          const visuals = (
            reopened.metadata as {
              visuals: { kind: string; chart?: { title: string; series: { values: number[] }[] } }[]
            }
          ).visuals
          expect(visuals.filter((visual) => visual.kind === 'chart')).toHaveLength(4)
          expect(visuals.map((visual) => visual.chart?.title).sort()).toEqual([
            'bar',
            'column',
            'line',
            'pie',
          ])
          expect(visuals[0]!.chart!.series[0]!.values).toEqual([120, -30])
        } finally {
          await source.close(reopened.sessionId)
        }
      }
    } finally {
      await opened.dispose()
      client.stop()
    }
  })
})
