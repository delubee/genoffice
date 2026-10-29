import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
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
import type { SheetChartVisual as SheetVisual } from '@genoffice/xlsx-gateway/domain/chart-visual'

it('creates all nine chart types from current unsaved data and saves each exactly once', async () => {
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
        A1: { kind: 'string', rawValue: 'Year' },
        B1: { kind: 'string', rawValue: 'Sales' },
        A2: { kind: 'number', rawValue: 2020 },
        B2: { kind: 'number', rawValue: 120 },
        A3: { kind: 'number', rawValue: 2021 },
        B3: { kind: 'number', rawValue: -30 },
        A4: { kind: 'number', rawValue: 2022 },
        B4: { kind: 'number', rawValue: 45 },
      }
      state.sheets[0]!.columns = 2
      state.sheets[0]!.rows = 4
      const operations = (
        ['column', 'bar', 'line', 'area', 'pie', 'scatter', 'radar', 'doughnut', 'combo'] as const
      ).map((chartType) => ({
        op: 'add_chart' as const,
        sheetId,
        dataRange: 'A1:B4',
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
          summary: 'Nine desktop chart families',
          operations,
        },
        material,
      )
      const charts = result.nextState.sheets[0]!.features!.charts as unknown as SheetVisual[]
      expect(charts).toHaveLength(9)
      expect(charts.map((chart) => chart.chart.series[0]!.values)).toEqual(
        Array(9).fill([120, -30, 45]),
      )
      expect(result.inverse.state).toEqual(state)
      const edits: WorkbookOperation[] = [
        {
          op: 'edit_visual',
          visualId: charts[0]!.id,
          altText: 'Current sales & comparisons',
          editAs: 'oneCell',
          zIndex: 8,
        },
        {
          op: 'edit_chart' as const,
          chartPath: charts[4]!.id,
          pointColors: { '0': { '1': '#123456' } },
        },
        {
          op: 'edit_chart',
          chartPath: charts[0]!.id,
          legend: 'bottom',
          dataLabels: 'value',
          dataLabelPosition: 'inside-end',
          dataLabelFormat: '0.00',
          gridlines: false,
          valueAxis: { min: -50, max: 200 },
          gapWidthPct: 60,
        },
        {
          op: 'edit_chart',
          chartPath: charts[7]!.id,
          holeSizePct: 65,
          explosionPct: 25,
          pointExplosions: { '1': 40 },
        },
      ]
      const edited = await computeCandidateState(
        result.nextState,
        {
          schemaVersion: 1,
          sessionId: 'charts',
          transactionId: 'pie-color',
          expectedRevision: 1,
          origin: 'manual',
          summary: 'Pie point color',
          operations: edits,
        },
        await prepareCandidateMaterial({
          state: result.nextState,
          operations: edits,
          source,
          sourceSessionId: opened.original.sidecarSessionId,
        }),
      )
      expect(edited.inverse.state).toEqual(result.nextState)
      const writer = new SessionPackageWriter(client, opened.metadata)
      for (const filename of ['first.xlsx', 'again.xlsx']) {
        const target = join(directory, filename),
          staged = await writer.stage(opened.original.path, edited.nextState, target)
        await writer.publish(staged.stagedPath, target, null)
        const reopened = await source.openSnapshot(target)
        try {
          const visuals = (reopened.metadata as { visuals: SheetVisual[] }).visuals
          expect(visuals.filter((visual) => visual.kind === 'chart')).toHaveLength(9)
          expect(visuals.map((visual) => visual.chart?.title).sort()).toEqual([
            'area',
            'bar',
            'column',
            'combo',
            'doughnut',
            'line',
            'pie',
            'radar',
            'scatter',
          ])
          expect(visuals[0]!.chart!.series[0]!.values).toEqual([120, -30, 45])
          expect(visuals.find((visual) => visual.chart.title === 'column')!.chart).toMatchObject({
            legend: 'bottom',
            dataLabels: 'value',
            dataLabelPosition: 'inside-end',
            dataLabelFormat: '0.00',
            gridlines: false,
            valueAxis: { min: -50, max: 200 },
            gapWidthPct: 60,
          })
          expect(visuals.find((v) => v.chart.title === 'column')).toMatchObject({
            altText: 'Current sales & comparisons',
            editAs: 'oneCell',
          })
          expect(visuals.at(-1)!.chart.title).toBe('column')
          const doughnut = visuals.find((visual) => visual.chart.title === 'doughnut')!.chart
          expect(doughnut.holeSizePct).toBe(65)
          expect(doughnut.series[0]!.explosionPct).toBe(25)
          expect(doughnut.series[0]!.pointExplosions).toContainEqual({ index: 1, pct: 40 })
          expect(
            visuals.find((visual) => visual.chart?.title === 'pie')!.chart!.series[0]!.pointColors,
          ).toContainEqual({ index: 1, color: '#123456' })
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

it('replaces original chart series for every supported family and preserves canonical data across repeated saves', async () => {
  const { planCellEditsToXlsx, createBufferEntrySource, assembleWithJsZip } =
    await import('@genoffice/xlsx-gateway/gateway/xlsx-gateway')
  await withTestWorkspace(async (directory) => {
    const client = new XlsxSidecarClient(
      fileURLToPath(
        new URL('../../sheets/native/xlsx-engine/target/release/xlsx-sidecar', import.meta.url),
      ),
    )
    const source = new SidecarWorkbookSource(client),
      path = join(directory, 'source-series.xlsx')
    const base = await manualWorkbook(),
      families = ['column', 'bar', 'line', 'area', 'pie', 'scatter', 'radar', 'doughnut'] as const
    const plan = await planCellEditsToXlsx(
      await createBufferEntrySource(base),
      [],
      [],
      [],
      undefined,
      [],
      [],
      [],
      [],
      [],
      null,
      families.map((chartType) => ({
        sheetName: 'Data',
        anchor: {
          fromRow: 0,
          fromColumn: 3,
          fromRowOffset: 0,
          fromColumnOffset: 0,
          toRow: 12,
          toColumn: 8,
          toRowOffset: 0,
          toColumnOffset: 0,
        },
        chart: {
          chartType,
          title: chartType,
          series: [{ name: 'Source', values: [1, 2], categories: ['A', 'B'], color: '#123456' }],
        },
      })),
    )
    await writeFile(path, (await assembleWithJsZip(base, plan)).buffer)
    const opened = await openWorkbookSnapshot({
      uri: new URL(`file://${path}`).href,
      storageDirectory: directory,
      source,
    })
    try {
      const state = initialState(opened.metadata),
        visuals = (opened.metadata as { visuals: SheetVisual[] }).visuals
      expect(visuals).toHaveLength(8)
      const operations: WorkbookOperation[] = visuals.flatMap((v): WorkbookOperation[] => [
        { op: 'edit_chart', chartPath: v.chartPath!, seriesColors: { '0': '#123456' } },
        {
          op: 'edit_chart',
          chartPath: v.chartPath!,
          seriesSelection: {
            mode: 'rows',
            rows: [
              { name: 'Current data', valuesRef: "'Other'!A1", categoriesRef: "'Data'!A1" },
              { name: 'Retained', sourceIndex: 0 },
            ],
          },
        },
      ])
      operations.push({
        op: 'edit_visual',
        visualId: visuals[0]!.id,
        altText: 'Source sales & comparisons',
        editAs: 'absolute',
        zIndex: 7,
      })
      const candidate = await computeCandidateState(
        state,
        {
          schemaVersion: 1,
          sessionId: 'series',
          transactionId: 'replace',
          expectedRevision: 0,
          origin: 'manual',
          summary: 'Replace original series',
          operations,
        },
        await prepareCandidateMaterial({
          state,
          operations,
          source,
          sourceSessionId: opened.original.sidecarSessionId,
        }),
      )
      const writer = new SessionPackageWriter(client, opened.metadata)
      for (const filename of ['first-series.xlsx', 'again-series.xlsx']) {
        const target = join(directory, filename),
          staged = await writer.stage(opened.original.path, candidate.nextState, target)
        await writer.publish(staged.stagedPath, target, null)
        const reopened = await source.openSnapshot(target)
        try {
          const saved = (reopened.metadata as { visuals: SheetVisual[] }).visuals
          expect(saved).toHaveLength(8)
          expect(saved.find((v) => v.chart.title === 'column')).toMatchObject({
            altText: 'Source sales & comparisons',
            editAs: 'absolute',
          })
          expect(saved.map((v) => v.chart.title)).toEqual([...families.slice(1), 'column'])
          for (const visual of saved) {
            expect(visual.chart.series).toHaveLength(2)
            expect(visual.chart.series[0]).toMatchObject({
              name: 'Current data',
              values: [7],
              categories: visual.chart.chartTypes.includes('scatterChart') ? ['123'] : ['00123'],
              valuesRef: "'Other'!$A$1:$A$1",
              categoriesRef: "'Data'!$A$1:$A$1",
            })
            expect(visual.chart.series[1]).toMatchObject({
              name: 'Retained',
              values: [1, 2],
              categories: visual.chart.chartTypes.includes('scatterChart')
                ? ['0', '1']
                : ['A', 'B'],
              color: '#123456',
            })
          }
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

it('persists mixed session shapes, text boxes and charts in their authored layer order on repeated saves', async () => {
  await withTestWorkspace(async (directory) => {
    const client = new XlsxSidecarClient(
        fileURLToPath(
          new URL('../../sheets/native/xlsx-engine/target/release/xlsx-sidecar', import.meta.url),
        ),
      ),
      source = new SidecarWorkbookSource(client),
      path = join(directory, 'mixed-visuals.xlsx')
    await writeFile(path, await manualWorkbook())
    const opened = await openWorkbookSnapshot({
      uri: new URL(`file://${path}`).href,
      storageDirectory: directory,
      source,
    })
    try {
      let state = structuredClone(initialState(opened.metadata))
      const sheetId = state.sheets[0]!.id
      state.sheets[0]!.cells = {
        A1: { kind: 'string', rawValue: 'Name' },
        B1: { kind: 'string', rawValue: 'Value' },
        A2: { kind: 'string', rawValue: 'A' },
        B2: { kind: 'number', rawValue: 4 },
      }
      const apply = async (operations: WorkbookOperation[]) => {
        const before = structuredClone(state),
          result = await computeCandidateState(
            state,
            {
              schemaVersion: 1,
              sessionId: 'mixed',
              transactionId: crypto.randomUUID(),
              expectedRevision: 0,
              origin: 'manual',
              summary: 'Mixed visuals',
              operations,
            },
            await prepareCandidateMaterial({
              state,
              operations,
              source,
              sourceSessionId: opened.original.sidecarSessionId,
            }),
          )
        expect(result.inverse.state).toEqual(before)
        state = result.nextState
      }
      await apply([
        { op: 'add_chart', sheetId, chartType: 'column', dataRange: 'A1:B2', title: 'Mixed chart' },
        { op: 'add_shape', sheetId, shapeType: 'rect', anchorCell: 'D5', text: 'Authored shape' },
        { op: 'add_shape', sheetId, shapeType: 'textbox', anchorCell: 'F12' },
        {
          op: 'add_picture',
          sheetId,
          name: 'Mixed.png',
          mediaType: 'image/png',
          base64:
            'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
          anchor: {
            fromRow: 15,
            fromColumn: 3,
            fromRowOffset: 0,
            fromColumnOffset: 0,
            toRow: 17,
            toColumn: 5,
            toRowOffset: 0,
            toColumnOffset: 0,
          },
        },
      ])
      const visuals = state.sheets[0]!.features!.visuals as any[],
        shape = visuals[1]
      await apply([
        { op: 'edit_shape', visualId: shape.id, text: 'Saved & manual text' },
        {
          op: 'edit_visual',
          visualId: shape.id,
          fillColor: 'none',
          lineColor: '#123456',
          altText: 'Mixed shape',
          rotation: 35,
          flipH: true,
          zIndex: 0,
        },
      ])
      await apply([
        {
          op: 'edit_visual',
          visualId: visuals[3].id,
          altText: 'Mixed picture',
          rotation: 90,
          flipV: true,
          frameSize: { width: 120, height: 80 },
        },
      ])
      await apply(
        Array.from({ length: 51 }, (_, i) => ({
          op: 'add_shape' as const,
          sheetId,
          shapeType: 'rect' as const,
          anchorCell: 'D20',
          text: `Extra ${i}`,
        })),
      )
      const writer = new SessionPackageWriter(client, opened.metadata)
      for (const name of ['first-mixed.xlsx', 'again-mixed.xlsx']) {
        const target = join(directory, name),
          staged = await writer.stage(opened.original.path, state, target)
        await writer.publish(staged.stagedPath, target, null)
        const reopened = await source.openSnapshot(target)
        try {
          const saved = (reopened.metadata as any).visuals
          expect(saved).toHaveLength(55)
          expect(saved.slice(0, 3).map((v: any) => v.kind)).toEqual(['shape', 'chart', 'shape'])
          expect(saved.slice(4).map((v: any) => v.text)).toEqual(
            Array.from({ length: 51 }, (_, i) => `Extra ${i}`),
          )
          expect(saved[0]).toMatchObject({
            altText: 'Mixed shape',
            text: 'Saved & manual text',
            fillColor: 'none',
            lineColor: '#123456',
            rotation: 35,
            flipH: true,
          })
          expect(saved[1].chart.title).toBe('Mixed chart')
          expect(saved[2].text).toBe('Text')
          expect(saved[3]).toMatchObject({
            kind: 'image',
            mediaType: 'image/png',
            altText: 'Mixed picture',
            rotation: 90,
            flipV: true,
            frameWidth: 120,
            frameHeight: 80,
          })
          expect((await source.readMedia(reopened.sessionId, saved[3].id)).base64).toBe(
            'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
          )
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
