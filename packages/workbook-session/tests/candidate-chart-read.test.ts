import { expect, it } from 'vitest'
import { initialState } from '../src/state'
import { prepareCandidateMaterial } from '../src/candidate-material'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import { parseRange, rangeCellCount } from '@genoffice/xlsx-gateway/domain/cell-address'
import type { WorkbookSource } from '../src/ports'

it('hydrates exact chart references in a 100k-cell source rather than rereading the complete worksheet for a scalar edit', async () => {
  const state = structuredClone(
      initialState({ sheets: [{ id: 'base', name: 'Data', rowCount: 10000, columnCount: 10 }] }),
    ),
    sheetId = state.sheets[0]!.id
  state.sheets[0]!.cells.D2 = { kind: 'number', rawValue: 120 }
  const reads: number[] = []
  const source: WorkbookSource = {
    async openSnapshot() {
      throw new Error('Not used')
    },
    async close() {},
    async archiveManifest() {
      return {}
    },
    async readFeatures() {
      return { indexingComplete: true, sheetProtection: null, styles: [] }
    },
    async readBaseRange(_session, _sheet, range) {
      const bounds = parseRange(range),
        count = rangeCellCount(bounds)
      reads.push(count)
      if (count > 20) throw new Error('Chart dependencies must not read unrelated source cells')
      const cells = []
      for (let row = bounds.startRow; row <= bounds.endRow; row++)
        for (let column = bounds.startColumn; column <= bounds.endColumn; column++)
          cells.push({
            row,
            column,
            value:
              row === 0
                ? column === 0
                  ? 'Region'
                  : 'Sales'
                : column === 0
                  ? `Region ${row}`
                  : row,
          })
      return { cells }
    },
  }
  const chartIntent = {
    schemaVersion: 1,
    sessionId: 'chart-reference',
    transactionId: 'chart',
    expectedRevision: 0,
    origin: 'agent',
    summary: 'Small chart on a large worksheet',
    operations: [
      { op: 'add_chart' as const, sheetId, chartType: 'column' as const, dataRange: 'A1:B5' },
    ],
  }
  const chart = await computeCandidateState(
    state,
    chartIntent,
    await prepareCandidateMaterial({
      state,
      operations: chartIntent.operations,
      source,
      sourceSessionId: 'base',
    }),
  )
  const edit = {
    ...chartIntent,
    transactionId: 'edit',
    expectedRevision: 1,
    operations: [{ op: 'set_cell' as const, sheetId, address: 'B2', value: 99 }],
  }
  const material = await prepareCandidateMaterial({
    state: chart.nextState,
    operations: edit.operations,
    source,
    sourceSessionId: 'base',
  })
  expect(material.completeSheetIds).not.toContain(sheetId)
  expect(material.referenceRanges[sheetId]!.length).toBeGreaterThan(1)
  const changed = await computeCandidateState(chart.nextState, edit, material)
  const charts = changed.nextState.sheets[0]!.features!.charts as {
    chart: { series: { values: number[] }[] }
  }[]
  expect(charts[0]!.chart.series[0]!.values[0]).toBe(99)
  expect(Object.keys(changed.nextState.sheets[0]!.cells).sort()).toEqual(['B2', 'D2'])
  expect(Math.max(...reads)).toBeLessThanOrEqual(10)
})
