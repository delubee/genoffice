import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import { expect, it, vi } from 'vitest'
import { prepareCandidateMaterial } from '../src/candidate-material'
import { initialState } from '../src/state'
import type { WorkbookSource } from '../src/ports'

it('waits for complete protection metadata before reading edit material or granting write access', async () => {
  const state = initialState({
      sheets: [{ id: 'base', name: 'Data', rowCount: 1, columnCount: 1 }],
    }),
    sheetId = state.sheets[0]!.id
  let indexed = false
  const readFeatures = vi.fn(async () => {
    if (!indexed) {
      indexed = true
      return { indexingComplete: false }
    }
    return { indexingComplete: true, sheetProtection: { protected: true } }
  })
  const source: WorkbookSource = {
    async openSnapshot() {
      throw new Error('Unused')
    },
    async archiveManifest() {
      return {}
    },
    async close() {},
    readFeatures,
    async readBaseRange() {
      expect(indexed).toBe(true)
      expect(readFeatures.mock.calls.length).toBeGreaterThanOrEqual(2)
      return { cells: [{ row: 0, column: 0, value: 7 }] }
    },
  }
  const material = await prepareCandidateMaterial({
    state,
    source,
    sourceSessionId: 'source',
    operations: [{ op: 'set_cell', sheetId, address: 'A1', value: 8 }],
  })
  expect(material.writableSheetIds).toEqual([])
  expect(material.state.sheets[0]!.protected).toBe(true)
  expect(material.state.sheets[0]!.cells.A1!.rawValue).toBe(7)
})

it('hydrates both replacement and retained original-series dependencies before interpreting a selection', async () => {
  const state = structuredClone(
    initialState({
      sheets: [
        { id: 'base', name: 'Data', rowCount: 3, columnCount: 2 },
        { id: 'other', name: 'Other', rowCount: 2, columnCount: 1 },
      ],
    }),
  )
  const owner = state.sheets[0]!,
    other = state.sheets[1]!
  state.features = {
    originalMetadata: {
      visuals: [
        {
          id: 'source',
          kind: 'chart',
          sheetId: 'base',
          chartPath: 'xl/charts/chart1.xml',
          chart: {
            chartTypes: ['barChart'],
            series: [
              {
                name: 'Sales',
                values: [],
                categories: [],
                valuesRef: "'Data'!$B$2:$B$3",
                categoriesRef: "'Data'!$A$2:$A$3",
              },
            ],
          },
        },
      ],
    },
  }
  const reads: string[] = []
  const source: WorkbookSource = {
    async openSnapshot() {
      throw new Error('unused')
    },
    async archiveManifest() {
      return {}
    },
    async close() {},
    async readFeatures() {
      return { indexingComplete: true }
    },
    async readBaseRange(_session, sheet, range) {
      reads.push(`${sheet}:${range}`)
      return { cells: [] }
    },
  }
  const material = await prepareCandidateMaterial({
    state,
    source,
    sourceSessionId: 'source',
    operations: [
      {
        op: 'edit_chart',
        chartPath: 'xl/charts/chart1.xml',
        seriesSelection: {
          mode: 'rows',
          rows: [
            { sourceIndex: 0, name: 'Retained' },
            { name: 'Added', valuesRef: "'Other'!A1:A2" },
          ],
        },
      },
    ],
  })
  expect(material.referenceRanges[owner.id]).toEqual(expect.arrayContaining(['B2:B3', 'A2:A3']))
  expect(material.referenceRanges[other.id]).toEqual(['A1:A2'])
  expect(reads).toHaveLength(3)
  owner.features = {
    chartEdits: {
      'xl/charts/chart1.xml': {
        chartPath: 'xl/charts/chart1.xml',
        seriesSet: [{ name: 'New', values: [1, 2], valuesRef: "'Other'!$A$1:$A$2" }],
      },
    },
  }
  const refresh = await prepareCandidateMaterial({
    state,
    source,
    sourceSessionId: 'source',
    operations: [{ op: 'set_cell', sheetId: owner.id, address: 'B1', value: 1 }],
  })
  expect(refresh.referenceRanges[other.id]).toEqual(['A1:A2'])
})

it('hydrates complete worksheets for a cut so dependencies outside both selected blocks are visible', async () => {
  const state = initialState({
    sheets: [
      { id: 'base', name: 'Data', rowCount: 90, columnCount: 10 },
      { id: 'other', name: 'Other', rowCount: 80, columnCount: 5 },
    ],
  })
  const reads: string[] = []
  const source: WorkbookSource = {
    async openSnapshot() {
      throw new Error('unused')
    },
    async archiveManifest() {
      return {}
    },
    async close() {},
    async readFeatures() {
      return { indexingComplete: true, styles: [] }
    },
    async readBaseRange(_session, sheetId, range) {
      reads.push(`${sheetId}:${range}`)
      return {
        cells: sheetId === 'base' ? [{ row: 89, column: 9, value: 7, formula: '=$A$2' }] : [],
      }
    },
  }
  const result = await prepareCandidateMaterial({
    state,
    source,
    sourceSessionId: 'source',
    operations: [
      {
        op: 'relocate_range',
        sourceSheetId: state.sheets[0]!.id,
        sheetId: state.sheets[1]!.id,
        source: 'A2:B3',
        start: 'D5',
      },
    ],
  })
  expect(reads).toContain('base:A1:J90')
  expect(reads).toContain('other:A1:E80')
  expect(result.completeSheetIds).toEqual(state.sheets.map((s) => s.id))
  expect(result.state.sheets[0]!.cells.J90!.formula).toBe('=$A$2')
})

it('indexes protection without reading unrelated chart data for print metadata', async () => {
  const state = structuredClone(
    initialState({
      sheets: [{ id: 'base', name: 'Data', rowCount: 3, columnCount: 2 }],
    }),
  )
  const sheet = state.sheets[0]!
  sheet.features = {
    chartEdits: { 'xl/charts/chart1.xml': { series: [{ valuesRef: "'Missing'!$A$1:$A$3" }] } },
  }
  const readBaseRange = vi.fn(async () => ({ cells: [] }))
  const readFeatures = vi.fn(async () => ({
    indexingComplete: true,
    sheetProtection: { protected: true },
  }))
  const source: WorkbookSource = {
    async openSnapshot() {
      throw new Error('unused')
    },
    async archiveManifest() {
      return {}
    },
    async close() {},
    readFeatures,
    readBaseRange,
  }
  const material = await prepareCandidateMaterial({
    state,
    source,
    sourceSessionId: 'source',
    operations: [{ op: 'set_page_setup', sheetId: sheet.id, orientation: 'landscape' }],
  })
  expect(readFeatures).toHaveBeenCalled()
  expect(readBaseRange).not.toHaveBeenCalled()
  expect(material.referenceRanges).toEqual({})
  expect(material.writableSheetIds).toEqual([])
  expect(material.state.sheets[0]!.protected).toBe(true)
  expect(material.state.sheets[0]!.features!.chartEdits).toEqual(sheet.features.chartEdits)
})

it('reads imported hidden-row metadata over the sort span instead of relying on viewport rows', async () => {
  const state = initialState({
    sheets: [{ id: 'base', name: 'Data', rowCount: 80, columnCount: 2 }],
  })
  const sheetId = state.sheets[0]!.id
  const source: WorkbookSource = {
    async openSnapshot() {
      throw new Error('Unused')
    },
    async archiveManifest() {
      return {}
    },
    async close() {},
    async readFeatures() {
      return {
        indexingComplete: true,
        merges: [],
        autoFilter: null,
        autoFilterColumns: [],
        sheetProtection: null,
      }
    },
    async readBaseRange(_session, _sheet, range) {
      expect(range).toContain('80')
      return {
        cells: [],
        rows: [
          { row: 74, hidden: true },
          { row: 79, hidden: true },
        ],
      }
    },
  }
  const material = await prepareCandidateMaterial({
    state,
    source,
    sourceSessionId: 'source',
    operations: [{ op: 'sort_range', sheetId, range: 'A70:B80', byColumn: 'B', order: 'asc' }],
  })
  expect(material.hiddenRows[sheetId]).toEqual([74, 79])
  const operations = [{ op: 'set_filter' as const, sheetId, range: 'A70:B80' }]
  const filterMaterial = await prepareCandidateMaterial({
    state,
    source,
    sourceSessionId: 'source',
    operations,
  })
  const result = await computeCandidateState(
    state,
    {
      schemaVersion: 1,
      sessionId: 'source',
      transactionId: crypto.randomUUID(),
      expectedRevision: 0,
      origin: 'manual',
      summary: 'Create filter',
      operations,
    },
    filterMaterial,
  )
  expect(result.nextState.sheets[0]!.features!.visibilityEdits).toEqual([
    { axis: 'rows', start: 74, count: 1, hidden: true },
    { axis: 'rows', start: 79, count: 1, hidden: true },
  ])
})
