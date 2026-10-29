import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { expect, it } from 'vitest'
import { prepareCandidateMaterial } from '../src/candidate-material'
import { prepareStagedCandidate } from '../src/staged-candidate'
import { WorkbookTransactions } from '../src/transactions'
import { FileResourceStore } from '../src/store'
import { initialState } from '../src/state'
import type { WorkbookSource } from '../src/ports'
import { useSessionTestResources } from './helpers'
const resources = useSessionTestResources()

const source: WorkbookSource = {
  async openSnapshot() {
    throw new Error('Unused')
  },
  async archiveManifest() {
    return {}
  },
  async close() {},
  async readFeatures() {
    return { indexingComplete: true }
  },
  async readBaseRange() {
    return { cells: [], indexingComplete: true }
  },
}

function fixture() {
  const state = initialState({
    sheets: [{ id: 'base', name: 'Data', rowCount: 2, columnCount: 2 }],
  })
  const sheetId = state.sheets[0]!.id
  const intent = {
    schemaVersion: 1 as const,
    sessionId: 'session',
    transactionId: 'bulk',
    expectedRevision: 0,
    origin: 'manual' as const,
    summary: 'Import',
    operations: [] as WorkbookOperation[],
  }
  // 4,500 cells: above the 2,000-cell ordinary batch, in three bounded stages.
  const band = (startRow: number): WorkbookOperation[] => [
    {
      op: 'set_range',
      sheetId,
      start: `A${startRow + 1}`,
      values: Array.from({ length: 30 }, (_, row) =>
        Array.from({ length: 50 }, (_, column) => (startRow + row) * 100 + column),
      ),
    },
  ]
  const stages = [band(0), band(30), band(60)]
  return { state, sheetId, intent, stages }
}

it('publishes bounded stages above the ordinary write ceiling as one changeset', async () => {
  const { state, sheetId, intent, stages } = fixture()
  const all = stages.flat()
  const material = await prepareCandidateMaterial({
    state,
    operations: all,
    source,
    sourceSessionId: 'source',
  })
  await expect(
    computeCandidateState(state, { ...intent, operations: all }, material),
  ).rejects.toThrow('expanded_write_limit')
  const changes = await prepareStagedCandidate({
    state,
    intent,
    stages,
    source,
    sourceSessionId: 'source',
  })
  expect(changes.beforeStateId).toBe(state.stateId)
  const cells = changes.nextState.sheets.find((sheet) => sheet.id === sheetId)!.cells
  expect(cells.A1!.rawValue).toBe(0)
  expect(cells.AX90!.rawValue).toBe(8949)
  expect(Object.keys(cells)).toHaveLength(4500)
})

it('adds a sheet in its own stage before writing its content', async () => {
  const { state, intent } = fixture()
  const changes = await prepareStagedCandidate({
    state,
    intent,
    stages: [[{ op: 'add_sheet', name: 'Merged' }]],
    source,
    sourceSessionId: 'source',
  })
  const added = changes.nextState.sheets.find((sheet) => sheet.name === 'Merged')!
  const next = await prepareStagedCandidate({
    state: changes.nextState,
    intent,
    stages: [
      [{ op: 'add_sheet', name: 'Second' }],
      [{ op: 'set_cell', sheetId: added.id, address: 'B2', value: 'kept' }],
    ],
    source,
    sourceSessionId: 'source',
  })
  expect(next.nextState.sheets.map((sheet) => sheet.name)).toEqual(['Data', 'Merged', 'Second'])
  expect(next.nextState.sheets[1]!.cells.B2!.rawValue).toBe('kept')
})

it('refuses staged transactions above the desktop bulk budget', async () => {
  const { state, sheetId, intent } = fixture()
  const stage: WorkbookOperation[] = [
    {
      op: 'set_range',
      sheetId,
      start: 'A1',
      values: Array.from({ length: 500 }, () => Array.from({ length: 100 }, () => 1)),
    },
  ]
  await expect(
    prepareStagedCandidate({
      state,
      intent,
      stages: Array.from({ length: 2 }, () => stage),
      source,
      sourceSessionId: 'source',
    }),
  ).rejects.toMatchObject({ code: 'expanded_write_limit' })
})

it('commits staged work as one revision and one undoable history entry', async () => {
  const { state, sheetId, intent, stages } = fixture()
  const store = await FileResourceStore.create(await resources.directory())
  const engine = new WorkbookTransactions({
    sessionId: 'session',
    documentUri: 'file:///test.xlsx',
    state,
    store,
    interpreter: {
      async prepare() {
        throw new Error('ordinary prepare must not run')
      },
      prepareStaged: (before, staged, parts, signal) =>
        prepareStagedCandidate({
          state: before,
          intent: staged,
          stages: parts,
          source,
          sourceSessionId: 'source',
          signal,
        }),
      dispose() {},
    },
    host: { committed() {}, projectionChanged() {}, conflict() {} },
  })
  const { operations: _unused, ...raw } = intent
  const receipt = await engine.submitStaged(raw, stages)
  expect(receipt.revision).toBe(1)
  expect(engine.history).toHaveLength(1)
  expect(Object.keys(engine.state.sheets[0]!.cells)).toHaveLength(4500)
  expect(await engine.submitStaged(raw, stages)).toEqual(receipt)
  await engine.undo()
  expect(engine.state.sheets.find((sheet) => sheet.id === sheetId)!.cells).toEqual({})
  await engine.dispose()
})
