import { beforeEach, expect, it, vi } from 'vitest'
const mocks = vi.hoisted(() => ({ prepare: vi.fn(), load: vi.fn(), solve: vi.fn() }))
vi.mock('../src/candidate-material', () => ({ prepareCandidateMaterial: mocks.prepare }))
vi.mock('../src/calculation/source', () => ({ loadCalculationSource: mocks.load }))
vi.mock('../src/calculation/goal-seek', () => ({ seekCalculationSource: mocks.solve }))
import { WorkbookGoalSeekController } from '../src/goal-seek-controller'
import { initialState } from '../src/state'
import type { DocumentDescriptor, TransactionInput } from '../src/domain'
import type { WorkbookSource } from '../src/ports'
beforeEach(() => vi.resetAllMocks())
function fixture() {
  const state = structuredClone(
    initialState({
      sheets: [{ id: 'base', name: 'Data', rowCount: 1, columnCount: 2 }],
    }),
  )
  const sheet = state.sheets[0]!
  sheet.cells.A1 = { kind: 'number', rawValue: 1 }
  sheet.cells.B1 = { kind: 'number', rawValue: null, formula: '=A1*3+2' }
  const submit = vi.fn(async (_intent: TransactionInput) => ({}))
  const controller = new WorkbookGoalSeekController({
    sessionId: 'session',
    sourceSessionId: 'private-source',
    source: {} as WorkbookSource,
    getState: () => state,
    getDescriptor: () => ({ stateId: state.stateId, revision: 2 }) as DocumentDescriptor,
    submit,
  })
  mocks.prepare.mockImplementation(async () => ({
    state: structuredClone(state),
    writableSheetIds: [sheet.id],
  }))
  mocks.load.mockResolvedValue({ privateResource: 'never-return-this' })
  mocks.solve.mockResolvedValue({ found: true, solution: 5, reached: 17, iterations: 3 })
  const request = {
    stateId: state.stateId,
    revision: 2,
    sheetId: sheet.id,
    setCell: 'B1',
    byCell: 'A1',
    toValue: 17,
  }
  return { state, sheet, submit, controller, request }
}
it('publishes only one final input edit with the captured revision and no private source data', async () => {
  const f = fixture(),
    before = structuredClone(f.state)
  const response = await f.controller.handle(f.request)
  expect(response).toMatchObject({ found: true, solution: 5, reached: 17, iterations: 3 })
  expect(f.submit).toHaveBeenCalledTimes(1)
  expect(f.submit.mock.calls[0]![0]).toMatchObject({
    expectedRevision: 2,
    origin: 'manual',
    operations: [{ op: 'set_cell', sheetId: f.sheet.id, address: 'A1', value: 5 }],
  })
  expect(f.state).toEqual(before)
  f.controller.dispose()
})
it('returns a JSON-safe nonconvergence result without any write', async () => {
  const f = fixture()
  mocks.solve.mockResolvedValue({ found: false, solution: 4, reached: NaN, iterations: 2 })
  expect(await f.controller.handle(f.request)).toMatchObject({ found: false, reached: null })
  expect(f.submit).not.toHaveBeenCalled()
  f.controller.dispose()
})
it('rejects protected targets before loading calculation sources', async () => {
  const f = fixture()
  mocks.prepare.mockResolvedValue({ state: f.state, writableSheetIds: [] })
  await expect(f.controller.handle(f.request)).rejects.toMatchObject({ code: 'protected_target' })
  expect(mocks.load).not.toHaveBeenCalled()
  expect(f.submit).not.toHaveBeenCalled()
  f.controller.dispose()
})
it('rejects results overtaken by another revision or a disposed view', async () => {
  const f = fixture()
  mocks.solve.mockImplementation(async () => {
    f.state.stateId = 'changed'
    return { found: true, solution: 5, reached: 17, iterations: 3 }
  })
  await expect(f.controller.handle(f.request)).rejects.toMatchObject({ code: 'revision_conflict' })
  expect(f.submit).not.toHaveBeenCalled()
  const other = fixture()
  mocks.solve.mockImplementation(async () => {
    other.controller.dispose()
    return { found: true, solution: 5, reached: 17, iterations: 3 }
  })
  await expect(other.controller.handle(other.request)).rejects.toBeDefined()
  expect(other.submit).not.toHaveBeenCalled()
  f.controller.dispose()
})
it('rejects stale or extra request fields before starting evaluation', async () => {
  const f = fixture()
  await expect(f.controller.handle({ ...f.request, revision: 1 })).rejects.toMatchObject({
    code: 'revision_conflict',
  })
  await expect(f.controller.handle({ ...f.request, formula: '=malicious' })).rejects.toBeDefined()
  expect(mocks.prepare).not.toHaveBeenCalled()
  f.controller.dispose()
})

it('does not create history for an already satisfied input and rejects overlapping solves', async () => {
  const f = fixture()
  f.sheet.cells.A1!.rawValue = 5
  await f.controller.handle(f.request)
  expect(f.submit).not.toHaveBeenCalled()
  let release!: () => void
  mocks.solve.mockImplementation(async () => {
    await new Promise<void>((resolve) => {
      release = resolve
    })
    return { found: true, solution: 6, reached: 17, iterations: 3 }
  })
  const pending = f.controller.handle(f.request)
  await vi.waitFor(() => expect(release).toBeTypeOf('function'))
  await expect(f.controller.handle(f.request)).rejects.toMatchObject({ code: 'operation_busy' })
  release()
  await pending
  expect(f.submit).toHaveBeenCalledTimes(1)
  f.controller.dispose()
})
