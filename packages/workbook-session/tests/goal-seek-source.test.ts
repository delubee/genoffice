import { beforeAll, afterAll, expect, it } from 'vitest'
import { build } from 'esbuild'
import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createRequire } from 'node:module'
import { fileURLToPath } from 'node:url'
import type * as GoalSeek from '../src/calculation/goal-seek'
import type { CalculationSourceData } from '../src/calculation/source'
let directory: string
let seek: typeof GoalSeek.seekCalculationSource
beforeAll(async () => {
  directory = await mkdtemp(join(tmpdir(), 'genoffice-goal-seek-'))
  const outfile = join(directory, 'solve.cjs')
  await build({
    entryPoints: [fileURLToPath(new URL('../src/calculation/goal-seek.ts', import.meta.url))],
    outfile,
    bundle: true,
    platform: 'node',
    format: 'cjs',
    target: 'node22',
    logLevel: 'silent',
  })
  seek = (createRequire(import.meta.url)(outfile) as typeof GoalSeek).seekCalculationSource
})
afterAll(async () => {
  if (directory) await rm(directory, { recursive: true, force: true })
})
function fixture(): CalculationSourceData {
  return {
    stateId: 'frozen-goal-root',
    dateSystem: '1900',
    structuralLog: [],
    workbookFeatures: {},
    definedNames: [],
    tables: [],
    sourceContext: {
      sessionId: '00000000-0000-4000-8000-000000000001',
      names: [],
      calculationProperties: {},
      externalRelationshipIds: [],
    },
    sheets: [
      {
        id: 'data',
        name: 'Data',
        rows: 10,
        columns: 3,
        sourceFeatures: {},
        logicalFeatures: {},
        sourceArrays: [],
        cells: {
          A1: { kind: 'number', rawValue: 1 },
          B1: { kind: 'number', rawValue: null, formula: '=A1*3+Other!A1' },
        },
      },
      {
        id: 'other',
        name: 'Other',
        rows: 1,
        columns: 1,
        sourceFeatures: {},
        logicalFeatures: {},
        sourceArrays: [],
        cells: { A1: { kind: 'number', rawValue: 2 } },
      },
    ],
  }
}
const input = { sheetId: 'data', setCell: 'B1', byCell: 'A1', toValue: 17 }
it('solves a frozen cross-sheet source through the real Node formula engine without mutating it', async () => {
  const source = fixture(),
    before = structuredClone(source)
  const result = await seek(source, input, { isCurrent: () => true })
  expect(result).toMatchObject({
    found: true,
    stateId: source.stateId,
    address: 'A1',
    sheetId: 'data',
  })
  // TriggerCalculationController schedules delayed progress at one second.
  // Let it expire so disposed-service notifications surface as test failures.
  await new Promise((resolve) => setTimeout(resolve, 1100))
  expect(result.solution).toBeCloseTo(5)
  expect(result.reached).toBeCloseTo(17)
  expect(source).toEqual(before)
})
it('does not publish failed guesses for a constant target', async () => {
  const source = fixture()
  source.sheets[0]!.cells.B1!.formula = '=4'
  const before = structuredClone(source)
  expect(await seek(source, input, { isCurrent: () => true })).toMatchObject({
    found: false,
    iterations: 2,
  })
  expect(source).toEqual(before)
})
it('rejects stale roots, cancellation and invalid changing cells without source mutation', async () => {
  const source = fixture(),
    before = structuredClone(source)
  let checks = 0
  await expect(seek(source, input, { isCurrent: () => ++checks < 3 })).rejects.toMatchObject({
    code: 'revision_conflict',
  })
  await expect(
    seek(source, input, { isCurrent: () => true, signal: AbortSignal.abort() }),
  ).rejects.toBeDefined()
  source.sheets[0]!.cells.A1 = { kind: 'string', rawValue: '00123' }
  await expect(seek(source, input, { isCurrent: () => true })).rejects.toThrow('number or be blank')
  source.sheets[0]!.cells.A1 = before.sheets[0]!.cells.A1!
  expect(source).toEqual(before)
})
