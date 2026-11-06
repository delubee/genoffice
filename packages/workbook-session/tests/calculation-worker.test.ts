import { afterAll, beforeAll, expect, it } from 'vitest'
import { build } from 'esbuild'
import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createRequire } from 'node:module'
import { fileURLToPath } from 'node:url'
import type * as Client from '../src/calculation/worker-client'
import type { CalculationSourceData } from '../src/calculation/source'

let directory: string
let client: typeof Client
beforeAll(async () => {
  directory = await mkdtemp(join(tmpdir(), 'genoffice-calc-worker-'))
  // The same two bundles the extension ships: the worker entry and the client that starts it.
  const entry = (name: string) =>
    fileURLToPath(new URL(`../src/calculation/${name}.ts`, import.meta.url))
  await build({
    entryPoints: [entry('worker-entry')],
    outfile: join(directory, 'calc-worker.cjs'),
    bundle: true,
    platform: 'node',
    format: 'cjs',
    target: 'node22',
    logLevel: 'silent',
  })
  await build({
    entryPoints: [entry('worker-client')],
    outfile: join(directory, 'client.cjs'),
    bundle: true,
    platform: 'node',
    format: 'cjs',
    target: 'node22',
    logLevel: 'silent',
  })
  client = createRequire(import.meta.url)(join(directory, 'client.cjs')) as typeof Client
  client.setCalculationWorkerPath(join(directory, 'calc-worker.cjs'))
}, 120_000)
afterAll(async () => {
  client?.setCalculationWorkerPath(undefined)
  if (directory) await rm(directory, { recursive: true, force: true })
})
function fixture(): CalculationSourceData {
  return {
    stateId: 'frozen-worker-root',
    dateSystem: '1900',
    structuralLog: [],
    workbookFeatures: {},
    definedNames: [],
    tables: [],
    sourceContext: {
      sessionId: '00000000-0000-4000-8000-000000000002',
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
          A1: { kind: 'number', rawValue: 4 },
          B1: { kind: 'number', rawValue: null, formula: '=A1*3' },
        },
      },
    ],
  }
}
it('calculates on a worker thread, survives an abort and keeps serving calculations', async () => {
  const first = await client.calculateOffThread(fixture())
  expect(first.formulas).toHaveLength(1)
  expect(first.formulas[0]!.values).toEqual([[{ type: 'number', value: 12 }]])
  // An aborted calculation rejects and its worker is replaced.
  const controller = new AbortController()
  const aborted = client.calculateOffThread(fixture(), { signal: controller.signal })
  controller.abort(new Error('stop'))
  await expect(aborted).rejects.toThrow('stop')
  const again = await client.calculateOffThread(fixture())
  expect(again.formulas[0]!.values).toEqual([[{ type: 'number', value: 12 }]])
  // Without a worker the same call runs in-process.
  client.setCalculationWorkerPath(undefined)
  const inline = await client.calculateOffThread(fixture())
  expect(inline.formulas[0]!.values).toEqual([[{ type: 'number', value: 12 }]])
}, 120_000)
