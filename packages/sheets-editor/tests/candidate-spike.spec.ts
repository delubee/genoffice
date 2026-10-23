import { test, expect } from '@playwright/test'
import { createServer, type ViteDevServer } from 'vite'
import { fileURLToPath } from 'node:url'

let server: ViteDevServer
let base: string

test.beforeAll(async () => {
  server = await createServer({
    configFile: false,
    root: fileURLToPath(new URL('../../..', import.meta.url)),
    server: { host: '127.0.0.1', port: 0 },
  })
  await server.listen()
  base = server.resolvedUrls!.local[0]!
})
test.afterAll(async () => {
  await server?.close()
})

test('isolated Univer candidates, pre-command interception and serializable structure/chart undo', async ({
  page,
}) => {
  await page.goto(base)
  const result = await page.evaluate(async () => {
    const { runSpike } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/candidate-spike.ts')
    )
    return runSpike()
  })
  expect(result).toEqual({
    failedCandidateDiscarded: true,
    blockedBeforeMutation: true,
    structuralChange: true,
    serializedChart: true,
    restored: true,
    activeRuntimes: 0,
  })
})

test('production interpreter discards failures and captures bidirectional structure/chart roots', async ({
  page,
}) => {
  await page.goto(base)
  const result = await page.evaluate(async () => {
    const { runProductionCandidate } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/candidate-spike.ts')
    )
    return runProductionCandidate()
  })
  expect(result).toEqual({
    rejected: true,
    intact: true,
    restored: true,
    chart: true,
    moved: true,
    disposed: true,
  })
})

test('candidate shifts cross-sheet references, preserves literal type and rejects protected writes', async ({
  page,
}) => {
  await page.goto(base)
  const result = await page.evaluate(async () => {
    const { runCandidateGuards } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/candidate-spike.ts')
    )
    return runCandidateGuards()
  })
  expect(result).toEqual({
    formula: '=Renamed!A3',
    literalKind: 'string',
    protectedRejected: true,
    noRuntimes: true,
  })
})
