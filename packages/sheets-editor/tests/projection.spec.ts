import { test, expect } from '@playwright/test'
import { createServer, type ViteDevServer } from 'vite'
import { fileURLToPath } from 'node:url'
import { mkdtemp, writeFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { Webview } from 'vscode'
import { createRequire } from 'node:module'
import { build } from 'esbuild'
import type { createProjectionSession } from './fixtures/projection-host'

let server: ViteDevServer
let base: string
let hostDirectory: string
let createHost: typeof createProjectionSession

test.beforeAll(async () => {
  server = await createServer({
    configFile: false,
    root: fileURLToPath(new URL('../../..', import.meta.url)),
    server: { host: '127.0.0.1', port: 0 },
  })
  await server.listen()
  base = server.resolvedUrls!.local[0]!
  hostDirectory = await mkdtemp(join(tmpdir(), 'genoffice-test-host-'))
  const outfile = join(hostDirectory, 'host.cjs')
  await build({
    entryPoints: [fileURLToPath(new URL('./fixtures/projection-host.ts', import.meta.url))],
    outfile,
    bundle: true,
    platform: 'node',
    format: 'cjs',
    external: ['vscode'],
  })
  createHost = createRequire(import.meta.url)(outfile).createProjectionSession
})
test.afterAll(async () => {
  await server?.close()
  if (hostDirectory) await rm(hostDirectory, { recursive: true, force: true })
})

test('real Node session, sidecar and browser candidate commit once and rebuild a failed projection', async ({
  page,
}) => {
  const directory = await mkdtemp(join(tmpdir(), 'genoffice-projection-'))
  let host: Awaited<ReturnType<typeof createProjectionSession>> | undefined
  try {
    const sessionId = 'projection-session'
    await page.goto(base)
    let receive: (message: unknown) => void = () => {}
    await page.exposeFunction('sendToHost', (message: unknown) => receive(message))
    await page.evaluate(async (id) => {
      const { JsonBridge, attachCandidateHandler } = await import(
        /* @vite-ignore */ String('/packages/sheets-editor/src/bridge.ts')
      )
      const scope = window as unknown as Record<string, any>
      const bridge = new JsonBridge(id, {
        post: (message: unknown) => scope.sendToHost(message),
        subscribe: (handler: (message: unknown) => void) => {
          scope.fromHost = handler
          return () => {
            scope.fromHost = () => {}
          }
        },
      })
      attachCandidateHandler(bridge)
      bridge.handle('state.committed', () => {
        throw new Error('Injected lost projection')
      })
      bridge.handle('view.resync', (payload: unknown) => {
        scope.projected = payload
        return { rebuilt: true }
      })
      scope.documentBridge = bridge
    }, sessionId)
    const webview = {
      postMessage: async (message: unknown) => {
        await page.evaluate(
          (data) => (window as unknown as Record<string, any>).fromHost(data),
          message,
        )
        return true
      },
      onDidReceiveMessage: (listener: (message: unknown) => void) => {
        receive = listener
        return {
          dispose() {
            receive = () => {}
          },
        }
      },
    } as unknown as Webview
    host = await createHost(
      directory,
      fileURLToPath(
        new URL(
          '../../../apps/sheets/native/xlsx-engine/target/release/xlsx-sidecar',
          import.meta.url,
        ),
      ),
      webview,
    )
    const { engine, interpreter, state } = host
    const intent = {
      schemaVersion: 1,
      sessionId,
      transactionId: 'browser-commit',
      expectedRevision: 0,
      origin: 'manual',
      summary: 'Browser edit',
      operations: [{ op: 'set_cell', sheetId: state.sheets[0]!.id, address: 'B1', value: 42 }],
    }
    const receipt = await engine.submit(intent)
    expect(receipt.revision).toBe(1)
    expect(engine.state.sheets[0]!.cells.B1!.rawValue).toBe(42)
    expect(engine.projectionError).toBeDefined()
    expect(await engine.submit(intent)).toEqual(receipt)
    expect(host.edits()).toBe(1)
    await interpreter.bridge.request('view.resync', {
      descriptor: engine.descriptor,
      state: engine.state,
    })
    expect(
      await page.evaluate(
        () =>
          (window as unknown as Record<string, any>).projected.state.sheets[0].cells.B1.rawValue,
      ),
    ).toBe(42)
    await expect(
      engine.submit({
        ...intent,
        transactionId: 'failed-browser-batch',
        expectedRevision: 1,
        operations: [
          { op: 'set_cell', sheetId: state.sheets[0]!.id, address: 'B1', value: 999 },
          { op: 'set_formula', sheetId: 'missing-sheet', address: 'A1', formula: '=1' },
        ],
      }),
    ).rejects.toThrow()
    expect(engine.state.sheets[0]!.cells.B1!.rawValue).toBe(42)
    expect(host.edits()).toBe(1)
    await page.evaluate(() =>
      (window as unknown as Record<string, any>).documentBridge.request('history.request', {
        direction: 'undo',
      }),
    )
    expect(engine.state.stateId).toBe(state.stateId)
    expect(engine.descriptor.revision).toBe(2)
    await page.evaluate(() =>
      (window as unknown as Record<string, any>).documentBridge.request('history.request', {
        direction: 'redo',
      }),
    )
    expect(engine.state.sheets[0]!.cells.B1!.rawValue).toBe(42)
    expect(engine.descriptor.revision).toBe(3)
    expect(host.edits()).toBe(1)
    expect(host.historyRequests()).toEqual(['undo', 'redo'])
  } finally {
    await host?.dispose()
    await rm(directory, { recursive: true, force: true })
  }
})
