import { expect, it } from 'vitest'
import { readFile, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import JSZip from 'jszip'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { manualWorkbook } from '../../../apps/vscode-office/tests/fixtures/manual-workbook'
import { XlsxSidecarClient } from '../src/sidecar/client'
import { SidecarWorkbookSource } from '../src/sidecar/source'
import { openWorkbookSnapshot } from '../src/open'
import { initialState } from '../src/state'
import { prepareCandidateMaterial } from '../src/candidate-material'
import { SessionPackageWriter } from '../src/package-writer'
import type { WorkbookState } from '../src/domain'
import { useSessionTestResources } from './helpers'
const resources = useSessionTestResources()

it('saves the calculation mode into calcPr and reads it back', { timeout: 120_000 }, async () => {
  const directory = await resources.directory(),
    path = join(directory, 'calc.xlsx')
  await writeFile(path, await manualWorkbook())
  const client = new XlsxSidecarClient(
    fileURLToPath(
      new URL(
        '../../../apps/sheets/native/xlsx-engine/target/release/xlsx-sidecar' +
          (process.platform === 'win32' ? '.exe' : ''),
        import.meta.url,
      ),
    ),
  )
  try {
    const source = new SidecarWorkbookSource(client)
    const run = async (state: WorkbookState, sessionId: string, operations: WorkbookOperation[]) =>
      computeCandidateState(
        state,
        {
          schemaVersion: 1,
          sessionId: 'calc',
          transactionId: 'calc-tx',
          expectedRevision: 0,
          origin: 'manual',
          summary: 'Calculation',
          operations,
        },
        await prepareCandidateMaterial({ state, operations, source, sourceSessionId: sessionId }),
      )
    const opened = await openWorkbookSnapshot({
      uri: pathToFileURL(path).href,
      storageDirectory: directory,
      source,
    })
    const savedPath = join(directory, 'manual-calc.xlsx')
    try {
      const next = await run(initialState(opened.metadata), opened.original.sidecarSessionId, [
        { op: 'set_calculation_mode', mode: 'manual' },
      ])
      const staged = await new SessionPackageWriter(client, opened.metadata).stage(
        opened.original.path,
        next.nextState,
        join(directory, 'stage.xlsx'),
      )
      const xml = await (
        await JSZip.loadAsync(await readFile(staged.stagedPath))
      )
        .file('xl/workbook.xml')!
        .async('string')
      expect(xml).toMatch(/<calcPr\b[^>]*calcMode="manual"/)
      await writeFile(savedPath, await readFile(staged.stagedPath))
    } finally {
      await opened.dispose()
    }
    const reopened = await openWorkbookSnapshot({
      uri: pathToFileURL(savedPath).href,
      storageDirectory: directory,
      source,
    })
    try {
      const state = initialState(reopened.metadata)
      expect((state.features?.originalMetadata as { calcMode?: string }).calcMode).toBe('manual')
      const back = await run(state, reopened.original.sidecarSessionId, [
        { op: 'set_calculation_mode', mode: 'auto' },
      ])
      const staged = await new SessionPackageWriter(client, reopened.metadata).stage(
        reopened.original.path,
        back.nextState,
        join(directory, 'stage-2.xlsx'),
      )
      const xml = await (
        await JSZip.loadAsync(await readFile(staged.stagedPath))
      )
        .file('xl/workbook.xml')!
        .async('string')
      expect(xml).not.toMatch(/calcMode=/)
    } finally {
      await reopened.dispose()
    }
  } finally {
    client.stop()
  }
})
