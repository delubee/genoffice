import { expect, it } from 'vitest'
import { writeFile, readFile, readdir } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import { join } from 'node:path'
import { manualWorkbook } from './fixtures/manual-workbook'
import { withTestWorkspace } from './helpers'
import { XlsxSidecarClient } from '@genoffice/workbook-session/sidecar/client'
import { SidecarWorkbookSource } from '@genoffice/workbook-session/sidecar/source'
import { openWorkbookSnapshot } from '@genoffice/workbook-session/open'
import { initialState } from '@genoffice/workbook-session/state'
import { SessionPackageWriter } from '@genoffice/workbook-session/package-writer'
import type { ArchiveClient } from '@genoffice/xlsx-gateway/gateway/xlsx-package-io'

async function fixture(
  run: (context: {
    directory: string
    path: string
    original: Buffer
    opened: Awaited<ReturnType<typeof openWorkbookSnapshot>>
    client: XlsxSidecarClient
  }) => Promise<void>,
) {
  await withTestWorkspace(async (directory) => {
    const client = new XlsxSidecarClient(
        fileURLToPath(
          new URL('../../sheets/native/xlsx-engine/target/release/xlsx-sidecar', import.meta.url),
        ),
      ),
      source = new SidecarWorkbookSource(client),
      path = join(directory, 'input.xlsx'),
      original = await manualWorkbook()
    await writeFile(path, original)
    try {
      const opened = await openWorkbookSnapshot({
        uri: new URL(`file://${path}`).href,
        storageDirectory: directory,
        source,
      })
      try {
        await run({ directory, path, original, opened, client })
      } finally {
        await opened.dispose()
      }
    } finally {
      client.stop()
    }
  })
}

it('rejects damaged workbook bytes and removes the incomplete private snapshot', async () => {
  await withTestWorkspace(async (directory) => {
    const client = new XlsxSidecarClient(
        fileURLToPath(
          new URL('../../sheets/native/xlsx-engine/target/release/xlsx-sidecar', import.meta.url),
        ),
      ),
      source = new SidecarWorkbookSource(client),
      path = join(directory, 'damaged.xlsx')
    await writeFile(path, 'not an xlsx archive')
    try {
      await expect(
        openWorkbookSnapshot({
          uri: new URL(`file://${path}`).href,
          storageDirectory: directory,
          source,
        }),
      ).rejects.toThrow()
      expect(await readdir(directory)).toEqual(['damaged.xlsx'])
      expect(await readFile(path, 'utf8')).toBe('not an xlsx archive')
    } finally {
      client.stop()
    }
  })
})

it('leaves both source and target intact on disk-full staging failure and cleans private stages', async () => {
  await fixture(async ({ directory, path, original, opened, client }) => {
    const state = structuredClone(initialState(opened.metadata)),
      target = join(directory, 'copy.xlsx')
    state.sheets[0]!.cells.A2 = { kind: 'number', rawValue: 999 }
    const broken: ArchiveClient = {
      archiveManifest: client.archiveManifest.bind(client),
      readEntries: client.readEntries.bind(client),
      scanEntries: client.scanEntries.bind(client),
      async saveArchive() {
        throw Object.assign(new Error('No space left on device'), { code: 'ENOSPC' })
      },
    }
    await expect(
      new SessionPackageWriter(broken, opened.metadata).stage(opened.original.path, state, target),
    ).rejects.toMatchObject({ code: 'ENOSPC' })
    expect(await readFile(path)).toEqual(original)
    expect(
      (await readdir(directory)).filter((name) => name.startsWith('.genoffice-stage-')),
    ).toEqual([])
    await expect(readFile(target)).rejects.toMatchObject({ code: 'ENOENT' })
  })
})

it('preserves a valid staged file when the destination changes before publication', async () => {
  await fixture(async ({ directory, path, original, opened, client }) => {
    const state = structuredClone(initialState(opened.metadata)),
      writer = new SessionPackageWriter(client, opened.metadata),
      target = join(directory, 'copy.xlsx')
    state.sheets[0]!.cells.A2 = { kind: 'number', rawValue: 999 }
    const staged = await writer.stage(opened.original.path, state, target)
    await writeFile(target, 'external data')
    await expect(writer.publish(staged.stagedPath, target, null)).rejects.toMatchObject({
      code: 'external_change_conflict',
    })
    expect(await readFile(target, 'utf8')).toBe('external data')
    expect((await readFile(staged.stagedPath)).length).toBeGreaterThan(0)
    expect(await readFile(path)).toEqual(original)
  })
})

it('cancels staging without publishing or leaving partial private packages', async () => {
  await fixture(async ({ directory, path, original, opened, client }) => {
    const controller = new AbortController(),
      target = join(directory, 'copy.xlsx'),
      state = structuredClone(initialState(opened.metadata))
    state.sheets[0]!.cells.A2 = { kind: 'number', rawValue: 999 }
    const cancelling: ArchiveClient = {
      async archiveManifest(path) {
        const result = await client.archiveManifest(path)
        controller.abort()
        return result
      },
      readEntries: client.readEntries.bind(client),
      scanEntries: client.scanEntries.bind(client),
      saveArchive: client.saveArchive.bind(client),
    }
    await expect(
      new SessionPackageWriter(cancelling, opened.metadata).stage(
        opened.original.path,
        state,
        target,
        controller.signal,
      ),
    ).rejects.toMatchObject({ code: 'operation_cancelled' })
    expect(await readFile(path)).toEqual(original)
    expect(
      (await readdir(directory)).filter((name) => name.startsWith('.genoffice-stage-')),
    ).toEqual([])
    await expect(readFile(target)).rejects.toMatchObject({ code: 'ENOENT' })
  })
})
