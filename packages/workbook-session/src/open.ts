import { constants } from 'node:fs'
import { copyFile, realpath, stat, chmod } from 'node:fs/promises'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { assertNotCancelled, WorkbookError, type OriginalBase } from './domain'
import type { WorkbookSource } from './ports'
import { hashFile, privateDirectory, RetainedResource } from './resources'

export async function openWorkbookSnapshot(input: {
  uri: string
  storageDirectory: string
  source: WorkbookSource
  signal?: AbortSignal
}) {
  assertNotCancelled(input.signal)
  const url = new URL(input.uri)
  if (url.protocol !== 'file:')
    throw new WorkbookError(
      'unsupported_resource',
      'Only a file on the extension host is supported.',
    )
  const path = fileURLToPath(url),
    canonicalPath = await realpath(path),
    physical = await stat(canonicalPath)
  if (!physical.isFile())
    throw new WorkbookError('unsupported_resource', 'The workbook must be a regular file.')
  const directory = await privateDirectory(input.storageDirectory, 'base-')
  let sidecarId: string | undefined
  const owner = new RetainedResource(directory, async () => {
    if (sidecarId) await input.source.close(sidecarId)
  })
  try {
    const snapshotPath = join(directory, 'original.xlsx')
    await copyFile(canonicalPath, snapshotPath, constants.COPYFILE_EXCL)
    await chmod(snapshotPath, 0o400)
    assertNotCancelled(input.signal)
    const sha256 = await hashFile(snapshotPath)
    const opened = await input.source.openSnapshot(snapshotPath)
    sidecarId = opened.sessionId
    const manifest = await input.source.archiveManifest(snapshotPath)
    assertNotCancelled(input.signal)
    const original: OriginalBase = Object.freeze({
      path: snapshotPath,
      sha256,
      sourceUri: input.uri,
      sidecarSessionId: sidecarId,
    })
    return {
      original,
      metadata:
        opened.metadata && typeof opened.metadata === 'object'
          ? { ...opened.metadata, readOnly: (physical.mode & 0o222) === 0 }
          : opened.metadata,
      archiveManifest: manifest,
      canonicalPath,
      physicalId: `${physical.dev}:${physical.ino}`,
      retain: () => owner.retain(),
      dispose: () => owner.dispose(),
    }
  } catch (error) {
    await owner.dispose()
    throw error
  }
}
