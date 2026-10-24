import { createHash } from 'node:crypto'
import { chmod, open, rm } from 'node:fs/promises'
import { join } from 'node:path'
import { LIMITS, WorkbookError, assertNotCancelled, type WorkbookState } from '../domain'
import { privateDirectory } from '../resources'
import { openWorkspaceFile, readWorkspaceManifest, publishJson } from './paths'
import { readTypedNdjson } from './ndjson'
import { validateResultManifest, validateResultCell } from './results'
import type { SnapshotAuthority } from './snapshots'

export async function prepareResultCopy(input: {
  source: SnapshotAuthority
  resultManifestPath: string
  privateStorage: string
  state: WorkbookState
  sessionId: string
  revision: number
  signal?: AbortSignal
}) {
  assertNotCancelled(input.signal)
  const selected = await readWorkspaceManifest(input.resultManifestPath, input.source.directory)
  const manifest = validateResultManifest(
    selected.value,
    input.source,
    input.state,
    input.sessionId,
    input.revision,
  )
  const directory = await privateDirectory(input.privateStorage, 'prepared-import-'),
    dataPath = join(directory, 'result.ndjson')
  try {
    const source = await openWorkspaceFile(
      join(input.source.directory, manifest.output.fileName),
      input.source.directory,
      LIMITS.artifactBytes,
    )
    try {
      const target = await open(dataPath, 'wx', 0o600),
        hash = createHash('sha256')
      let bytes = 0
      try {
        for await (const chunk of source.handle.createReadStream({
          autoClose: false,
          highWaterMark: 64 * 1024,
        })) {
          assertNotCancelled(input.signal)
          const buffer = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk)
          bytes += buffer.length
          if (bytes > manifest.output.bytes || bytes > LIMITS.artifactBytes)
            throw new WorkbookError(
              'artifact_limit_exceeded',
              'Result data exceeds its declared byte budget.',
            )
          hash.update(buffer)
          await target.writeFile(buffer)
        }
        const after = await source.handle.stat({ bigint: true }),
          before = source.baseline
        if (
          before.size !== after.size ||
          before.mtimeNs !== after.mtimeNs ||
          before.ctimeNs !== after.ctimeNs ||
          bytes !== manifest.output.bytes ||
          hash.digest('hex') !== manifest.output.sha256
        )
          throw new WorkbookError(
            'artifact_hash_mismatch',
            'The result changed during staging or does not match its declared digest.',
          )
        await target.sync()
      } finally {
        await target.close()
      }
    } finally {
      await source.handle.close()
    }
    await chmod(dataPath, 0o400)
    let count = 0,
      indexedCells = 0
    const batches: { byteOffset: number; cellOffset: number }[] = []
    for await (const record of readTypedNdjson({
      path: dataPath,
      shape: manifest.shape,
      expected: manifest.output,
      signal: input.signal,
      onRecordPosition(record, byteOffset) {
        for (
          let start = Math.ceil(indexedCells / LIMITS.importBatchCells) * LIMITS.importBatchCells;
          start < indexedCells + record.cells.length;
          start += LIMITS.importBatchCells
        )
          batches.push({ byteOffset, cellOffset: start - indexedCells })
        indexedCells += record.cells.length
      },
    }))
      for (const cell of record.cells) {
        validateResultCell(cell, manifest, input.state.dateSystem)
        count++
      }
    assertNotCancelled(input.signal)
    await publishJson(join(directory, 'result.manifest.json'), manifest, input.signal)
    return {
      directory,
      dataPath,
      manifest,
      batches,
      resultManifestSha256: createHash('sha256').update(selected.bytes).digest('hex'),
      cellCount: count,
      dispose: () => rm(directory, { recursive: true, force: true }),
    }
  } catch (error) {
    await rm(directory, { recursive: true, force: true })
    throw error
  }
}
