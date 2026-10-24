import { createHash } from 'node:crypto'
import { chmod, copyFile, open, readFile, realpath, rm, stat } from 'node:fs/promises'
import { basename, dirname, isAbsolute, join, relative, sep } from 'node:path'
import { z } from 'zod'
import { LIMITS, WorkbookError, assertNotCancelled, type WorkbookState } from './domain'
import { FileResourceStore, resourceManifestSchema } from './store'
import { hashFile, privateDirectory } from './resources'
import { workbookStateSchema } from './schemas'

const digest = (bytes: Uint8Array) => createHash('sha256').update(bytes).digest('hex')
const payloadSchema = z.strictObject({
  kind: z.literal('workbook-backup'),
  originalUri: z.string(),
  createdAt: z.number().finite(),
  baseHash: z.string().regex(/^[a-f0-9]{64}$/),
  stateId: z.string(),
  savedStateId: z.string(),
  baselineHash: z.string().nullable(),
  bytes: z.number().int().min(1).max(LIMITS.transactionBytes),
  digest: z.string().regex(/^[a-f0-9]{64}$/),
  chunks: z.array(z.string()),
})

export async function createRecoveryPoint(input: {
  storageDirectory: string
  original: { path: string; sha256: string }
  documentUri: string
  state: Readonly<WorkbookState>
  savedStateId: string
  baselineHash: string | null
  signal?: AbortSignal
  byteLimit?: number
  fault?(
    point: 'afterBaseCopy' | 'afterChunkWrite' | 'beforeManifestPublish' | 'afterManifestPublish',
  ): void | Promise<void>
}) {
  assertNotCancelled(input.signal)
  const state = workbookStateSchema.parse(input.state),
    bytes = Buffer.from(JSON.stringify(state)),
    requestedLimit = input.byteLimit ?? LIMITS.transactionBytes
  if (!Number.isSafeInteger(requestedLimit) || requestedLimit <= 0)
    throw new WorkbookError('undo_budget_exceeded', 'The backup byte limit is invalid.')
  if (bytes.length > Math.min(requestedLimit, LIMITS.transactionBytes))
    throw new WorkbookError(
      'undo_budget_exceeded',
      'The complete backup exceeds its byte budget; no recovery point was acknowledged.',
    )
  const directory = await privateDirectory(input.storageDirectory, 'backup-')
  try {
    const basePath = join(directory, 'base.xlsx')
    await copyFile(input.original.path, basePath)
    // Windows requires write access for FlushFileBuffers; copied sources may be read-only.
    await chmod(basePath, 0o600)
    const baseHandle = await open(basePath, 'r+')
    try {
      await baseHandle.sync()
    } finally {
      await baseHandle.close()
    }
    await chmod(basePath, 0o400)
    if ((await hashFile(basePath)) !== input.original.sha256)
      throw new WorkbookError('recovery_invalid', 'The immutable backup source digest changed.')
    await input.fault?.('afterBaseCopy')
    const store = await FileResourceStore.create(directory),
      chunks: string[] = []
    for (let offset = 0; offset < bytes.length; offset += LIMITS.chunkBytes) {
      assertNotCancelled(input.signal)
      chunks.push(await store.writeChunk(bytes.subarray(offset, offset + LIMITS.chunkBytes)))
      await input.fault?.('afterChunkWrite')
    }
    await input.fault?.('beforeManifestPublish')
    assertNotCancelled(input.signal)
    await store.publishManifest('backup', {
      schemaVersion: 1,
      references: chunks,
      payload: {
        kind: 'workbook-backup',
        originalUri: input.documentUri,
        createdAt: Date.now(),
        baseHash: input.original.sha256,
        stateId: state.stateId,
        savedStateId: input.savedStateId,
        baselineHash: input.baselineHash,
        bytes: bytes.length,
        digest: digest(bytes),
        chunks,
      },
    })
    await input.fault?.('afterManifestPublish')
    const manifestPath = join(directory, 'manifests', 'backup.json'),
      manifestHash = digest(await readFile(manifestPath))
    return {
      directory,
      manifestPath,
      manifestHash,
      dispose: () => rm(directory, { recursive: true, force: true }),
    }
  } catch (error) {
    await rm(directory, { recursive: true, force: true })
    throw error
  }
}

export async function readRecoveryPoint(input: {
  manifestPath: string
  expectedManifestHash: string
  documentUri: string
  allowedRoots: readonly string[]
}) {
  try {
    if (!/^[a-f0-9]{64}$/.test(input.expectedManifestHash))
      throw new Error('Invalid manifest digest')
    const manifestPath = await realpath(input.manifestPath)
    const roots = await Promise.all(
      input.allowedRoots.map((root) => realpath(root).catch(() => undefined)),
    )
    if (
      !roots.some((root) => {
        if (!root) return false
        const child = relative(root, manifestPath)
        return child !== '' && child !== '..' && !child.startsWith(`..${sep}`) && !isAbsolute(child)
      }) ||
      basename(manifestPath) !== 'backup.json' ||
      basename(dirname(manifestPath)) !== 'manifests'
    )
      throw new Error('The backup is outside extension storage')
    if ((await stat(manifestPath)).size > LIMITS.manifestBytes)
      throw new Error('The backup manifest is too large')
    const manifestBytes = await readFile(manifestPath)
    if (digest(manifestBytes) !== input.expectedManifestHash)
      throw new Error('The acknowledged manifest digest does not match')
    const manifest = resourceManifestSchema.parse(JSON.parse(manifestBytes.toString('utf8'))),
      payload = payloadSchema.parse(manifest.payload),
      directory = dirname(dirname(manifestPath)),
      basePath = join(directory, 'base.xlsx')
    if ((await realpath(basePath)) !== basePath)
      throw new Error('The backup base must not be a symbolic link')
    if (
      payload.originalUri !== input.documentUri ||
      JSON.stringify(payload.chunks) !== JSON.stringify(manifest.references) ||
      (await hashFile(basePath)) !== payload.baseHash
    )
      throw new Error('The backup source or references do not match')
    const store = await FileResourceStore.create(directory),
      parts: Uint8Array[] = []
    let total = 0
    for (const reference of payload.chunks) {
      const part = await store.readChunk(reference)
      total += part.length
      if (total > payload.bytes) throw new Error('The backup exceeds its declared size')
      parts.push(part)
    }
    const bytes = Buffer.concat(parts)
    if (bytes.length !== payload.bytes || digest(bytes) !== payload.digest)
      throw new Error('The backup content is incomplete')
    const state = workbookStateSchema.parse(JSON.parse(bytes.toString('utf8')))
    if (state.stateId !== payload.stateId)
      throw new Error('The backup state identity does not match')
    return {
      state,
      savedStateId: payload.savedStateId,
      baselineHash: payload.baselineHash,
      basePath,
    }
  } catch (error) {
    throw new WorkbookError(
      'recovery_invalid',
      error instanceof Error ? error.message : String(error),
    )
  }
}
