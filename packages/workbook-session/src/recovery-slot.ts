import { randomUUID } from 'node:crypto'
import { basename, dirname, join } from 'node:path'
import { readFile, realpath, rm, stat, readdir } from 'node:fs/promises'
import { z } from 'zod'
import { FileResourceStore, resourceManifestSchema } from './store'
import { LIMITS, WorkbookError } from './domain'
import { readRecoveryPoint } from './recovery'

const namePattern = /^backup-slot-[0-9a-f-]{36}$/
const mutations = new Map<string, Promise<unknown>>()
async function mutateSlot<T>(path: string, action: () => Promise<T>): Promise<T> {
  const operation = (mutations.get(path) ?? Promise.resolve()).catch(() => {}).then(action)
  mutations.set(path, operation)
  try {
    return await operation
  } finally {
    if (mutations.get(path) === operation) mutations.delete(path)
  }
}
const slotSchema = z.strictObject({
  kind: z.literal('recovery-slot'),
  generation: z.string().uuid(),
  documentUri: z.string(),
  manifestPath: z.string(),
  manifestHash: z.string().regex(/^[0-9a-f]{64}$/),
})

/** Native serialized editor inputs can retain an older ID after a crash.
 * The stable slot always points to the last fully acknowledged immutable copy. */
export async function publishRecoverySlot(input: {
  storageDirectory: string
  documentUri: string
  slotPath?: string
  point: { manifestPath: string; manifestHash: string; dispose(): Promise<void> }
}) {
  const storage = await realpath(input.storageDirectory)
  const name = input.slotPath ? basename(input.slotPath, '.json') : `backup-slot-${randomUUID()}`
  const path = join(storage, 'manifests', `${name}.json`)
  if (!namePattern.test(name) || (input.slotPath && input.slotPath !== path))
    throw new WorkbookError('recovery_invalid', 'The backup slot is outside extension storage.')
  const store = await FileResourceStore.create(storage)
  const generation = randomUUID()
  await mutateSlot(path, () =>
    store.publishManifest(name, {
      schemaVersion: 1,
      references: [],
      payload: {
        kind: 'recovery-slot',
        generation,
        documentUri: input.documentUri,
        manifestPath: input.point.manifestPath,
        manifestHash: input.point.manifestHash,
      },
    }),
  )
  return {
    path,
    async dispose() {
      // Deleting an obsolete native backup must not delete its replacement.
      await mutateSlot(path, async () => {
        const current = await readFile(path).catch(() => undefined)
        if (current) {
          const slot = slotSchema.parse(
            resourceManifestSchema.parse(JSON.parse(current.toString())).payload,
          )
          if (slot.generation === generation) await rm(path, { force: true })
        }
      })
      await input.point.dispose()
    },
  }
}

export async function readRecoverySlot(input: {
  slotPath: string
  documentUri: string
  allowedRoots: readonly string[]
}) {
  try {
    const path = await realpath(input.slotPath)
    if (!namePattern.test(basename(path, '.json')) || basename(dirname(path)) !== 'manifests')
      throw new Error('Invalid backup slot name')
    const storage = await realpath(dirname(dirname(path)))
    const roots = await Promise.all(input.allowedRoots.map((root) => realpath(root)))
    if (!roots.includes(storage)) throw new Error('Backup slot is outside extension storage')
    if ((await stat(path)).size > LIMITS.manifestBytes)
      throw new Error('Backup slot exceeds its limit')
    const bytes = await readFile(path)
    if (bytes.length > LIMITS.manifestBytes) throw new Error('Backup slot exceeds its limit')
    const slot = slotSchema.parse(
      resourceManifestSchema.parse(JSON.parse(bytes.toString())).payload,
    )
    if (slot.documentUri !== input.documentUri)
      throw new Error('Backup slot belongs to another document')
    return await readRecoveryPoint({
      manifestPath: slot.manifestPath,
      expectedManifestHash: slot.manifestHash,
      documentUri: input.documentUri,
      allowedRoots: roots,
    })
  } catch (error) {
    throw new WorkbookError(
      'recovery_invalid',
      error instanceof Error ? error.message : String(error),
    )
  }
}

/** Lists only small private metadata; full copies are verified when selected. */
export async function listRecoverySlots(storageDirectory: string) {
  const directory = join(storageDirectory, 'manifests')
  const entries = await readdir(directory).catch((error: NodeJS.ErrnoException) => {
    if (error.code === 'ENOENT') return []
    throw error
  })
  const results: { path: string; documentUri: string; modifiedAt: number }[] = []
  for (const entry of entries) {
    if (!namePattern.test(basename(entry, '.json'))) continue
    const path = join(directory, entry),
      metadata = await stat(path)
    if (!metadata.isFile() || metadata.size > LIMITS.manifestBytes) continue
    if ((await realpath(path)) !== path) continue
    try {
      const bytes = await readFile(path)
      if (bytes.length > LIMITS.manifestBytes) continue
      const parsed = slotSchema.safeParse(
        resourceManifestSchema.parse(JSON.parse(bytes.toString())).payload,
      )
      if (parsed.success)
        results.push({ path, documentUri: parsed.data.documentUri, modifiedAt: metadata.mtimeMs })
    } catch {
      /* A corrupt candidate must not hide other valid recovery copies. */
    }
  }
  return results.sort((a, b) => b.modifiedAt - a.modifiedAt)
}
