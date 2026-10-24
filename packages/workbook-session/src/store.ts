import { createHash, randomUUID } from 'node:crypto'
import { mkdir, open, readFile, readdir, rename, rm } from 'node:fs/promises'
import { join } from 'node:path'
import { z } from 'zod'
import { LIMITS, WorkbookError } from './domain'
import type { ResourceStore } from './ports'

const digestSchema = z.string().regex(/^[a-f0-9]{64}$/)
const nameSchema = z.string().regex(/^[a-zA-Z0-9_-]{1,128}$/)
export const resourceManifestSchema = z.strictObject({
  schemaVersion: z.literal(1),
  references: z.array(digestSchema),
  payload: z.json(),
})
const digest = (data: Uint8Array) => createHash('sha256').update(data).digest('hex')

export class FileResourceStore implements ResourceStore {
  private constructor(
    readonly directory: string,
    private readonly fault?: (point: string) => void,
  ) {}
  static async create(directory: string, fault?: (point: string) => void) {
    await mkdir(join(directory, 'chunks'), { recursive: true, mode: 0o700 })
    await mkdir(join(directory, 'manifests'), { recursive: true, mode: 0o700 })
    return new FileResourceStore(directory, fault)
  }
  async writeChunk(data: Uint8Array) {
    if (data.byteLength > LIMITS.chunkBytes)
      throw new WorkbookError('chunk_limit', 'A chunk exceeds 256 KiB.')
    const reference = digest(data),
      path = join(this.directory, 'chunks', reference)
    try {
      await this.readChunk(reference)
      return reference
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error
    }
    await this.atomicWrite(path, data)
    return reference
  }
  async readChunk(reference: string) {
    digestSchema.parse(reference)
    const bytes = await readFile(join(this.directory, 'chunks', reference))
    if (bytes.length > LIMITS.chunkBytes || digest(bytes) !== reference)
      throw new WorkbookError('resource_corrupt', 'The chunk digest does not match its reference.')
    return bytes
  }
  async publishManifest(name: string, data: unknown) {
    nameSchema.parse(name)
    const manifest = resourceManifestSchema.parse(data)
    const bytes = Buffer.from(JSON.stringify(manifest))
    if (bytes.length > LIMITS.manifestBytes)
      throw new WorkbookError('manifest_limit', 'The manifest exceeds 256 KiB.')
    for (const reference of manifest.references) await this.readChunk(reference)
    await this.atomicWrite(
      join(this.directory, 'manifests', `${name}.json`),
      bytes,
      'beforeManifestRename',
    )
  }
  async readManifest(name: string) {
    const manifest = await this.readManifestMetadata(name)
    for (const reference of manifest.references) await this.readChunk(reference)
    return manifest
  }
  async readManifestMetadata(name: string) {
    nameSchema.parse(name)
    const bytes = await readFile(join(this.directory, 'manifests', `${name}.json`))
    if (bytes.length > LIMITS.manifestBytes)
      throw new WorkbookError('manifest_limit', 'The manifest exceeds 256 KiB.')
    const manifest = resourceManifestSchema.parse(JSON.parse(bytes.toString('utf8')))
    return manifest
  }
  async delete(reference: string) {
    digestSchema.parse(reference)
    await rm(join(this.directory, 'chunks', reference), { force: true })
  }
  async collect(retainedManifestNames: readonly string[]) {
    const retained = new Set<string>()
    for (const name of retainedManifestNames)
      for (const reference of (await this.readManifestMetadata(name)).references)
        retained.add(reference)
    for (const name of await readdir(join(this.directory, 'chunks')))
      if (digestSchema.safeParse(name).success && !retained.has(name)) await this.delete(name)
    const retainedNames = new Set(retainedManifestNames.map((name) => `${name}.json`))
    for (const name of await readdir(join(this.directory, 'manifests')))
      if (/^history-[a-zA-Z0-9_-]+\.json$/.test(name) && !retainedNames.has(name))
        await rm(join(this.directory, 'manifests', name), { force: true })
  }
  private async atomicWrite(path: string, bytes: Uint8Array, point?: string) {
    const temporary = `${path}.${randomUUID()}.tmp`
    let handle
    try {
      handle = await open(temporary, 'wx', 0o600)
      await handle.writeFile(bytes)
      await handle.sync()
      await handle.close()
      handle = undefined
      if (point) this.fault?.(point)
      await rename(temporary, path)
      // Directory fsync makes the rename durable on supported local filesystems.
      const parent = await open(join(path, '..'), 'r').catch(() => undefined)
      if (parent)
        try {
          await parent.sync()
        } catch {
          /* Some platforms cannot fsync directories. */
        } finally {
          await parent.close().catch(() => {})
        }
    } finally {
      await handle?.close()
      await rm(temporary, { force: true }).catch(() => {})
    }
  }
}
