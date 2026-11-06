import { constants } from 'node:fs'
import { chmod, copyFile, link, open, rename, stat, unlink } from 'node:fs/promises'
import { basename, dirname, join, resolve } from 'node:path'
import { randomUUID } from 'node:crypto'
import { WorkbookError } from './domain'
import { hashFile } from './resources'
import { resolveDocumentPath } from './document-path'

export interface FileBaseline {
  requestedPath: string
  canonicalPath: string
  sha256: string | null
  physicalId: string | null
  size: string | null
  mtimeNs: string | null
  ctimeNs: string | null
  mode: number
}

/** Hashes a stable file version; identity and metadata surround the read. */
export async function captureFileBaseline(path: string): Promise<FileBaseline> {
  const requestedPath = resolve(path)
  let canonicalPath: string
  try {
    canonicalPath = await resolveDocumentPath(requestedPath)
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error
    const parent = await resolveDocumentPath(dirname(requestedPath)).catch(
      (missing: NodeJS.ErrnoException) => {
        if (missing.code === 'ENOENT') return dirname(requestedPath)
        throw missing
      },
    )
    canonicalPath = join(parent, basename(requestedPath))
    // Check again: a destination may have appeared while resolving its parent.
    try {
      await stat(requestedPath)
    } catch (missing) {
      if ((missing as NodeJS.ErrnoException).code === 'ENOENT')
        return {
          requestedPath,
          canonicalPath,
          sha256: null,
          physicalId: null,
          size: null,
          mtimeNs: null,
          ctimeNs: null,
          mode: 0o600,
        }
      throw missing
    }
    return captureFileBaseline(requestedPath)
  }
  const before = await stat(canonicalPath, { bigint: true })
  if (!before.isFile())
    throw new WorkbookError('unsupported_resource', 'The save target must be a regular file.')
  const sha256 = await hashFile(canonicalPath),
    after = await stat(canonicalPath, { bigint: true })
  if (
    before.dev !== after.dev ||
    before.ino !== after.ino ||
    before.size !== after.size ||
    before.mtimeNs !== after.mtimeNs ||
    before.ctimeNs !== after.ctimeNs
  )
    throw new WorkbookError(
      'external_change_conflict',
      'The save target changed while its baseline was being read.',
      {},
      true,
    )
  return {
    requestedPath,
    canonicalPath,
    sha256,
    physicalId: `${after.dev}:${after.ino}`,
    size: String(after.size),
    mtimeNs: String(after.mtimeNs),
    ctimeNs: String(after.ctimeNs),
    mode: Number(after.mode & 0o777n),
  }
}

export function sameFileBaseline(a: FileBaseline, b: FileBaseline) {
  return (
    a.canonicalPath === b.canonicalPath &&
    a.sha256 === b.sha256 &&
    a.physicalId === b.physicalId &&
    a.size === b.size &&
    a.mtimeNs === b.mtimeNs &&
    a.ctimeNs === b.ctimeNs
  )
}

/** Filesystem checks detect races; external writers do not participate in a CAS. */
export async function publishStagedWorkbook(input: {
  targetPath: string
  stagedPath: string
  baseline: FileBaseline
  fault?(point: 'beforePublish' | 'afterPublish'): void | Promise<void>
}) {
  const { baseline, stagedPath } = input,
    stagedHash = await hashFile(stagedPath)
  const conflict = (message: string, details: Record<string, unknown> = {}) =>
    new WorkbookError('external_change_conflict', message, {
      targetPath: input.targetPath,
      stagedPath,
      ...details,
    })
  if (!sameFileBaseline(await captureFileBaseline(input.targetPath), baseline))
    throw conflict('The save target changed; the staged workbook has been preserved.')
  const retainedPath = join(dirname(stagedPath), `candidate-${randomUUID()}.xlsx`),
    previousPath =
      baseline.sha256 === null
        ? undefined
        : join(dirname(stagedPath), `previous-${randomUUID()}.xlsx`)
  let published = false
  try {
    // Copies must have separate inodes: a later external in-place write must
    // not change the retained candidate or previous package as well.
    await copyFile(stagedPath, retainedPath, constants.COPYFILE_EXCL)
    await chmod(retainedPath, 0o600)
    if (previousPath) {
      await copyFile(baseline.canonicalPath, previousPath, constants.COPYFILE_EXCL)
      await chmod(previousPath, 0o600)
      if ((await hashFile(previousPath)) !== baseline.sha256)
        throw conflict('The previous package changed while creating its safety copy.', {
          previousPath,
        })
    }
    await input.fault?.('beforePublish')
    if (!sameFileBaseline(await captureFileBaseline(input.targetPath), baseline))
      throw conflict(
        'The save target changed before publication.',
        previousPath ? { previousPath } : {},
      )
    if (
      (await hashFile(stagedPath)) !== stagedHash ||
      (await hashFile(retainedPath)) !== stagedHash
    )
      throw new WorkbookError(
        'save_failed',
        'The private staged package changed before publication.',
      )
    await chmod(stagedPath, 0o600)
    const handle = await open(stagedPath, 'r+')
    try {
      await handle.sync()
    } finally {
      await handle.close()
    }
    await chmod(stagedPath, baseline.mode)
    const stagedIdentity = await stat(stagedPath, { bigint: true })
    if (baseline.sha256 === null) {
      await link(stagedPath, baseline.canonicalPath)
      published = true
      await unlink(stagedPath).catch(() => {})
    } else {
      for (let attempt = 0; ; attempt++) {
        if (!sameFileBaseline(await captureFileBaseline(input.targetPath), baseline))
          throw conflict(
            'The save target changed before replacement.',
            previousPath ? { previousPath } : {},
          )
        try {
          await rename(stagedPath, baseline.canonicalPath)
          break
        } catch (error) {
          if (
            attempt >= 3 ||
            !['EPERM', 'EACCES', 'EBUSY'].includes((error as NodeJS.ErrnoException).code ?? '')
          )
            throw error
          await new Promise((resolve) => setTimeout(resolve, 50 * 2 ** attempt))
        }
      }
      published = true
    }
    await input.fault?.('afterPublish')
    const next = await captureFileBaseline(input.targetPath)
    if (
      next.canonicalPath !== baseline.canonicalPath ||
      next.sha256 !== stagedHash ||
      next.physicalId !== `${stagedIdentity.dev}:${stagedIdentity.ino}`
    )
      throw conflict(
        'The destination changed after publication; inspect the retained previous and candidate packages.',
        { published: true, ...(previousPath ? { previousPath } : {}), stagedPath: retainedPath },
      )
    await unlink(retainedPath).catch(() => {})
    if (previousPath) await unlink(previousPath).catch(() => {})
    return next
  } catch (error) {
    if (error instanceof WorkbookError) throw error
    throw new WorkbookError('save_failed', error instanceof Error ? error.message : String(error), {
      published,
      stagedPath: retainedPath,
      ...(previousPath ? { previousPath } : {}),
      targetPath: input.targetPath,
    })
  }
}
