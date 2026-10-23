import { constants } from 'node:fs'
import { mkdir, open, realpath, stat, rename, rm } from 'node:fs/promises'
import { createHash, randomUUID } from 'node:crypto'
import { dirname, isAbsolute, join, relative, sep } from 'node:path'
import { LIMITS, WorkbookError, assertNotCancelled } from '../domain'

export function isInside(root: string, path: string) {
  const child = relative(root, path)
  return child !== '' && child !== '..' && !child.startsWith(`..${sep}`) && !isAbsolute(child)
}
async function assertWorkspaceParent(root: string, path: string) {
  let ancestor = path
  for (;;) {
    try {
      const canonical = await realpath(ancestor)
      if (canonical !== root && !isInside(root, canonical))
        throw new WorkbookError('path_not_allowed', 'The exchange path escapes its workspace root.')
      return
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error
      const parent = dirname(ancestor)
      if (parent === ancestor)
        throw new WorkbookError('path_not_allowed', 'No workspace parent exists.')
      ancestor = parent
    }
  }
}
export async function workspaceTaskDirectory(workspaceRoot: string, snapshotId: string) {
  const root = await realpath(workspaceRoot)
  if (!(await stat(root)).isDirectory())
    throw new WorkbookError('path_not_allowed', 'Choose a real workspace directory.')
  const work = join(root, '.genoffice', 'work')
  await assertWorkspaceParent(root, work)
  await mkdir(work, { recursive: true, mode: 0o700 })
  const canonicalWork = await realpath(work)
  if (!isInside(root, canonicalWork))
    throw new WorkbookError(
      'path_not_allowed',
      'The workspace exchange directory escapes its root.',
    )
  const directory = join(canonicalWork, snapshotId)
  await mkdir(directory, { mode: 0o700 })
  return { root, directory: await realpath(directory) }
}

/** Opens one verified file descriptor; later pathname replacement cannot swap the input. */
export async function openWorkspaceFile(path: string, allowedDirectory: string, limit: number) {
  const root = await realpath(allowedDirectory)
  if (root !== allowedDirectory)
    throw new WorkbookError(
      'path_not_allowed',
      'The registered task directory was replaced or redirected; export a new snapshot.',
    )
  const canonical = await realpath(path)
  if (!isInside(root, canonical) || dirname(canonical) !== root)
    throw new WorkbookError(
      'path_not_allowed',
      'Use a regular result file inside this snapshot task directory.',
    )
  const handle = await open(canonical, constants.O_RDONLY | (constants.O_NOFOLLOW ?? 0))
  try {
    const file = await handle.stat({ bigint: true }),
      current = await stat(canonical, { bigint: true })
    if (
      !file.isFile() ||
      file.size > BigInt(limit) ||
      file.dev !== current.dev ||
      file.ino !== current.ino ||
      !isInside(root, await realpath(path))
    )
      throw new WorkbookError(
        'path_not_allowed',
        'The selected file changed identity, is not regular or exceeds its size limit.',
      )
    return { handle, canonical, baseline: file }
  } catch (error) {
    await handle.close()
    throw error
  }
}

export async function readWorkspaceManifest(path: string, directory: string) {
  const opened = await openWorkspaceFile(path, directory, LIMITS.manifestBytes)
  try {
    const bytes = await opened.handle.readFile()
    if (bytes.length > LIMITS.manifestBytes)
      throw new WorkbookError('artifact_limit_exceeded', 'Manifest exceeds 256 KiB.')
    let value: unknown
    try {
      value = JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(bytes))
    } catch {
      throw new WorkbookError('type_mismatch', 'Manifest must be valid UTF-8 JSON.')
    }
    return { value, bytes }
  } finally {
    await opened.handle.close()
  }
}

export async function publishJson(path: string, value: unknown, signal?: AbortSignal) {
  const bytes = Buffer.from(JSON.stringify(value))
  if (bytes.length > LIMITS.manifestBytes)
    throw new WorkbookError('artifact_limit_exceeded', 'Manifest exceeds 256 KiB.')
  const temporary = `${path}.${randomUUID()}.tmp`,
    handle = await open(temporary, 'wx', 0o600)
  try {
    await handle.writeFile(bytes)
    await handle.sync()
    await handle.close()
    assertNotCancelled(signal)
    await rename(temporary, path)
    const parent = await open(dirname(path), 'r').catch(() => undefined)
    try {
      await parent?.sync()
    } catch {
      /* Directory sync is not supported on every filesystem. */
    } finally {
      await parent?.close()
    }
    return createHash('sha256').update(bytes).digest('hex')
  } finally {
    await handle.close().catch(() => {})
    await rm(temporary, { force: true }).catch(() => {})
  }
}
