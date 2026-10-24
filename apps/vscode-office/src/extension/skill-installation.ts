import { createHash, randomUUID } from 'node:crypto'
import {
  lstat,
  mkdir,
  mkdtemp,
  readFile,
  readdir,
  realpath,
  rename,
  rm,
  writeFile,
} from 'node:fs/promises'
import { dirname, join, relative, sep } from 'node:path'
import { z } from 'zod'
import { WorkbookError } from '@genoffice/workbook-session/domain'
import { isInside } from '@genoffice/workbook-session/exchange/paths'

const RECEIPT = '.genoffice-install.json'
const sha = (bytes: Uint8Array | string) => createHash('sha256').update(bytes).digest('hex')
const receiptSchema = z.strictObject({
  schemaVersion: z.literal(1),
  version: z.string(),
  installedDigest: z.string(),
  bundleDigest: z.string(),
  installedAt: z.number(),
  fileDigests: z.record(z.string(), z.string()),
})
type Tree = Map<string, Buffer>
async function tree(directory: string): Promise<Tree> {
  const files: Tree = new Map()
  let total = 0
  async function walk(path: string) {
    for (const entry of await readdir(path, { withFileTypes: true })) {
      if (entry.name === RECEIPT && path === directory) continue
      const child = join(path, entry.name)
      if (entry.isSymbolicLink())
        throw new WorkbookError('skill_path_not_allowed', 'Skill files must not be symbolic links.')
      if (entry.isDirectory()) await walk(child)
      else if (entry.isFile()) {
        const info = await lstat(child)
        total += info.size
        if (total > 16 * 1024 * 1024 || files.size >= 100)
          throw new WorkbookError(
            'skill_install_too_large',
            'The skill folder exceeds the reviewable installation limit; existing files were preserved.',
          )
        const bytes = await readFile(child)
        if (bytes.length !== info.size)
          throw new WorkbookError('skill_changed', 'A skill file changed during inspection.')
        files.set(relative(directory, child).split(sep).join('/'), bytes)
      } else
        throw new WorkbookError(
          'skill_path_not_allowed',
          'Only regular skill resources may be installed.',
        )
    }
  }
  await walk(directory)
  return files
}
function fileDigests(files: Tree) {
  return Object.fromEntries(
    [...files]
      .sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0))
      .map(([path, bytes]) => [path, sha(bytes)]),
  )
}
function digest(files: Tree) {
  return sha(JSON.stringify(fileDigests(files)))
}
async function within(root: string, path: string) {
  let ancestor = path
  for (;;) {
    try {
      const resolved = await realpath(ancestor)
      if (resolved !== root && !isInside(root, resolved))
        throw new WorkbookError(
          'skill_path_not_allowed',
          'The skill path escapes the selected workspace.',
        )
      return
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error
      const parent = dirname(ancestor)
      if (parent === ancestor)
        throw new WorkbookError('skill_path_not_allowed', 'No workspace parent exists.')
      ancestor = parent
    }
  }
}
async function directory(root: string, path: string) {
  await within(root, path)
  await mkdir(path, { recursive: true, mode: 0o700 })
  await within(root, path)
}
async function current(root: string, path: string) {
  await within(root, path)
  const info = await lstat(path).catch((error: NodeJS.ErrnoException) => {
    if (error.code === 'ENOENT') return undefined
    throw error
  })
  if (!info) return undefined
  if (!info.isDirectory() || info.isSymbolicLink())
    throw new WorkbookError(
      'skill_path_not_allowed',
      'The installed skill must be a regular workspace directory.',
    )
  return tree(path)
}
async function receipt(path: string) {
  const file = join(path, RECEIPT)
  const info = await lstat(file).catch(() => undefined)
  if (!info?.isFile() || info.isSymbolicLink() || info.size > 256 * 1024) return undefined
  try {
    const value = receiptSchema.safeParse(JSON.parse(await readFile(file, 'utf8')))
    return value.success ? value.data : undefined
  } catch {
    return undefined
  }
}

export interface SkillPreview {
  workspaceRoot: string
  bundleDirectory: string
  version: string
  targetDirectory: string
  exists: boolean
  userModified: boolean
  currentDigest: string | null
  bundleDigest: string
  installedVersion: string | null
  updateAvailable: boolean
  differences: { path: string; kind: 'added' | 'changed' | 'same' | 'preserved' }[]
}
export async function inspectSkillInstallation(input: {
  workspaceRoot: string
  bundleDirectory: string
  version: string
}): Promise<SkillPreview> {
  const workspaceRoot = await realpath(input.workspaceRoot),
    bundleDirectory = await realpath(input.bundleDirectory),
    targetDirectory = join(workspaceRoot, '.github', 'skills', 'genoffice-workbook')
  const packaged = await tree(bundleDirectory)
  if (!packaged.has('SKILL.md'))
    throw new WorkbookError('skill_bundle_invalid', 'The packaged workbook skill is incomplete.')
  const existing = await current(workspaceRoot, targetDirectory),
    prior = existing ? await receipt(targetDirectory) : undefined
  const currentDigest = existing ? digest(existing) : null,
    bundleDigest = digest(packaged)
  return {
    workspaceRoot,
    bundleDirectory,
    version: input.version,
    targetDirectory,
    exists: !!existing,
    currentDigest,
    bundleDigest,
    installedVersion: prior?.version ?? null,
    userModified: !!existing && (!prior || prior.installedDigest !== currentDigest),
    updateAvailable:
      !prior || prior.version !== input.version || prior.bundleDigest !== bundleDigest,
    differences: [...new Set([...packaged.keys(), ...(existing?.keys() ?? [])])]
      .sort()
      .map((path) => ({
        path,
        kind: !packaged.has(path)
          ? 'preserved'
          : !existing?.has(path)
            ? 'added'
            : sha(packaged.get(path)!) === sha(existing.get(path)!)
              ? 'same'
              : 'changed',
      })),
  }
}

export async function installWorkbookSkill(
  input: SkillPreview & { decision: 'install' | 'update' | 'keep' },
) {
  if (input.decision === 'keep')
    return {
      status: 'kept' as const,
      targetDirectory: input.targetDirectory,
      backupDirectory: undefined,
    }
  const fresh = await inspectSkillInstallation(input)
  if (
    fresh.targetDirectory !== input.targetDirectory ||
    fresh.currentDigest !== input.currentDigest ||
    fresh.bundleDigest !== input.bundleDigest
  )
    throw new WorkbookError(
      'skill_changed',
      'The skill changed after the preview; inspect it again before updating.',
    )
  if (
    (input.decision === 'install' && fresh.exists) ||
    (input.decision === 'update' && !fresh.exists)
  )
    throw new WorkbookError('skill_changed', 'The installation target changed; review it again.')
  const packaged = await tree(fresh.bundleDirectory),
    existing = await current(fresh.workspaceRoot, fresh.targetDirectory)
  const merged = new Map(existing ?? [])
  for (const [path, bytes] of packaged) merged.set(path, bytes)
  const stagingParent = join(fresh.workspaceRoot, '.genoffice', 'skill-staging'),
    backupParent = join(fresh.workspaceRoot, '.genoffice', 'skill-backups')
  await directory(fresh.workspaceRoot, stagingParent)
  const staging = await mkdtemp(join(stagingParent, 'workbook-'))
  let backupDirectory: string | undefined,
    published = false
  try {
    for (const [path, bytes] of merged) {
      const destination = join(staging, ...path.split('/'))
      await mkdir(dirname(destination), { recursive: true, mode: 0o700 })
      await writeFile(destination, bytes, { flag: 'wx', mode: 0o600 })
    }
    await writeFile(
      join(staging, RECEIPT),
      JSON.stringify({
        schemaVersion: 1,
        version: fresh.version,
        bundleDigest: fresh.bundleDigest,
        installedDigest: digest(merged),
        fileDigests: fileDigests(merged),
        installedAt: Date.now(),
      }),
      { flag: 'wx', mode: 0o600 },
    )
    await directory(fresh.workspaceRoot, dirname(fresh.targetDirectory))
    const last = await current(fresh.workspaceRoot, fresh.targetDirectory)
    if ((last ? digest(last) : null) !== input.currentDigest)
      throw new WorkbookError(
        'skill_changed',
        'The installed skill changed during staging; no packaged files were published.',
      )
    if (last) {
      await directory(fresh.workspaceRoot, backupParent)
      backupDirectory = join(backupParent, `genoffice-workbook-${randomUUID()}`)
      await rename(fresh.targetDirectory, backupDirectory)
    }
    try {
      await rename(staging, fresh.targetDirectory)
      published = true
    } catch (error) {
      if (backupDirectory && !(await lstat(fresh.targetDirectory).catch(() => undefined))) {
        await rename(backupDirectory, fresh.targetDirectory)
        backupDirectory = undefined
      }
      throw error
    }
    const installed = await current(fresh.workspaceRoot, fresh.targetDirectory)
    if (!installed || digest(installed) !== digest(merged))
      throw new WorkbookError(
        'skill_changed',
        'The skill changed during publication; its previous copy is retained.',
        { published: true, backupDirectory: backupDirectory ?? null },
      )
    return { status: 'installed' as const, targetDirectory: fresh.targetDirectory, backupDirectory }
  } catch (error) {
    if (error instanceof WorkbookError) throw error
    throw new WorkbookError(
      'skill_install_failed',
      error instanceof Error ? error.message : String(error),
      { published, backupDirectory: backupDirectory ?? null },
    )
  } finally {
    await rm(staging, { recursive: true, force: true }).catch(() => {})
  }
}
