import { createHash, randomUUID } from 'node:crypto'
import { mkdir, readFile, realpath, stat, writeFile } from 'node:fs/promises'
import { join, dirname, relative } from 'node:path'
import { z } from 'zod'
import type { FileBaseline } from '@genoffice/workbook-session/publish'
import { DOCX_LIMITS } from '@genoffice/docs-editor/shared/policy'
import { DocxSession } from './session'

const baselineSchema = z
  .object({
    requestedPath: z.string(),
    canonicalPath: z.string(),
    sha256: z.string().nullable(),
    physicalId: z.string().nullable(),
    size: z.string().nullable(),
    mtimeNs: z.string().nullable(),
    ctimeNs: z.string().nullable(),
    mode: z.number().int(),
  })
  .strict()
const backupSchema = z
  .object({
    schemaVersion: z.literal(1),
    sourceUri: z.string(),
    baseline: baselineSchema,
    sha256: z.string().regex(/^[a-f0-9]{64}$/),
    document: z.string().max(Math.ceil(DOCX_LIMITS.file / 3) * 4),
  })
  .strict()
const hash = (bytes: Uint8Array) => createHash('sha256').update(bytes).digest('hex')

export async function writeBackup(input: {
  session: DocxSession
  baseline: FileBaseline
  sourceUri: string
  storage: string
}) {
  const directory = join(input.storage, 'docx-backups')
  await mkdir(directory, { recursive: true, mode: 0o700 })
  const bytes = await input.session.save()
  if (bytes.length > DOCX_LIMITS.file) throw new Error('DOCX backup exceeds size limit')
  const payload = backupSchema.parse({
    schemaVersion: 1,
    sourceUri: input.sourceUri,
    baseline: input.baseline,
    sha256: hash(bytes),
    document: Buffer.from(bytes).toString('base64'),
  })
  const path = join(directory, `${randomUUID()}.json`)
  await writeFile(path, JSON.stringify(payload), { mode: 0o600, flag: 'wx' })
  return path
}

export async function restoreBackup(input: { path: string; storage: string; sourceUri: string }) {
  const expected = await realpath(join(input.storage, 'docx-backups'))
  const path = await realpath(input.path)
  if (relative(expected, dirname(path)) !== '') throw new Error('Invalid DOCX recovery path')
  const info = await stat(path)
  if (!info.isFile() || info.size > 24 * 1024 ** 2)
    throw new Error('DOCX backup exceeds size limit')
  const payload = backupSchema.parse(JSON.parse(await readFile(path, 'utf8')))
  if (payload.sourceUri !== input.sourceUri)
    throw new Error('DOCX backup belongs to another document')
  const bytes = Buffer.from(payload.document, 'base64')
  if (bytes.length > DOCX_LIMITS.file || hash(bytes) !== payload.sha256)
    throw new Error('DOCX backup integrity check failed')
  const session = await DocxSession.open(bytes)
  session.markUnsaved()
  return { session, baseline: payload.baseline }
}
