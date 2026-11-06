import { readFile, stat } from 'node:fs/promises'
import { createHash } from 'node:crypto'
import { captureFileBaseline, sameFileBaseline } from '@genoffice/workbook-session/publish'
import { DOCX_LIMITS } from '@genoffice/docs-editor/shared/policy'
import { DocxSession } from './session'

export async function openDocument(path: string) {
  const baseline = await captureFileBaseline(path)
  if (!baseline.sha256) throw new Error('Document does not exist')
  const info = await stat(baseline.canonicalPath)
  if (info.size > DOCX_LIMITS.file) throw new Error('DOCX exceeds 16 MiB limit')
  const bytes = await readFile(baseline.canonicalPath)
  if (
    createHash('sha256').update(bytes).digest('hex') !== baseline.sha256 ||
    !sameFileBaseline(baseline, await captureFileBaseline(path))
  )
    throw new Error('Document changed while opening')
  return { session: await DocxSession.open(bytes), baseline }
}
