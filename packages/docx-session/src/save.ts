import { mkdtemp, writeFile, rm } from 'node:fs/promises'
import { join } from 'node:path'
import {
  captureFileBaseline,
  publishStagedWorkbook,
  type FileBaseline,
} from '@genoffice/workbook-session/publish'
import type { DocxSession } from './session'

export async function saveDocument(input: {
  session: DocxSession
  target: string
  baseline?: FileBaseline
  storage: string
  assertActive?(): void
}) {
  const baseline = input.baseline ?? (await captureFileBaseline(input.target))
  const directory = await mkdtemp(join(input.storage, 'docx-save-'))
  const stagedPath = join(directory, 'candidate.docx')
  try {
    await writeFile(stagedPath, await input.session.save(), { mode: 0o600, flag: 'wx' })
    input.assertActive?.()
    const next = await publishStagedWorkbook({ targetPath: input.target, stagedPath, baseline })
    await rm(directory, { recursive: true, force: true })
    return next
  } catch (error) {
    if (
      error &&
      typeof error === 'object' &&
      'code' in error &&
      error.code === 'external_change_conflict'
    )
      input.session.markConflict()
    // Keep the staged document for inspection after an external writer conflict.
    throw new Error(`DOCX save failed; recovery candidate: ${stagedPath}`, { cause: error })
  }
}
