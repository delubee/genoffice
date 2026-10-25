import { createHash, randomUUID } from 'node:crypto'
import { readFile, realpath, stat, writeFile } from 'node:fs/promises'
import { basename, dirname, isAbsolute, join, relative } from 'node:path'
import { z } from 'zod'
import type { Commit } from '@genoffice/docs-editor/shared/protocol'
import type { DocxOperation } from '@genoffice/docs-editor/shared/operations'
import { jsonBytes } from '@genoffice/docs-editor/shared/policy'
import type { BlockView } from './read'
import type { DocxSession } from './session'

const maxBytes = 2 * 1024 ** 2
const snapshotSchema = z
  .object({
    schemaVersion: z.literal(1),
    snapshotId: z.string().uuid(),
    sessionId: z.string().uuid(),
    stateId: z.string().uuid(),
    revision: z.number().int().nonnegative(),
    blocks: z
      .array(
        z
          .object({
            blockId: z.string(),
            type: z.string(),
            protected: z.boolean(),
            text: z.string().optional(),
          })
          .strict(),
      )
      .max(1000),
  })
  .strict()
type Snapshot = z.infer<typeof snapshotSchema>
interface Prepared {
  snapshot: Snapshot
  path: string
  sha256: string
  operations: DocxOperation[]
  receipt?: { transactionId: string; commit: Commit }
}
const hash = (bytes: Uint8Array) => createHash('sha256').update(bytes).digest('hex')

async function confined(path: string, roots: string[], writing = false) {
  if (!isAbsolute(path)) throw new Error('Exchange requires an absolute workspace path')
  const canonical = writing
    ? join(await realpath(dirname(path)), basename(path))
    : await realpath(path)
  const allowed = await Promise.all(roots.map((root) => realpath(root)))
  if (
    !allowed.some((root) => {
      const child = relative(root, canonical)
      return child !== '' && !child.startsWith('..') && !isAbsolute(child)
    })
  )
    throw new Error('Exchange path must be inside the workspace')
  return canonical
}
async function readResult(path: string) {
  const info = await stat(path)
  if (!info.isFile() || info.size > maxBytes) throw new Error('Exchange result exceeds 2 MiB')
  const bytes = await readFile(path)
  if (bytes.length > maxBytes) throw new Error('Exchange result exceeds 2 MiB')
  return { data: snapshotSchema.parse(JSON.parse(bytes.toString('utf8'))), sha256: hash(bytes) }
}

export class DocumentExchange {
  private snapshots = new Map<string, Snapshot>()
  private plans = new Map<string, Prepared>()
  async export(
    session: DocxSession,
    input: { path: string; roots: string[]; expectedRevision: number },
  ) {
    const descriptor = session.descriptor()
    if (input.expectedRevision !== descriptor.revision) throw new Error('Stale document version')
    const count = session.content().content?.length ?? 0
    if (count > 1000) throw new Error('Export exceeds 1000 blocks')
    const blocks: BlockView[] = []
    for (let offset = 0; offset < count; offset += 100)
      blocks.push(...session.read(offset, Math.min(100, count - offset)).blocks)
    const snapshot: Snapshot = {
      schemaVersion: 1,
      snapshotId: randomUUID(),
      sessionId: descriptor.sessionId,
      stateId: descriptor.stateId,
      revision: descriptor.revision,
      blocks,
    }
    if (jsonBytes(snapshot) > maxBytes) throw new Error('Export exceeds 2 MiB')
    const path = await confined(input.path, input.roots, true)
    const serialized = JSON.stringify(snapshot, null, 2)
    if (Buffer.byteLength(serialized) > maxBytes) throw new Error('Export exceeds 2 MiB')
    await writeFile(path, serialized, { mode: 0o600, flag: 'wx' })
    this.snapshots.set(snapshot.snapshotId, snapshot)
    if (this.snapshots.size > 32) this.snapshots.delete(this.snapshots.keys().next().value!)
    return {
      snapshotId: snapshot.snapshotId,
      path,
      descriptor,
      blocks: blocks.length,
      instructions:
        'Transform block text in this JSON using visible workspace Python. Preserve all IDs, metadata, order and protected blocks. Write the result to a new JSON file, then prepareImport. Do not overwrite the open DOCX.',
    }
  }
  async prepareImport(
    session: DocxSession,
    input: { snapshotId: string; path: string; roots: string[]; expectedRevision: number },
  ) {
    const snapshot = this.snapshots.get(input.snapshotId)
    const descriptor = session.descriptor()
    if (!snapshot || snapshot.sessionId !== descriptor.sessionId)
      throw new Error('Unknown export snapshot')
    if (
      input.expectedRevision !== descriptor.revision ||
      snapshot.stateId !== descriptor.stateId ||
      snapshot.revision !== descriptor.revision
    )
      throw new Error('Export snapshot is stale; export again')
    const path = await confined(input.path, input.roots)
    const result = await readResult(path)
    const data = result.data
    if (
      data.snapshotId !== snapshot.snapshotId ||
      data.sessionId !== snapshot.sessionId ||
      data.stateId !== snapshot.stateId ||
      data.revision !== snapshot.revision ||
      data.blocks.length !== snapshot.blocks.length
    )
      throw new Error('Export metadata was changed')
    const operations: DocxOperation[] = []
    for (let index = 0; index < snapshot.blocks.length; index++) {
      const before = snapshot.blocks[index]!,
        after = data.blocks[index]!
      if (
        before.blockId !== after.blockId ||
        before.type !== after.type ||
        before.protected !== after.protected
      )
        throw new Error('Export block identity or order was changed')
      if (before.text === after.text) continue
      if (
        before.protected ||
        !['docParagraph', 'docHeading', 'docListItem'].includes(before.type) ||
        typeof after.text !== 'string'
      )
        throw new Error('Unsupported or protected import target')
      operations.push({
        type: 'replaceText',
        blockId: before.blockId,
        start: 0,
        end: before.text?.length ?? 0,
        text: after.text,
      })
    }
    if (!operations.length || operations.length > 100)
      throw new Error('Import requires between 1 and 100 changed paragraphs')
    if (
      session.descriptor().stateId !== snapshot.stateId ||
      session.descriptor().revision !== snapshot.revision
    )
      throw new Error('Export snapshot is stale; export again')
    const importId = randomUUID()
    this.plans.set(importId, { snapshot, path, sha256: result.sha256, operations })
    if (this.plans.size > 32) this.plans.delete(this.plans.keys().next().value!)
    return {
      importId,
      sha256: result.sha256,
      expectedRevision: snapshot.revision,
      changedBlocks: operations.length,
    }
  }
  async applyImport(
    session: DocxSession,
    input: {
      importId: string
      sha256: string
      transactionId: string
      expectedRevision: number
      roots: string[]
      assertActive?(): void
    },
  ): Promise<Commit> {
    const plan = this.plans.get(input.importId)
    if (
      !plan ||
      plan.snapshot.sessionId !== session.sessionId ||
      input.sha256 !== plan.sha256 ||
      input.expectedRevision !== plan.snapshot.revision
    )
      throw new Error('Invalid import plan or result hash')
    if (plan.receipt) {
      if (input.transactionId !== plan.receipt.transactionId)
        throw new Error('Import already applied; use its original transactionId')
      return structuredClone(plan.receipt.commit)
    }
    const path = await confined(plan.path, input.roots)
    if (path !== plan.path || (await readResult(path)).sha256 !== plan.sha256)
      throw new Error('Import result changed after preview')
    const descriptor = session.descriptor()
    if (
      descriptor.stateId !== plan.snapshot.stateId ||
      descriptor.revision !== plan.snapshot.revision
    )
      throw new Error('Import plan is stale')
    input.assertActive?.()
    const commit = session.applyOperations({
      transactionId: input.transactionId,
      expectedRevision: input.expectedRevision,
      operations: plan.operations,
    })
    plan.receipt = { transactionId: input.transactionId, commit: structuredClone(commit) }
    return commit
  }
}
