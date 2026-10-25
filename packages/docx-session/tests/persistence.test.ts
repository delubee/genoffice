import { it, expect } from 'vitest'
import { mkdtemp, writeFile, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { buildDocx } from '../../docx-engine/tests/helpers/build-docx'
import { openDocument } from '../src/open'
import { saveDocument } from '../src/save'

it('refuses to overwrite an externally changed file and retains the edited recovery candidate', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'docx-persistence-'))
  try {
    const path = join(directory, 'document.docx')
    const original = await buildDocx({ bodyXml: '<w:p><w:r><w:t>Original</w:t></w:r></w:p>' })
    await writeFile(path, original)
    const { session, baseline } = await openDocument(path)
    const blockId = session.read().blocks[0]!.blockId
    const input = {
      transactionId: 'edit',
      expectedRevision: 0,
      operations: [{ type: 'replaceText', blockId, start: 0, end: 8, text: 'Edited' }],
    }
    const receipt = session.applyOperations(input)
    expect(session.applyOperations(input)).toEqual(receipt)
    const external = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>External writer</w:t></w:r></w:p>',
    })
    await writeFile(path, external)
    await expect(
      saveDocument({ session, target: path, baseline, storage: directory }),
    ).rejects.toThrow('recovery candidate')
    expect(session.descriptor()).toMatchObject({ dirty: true, conflict: true })
    expect(await readFile(path)).toEqual(Buffer.from(external))
  } finally {
    await rm(directory, { recursive: true, force: true })
  }
})

it('saves and reopens edits while keeping the immutable opened base for subsequent saves', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'docx-persistence-'))
  try {
    const path = join(directory, 'document.docx')
    const original = await buildDocx({ bodyXml: '<w:p><w:r><w:t>Original</w:t></w:r></w:p>' })
    await writeFile(path, original)
    const opened = await openDocument(path)
    expect(Buffer.from(await opened.session.save())).toEqual(Buffer.from(original))
    opened.session.applyOperations({
      transactionId: 'edit',
      expectedRevision: 0,
      operations: [
        {
          type: 'replaceText',
          blockId: opened.session.read().blocks[0]!.blockId,
          start: 0,
          end: 8,
          text: 'Edited',
        },
      ],
    })
    const baseline = await saveDocument({ ...opened, target: path, storage: directory })
    expect((await openDocument(path)).session.read().blocks[0]!.text).toBe('Edited')
    opened.session.history('undo')
    await saveDocument({ session: opened.session, baseline, target: path, storage: directory })
    expect(await readFile(path)).toEqual(Buffer.from(original))
  } finally {
    await rm(directory, { recursive: true, force: true })
  }
})
