import { it, expect } from 'vitest'
import { mkdtemp, writeFile, readFile, rm } from 'node:fs/promises'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { pathToFileURL } from 'node:url'
import { buildDocx } from '../../docx-engine/tests/helpers/build-docx'
import { openDocument } from '../src/open'
import { saveDocument } from '../src/save'
import { writeBackup, restoreBackup } from '../src/backup'

it('restores unsaved text into a fresh session and preserves the pre-crash disk baseline', async () => {
  const storage = await mkdtemp(join(tmpdir(), 'docx-backup-'))
  try {
    const path = join(storage, 'original.docx'),
      sourceUri = pathToFileURL(path).toString()
    await writeFile(path, await buildDocx({ bodyXml: '<w:p><w:r><w:t>Original</w:t></w:r></w:p>' }))
    const opened = await openDocument(path)
    opened.session.applyOperations({
      transactionId: 'edit',
      expectedRevision: 0,
      operations: [
        {
          type: 'replaceText',
          blockId: opened.session.read().blocks[0]!.blockId,
          start: 0,
          end: 8,
          text: 'Unsaved',
        },
      ],
    })
    const backup = await writeBackup({ ...opened, sourceUri, storage })
    await writeFile(
      path,
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>External change during crash</w:t></w:r></w:p>' }),
    )
    const restored = await restoreBackup({ path: backup, storage, sourceUri })
    expect(restored.session.sessionId).not.toBe(opened.session.sessionId)
    expect(restored.session.descriptor().dirty).toBe(true)
    expect(restored.session.read().blocks[0]!.text).toBe('Unsaved')
    await expect(saveDocument({ ...restored, target: path, storage })).rejects.toThrow(
      'recovery candidate',
    )
    await expect(
      restoreBackup({ path: backup, storage, sourceUri: 'file:///another.docx' }),
    ).rejects.toThrow('another document')
    const payload = JSON.parse(await readFile(backup, 'utf8'))
    payload.sha256 = '0'.repeat(64)
    await writeFile(backup, JSON.stringify(payload))
    await expect(restoreBackup({ path: backup, storage, sourceUri })).rejects.toThrow('integrity')
  } finally {
    await rm(storage, { recursive: true, force: true })
  }
})
