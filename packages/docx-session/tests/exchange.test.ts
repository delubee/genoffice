import { it, expect } from 'vitest'
import { mkdtemp, readFile, writeFile, rm, symlink } from 'node:fs/promises'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { DocxSession } from '../src/session'
import { DocumentExchange } from '../src/exchange'
import { buildDocx } from '../../docx-engine/tests/helpers/build-docx'

it('previews without editing, binds result hash/version and imports into shared history once', async () => {
  const root = await mkdtemp(join(tmpdir(), 'docx-exchange-'))
  try {
    const session = await DocxSession.open(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>Original</w:t></w:r></w:p>' }),
    )
    const exchange = new DocumentExchange()
    const snapshot = await exchange.export(session, {
      path: join(root, 'input.json'),
      roots: [root],
      expectedRevision: 0,
    })
    const data = JSON.parse(await readFile(snapshot.path, 'utf8'))
    data.blocks[0].text = 'Python result'
    const path = join(root, 'result.json')
    await writeFile(path, JSON.stringify(data))
    const plan = await exchange.prepareImport(session, {
      snapshotId: snapshot.snapshotId,
      path,
      roots: [root],
      expectedRevision: 0,
    })
    expect(session.read().blocks[0]!.text).toBe('Original')
    const input = { ...plan, roots: [root], transactionId: 'python-result' }
    await expect(
      exchange.applyImport(session, {
        ...input,
        assertActive: () => {
          throw new Error('Canceled before commit')
        },
      }),
    ).rejects.toThrow('Canceled')
    expect(session.descriptor().revision).toBe(0)
    const committed = await exchange.applyImport(session, input)
    expect(await exchange.applyImport(session, input)).toEqual(committed)
    expect(session.descriptor().revision).toBe(1)
    expect(session.read().blocks[0]!.text).toBe('Python result')
    session.history('undo')
    expect(session.read().blocks[0]!.text).toBe('Original')
    await expect(
      exchange.prepareImport(session, {
        snapshotId: snapshot.snapshotId,
        path,
        roots: [root],
        expectedRevision: 2,
      }),
    ).rejects.toThrow('stale')
  } finally {
    await rm(root, { recursive: true, force: true })
  }
})

it('rejects workspace escapes, tampered metadata and changes after preview without partial edits', async () => {
  const root = await mkdtemp(join(tmpdir(), 'docx-exchange-')),
    outside = await mkdtemp(join(tmpdir(), 'docx-outside-'))
  try {
    const session = await DocxSession.open(
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>Original</w:t></w:r></w:p>' }),
    )
    const exchange = new DocumentExchange()
    await expect(
      exchange.export(session, {
        path: join(outside, 'escape.json'),
        roots: [root],
        expectedRevision: 0,
      }),
    ).rejects.toThrow('workspace')
    await symlink(
      outside,
      join(root, 'junction'),
      process.platform === 'win32' ? 'junction' : 'dir',
    )
    await expect(
      exchange.export(session, {
        path: join(root, 'junction', 'escape.json'),
        roots: [root],
        expectedRevision: 0,
      }),
    ).rejects.toThrow('workspace')
    const snapshot = await exchange.export(session, {
      path: join(root, 'input.json'),
      roots: [root],
      expectedRevision: 0,
    })
    const data = JSON.parse(await readFile(snapshot.path, 'utf8')),
      path = join(root, 'result.json')
    data.blocks[0].text = 'Changed'
    data.revision = 99
    await writeFile(path, JSON.stringify(data))
    await expect(
      exchange.prepareImport(session, {
        snapshotId: snapshot.snapshotId,
        path,
        roots: [root],
        expectedRevision: 0,
      }),
    ).rejects.toThrow('metadata')
    data.revision = 0
    await writeFile(path, JSON.stringify(data))
    const plan = await exchange.prepareImport(session, {
      snapshotId: snapshot.snapshotId,
      path,
      roots: [root],
      expectedRevision: 0,
    })
    data.blocks[0].text = 'Changed after preview'
    await writeFile(path, JSON.stringify(data))
    await expect(
      exchange.applyImport(session, { ...plan, roots: [root], transactionId: 'changed-file' }),
    ).rejects.toThrow('changed after preview')
    expect(session.descriptor().revision).toBe(0)
    expect(session.read().blocks[0]!.text).toBe('Original')
  } finally {
    await Promise.all([
      rm(root, { recursive: true, force: true }),
      rm(outside, { recursive: true, force: true }),
    ])
  }
})
