import { expect, it } from 'vitest'
import { writeFile, readFile, readdir, chmod } from 'node:fs/promises'
import { join } from 'node:path'
import { createRecoveryPoint, readRecoveryPoint } from '../src/recovery'
import { initialState } from '../src/state'
import { hashFile } from '../src/resources'
import { useSessionTestResources } from './helpers'
const resources = useSessionTestResources()

async function context() {
  const directory = await resources.directory(),
    path = join(directory, 'original.xlsx')
  await writeFile(path, 'immutable original bytes')
  const state = structuredClone(
    initialState({ sheets: [{ id: 'base', name: 'Data', rowCount: 2, columnCount: 1 }] }),
  )
  state.sheets[0]!.cells.A1 = { kind: 'string', rawValue: '00123' }
  return {
    directory,
    path,
    state,
    input: {
      storageDirectory: directory,
      original: { path, sha256: await hashFile(path) },
      state,
      savedStateId: 'saved-before-edit',
      documentUri: 'file:///fixture.xlsx',
      baselineHash: await hashFile(path),
    },
  }
}

it('keeps the last confirmed backup when a newer backup is interrupted and verifies the complete feature root', async () => {
  const { directory, input, state } = await context()
  state.sheets[0]!.features = { charts: [{ id: 'chart', data: [1, 2] }] }
  const first = await createRecoveryPoint(input)
  const modified = structuredClone(state)
  modified.stateId = 'new-state'
  modified.sheets[0]!.cells.A1 = { kind: 'number', rawValue: 8 }
  await expect(
    createRecoveryPoint({
      ...input,
      state: modified,
      fault(point) {
        if (point === 'beforeManifestPublish') throw new Error('Injected interrupted backup')
      },
    }),
  ).rejects.toThrow('interrupted')
  const restored = await readRecoveryPoint({
    manifestPath: first.manifestPath,
    expectedManifestHash: first.manifestHash,
    documentUri: input.documentUri,
    allowedRoots: [directory],
  })
  expect(restored.state).toEqual(state)
  expect(restored.savedStateId).toBe('saved-before-edit')
  expect('sessionId' in restored).toBe(false)
  expect((await readdir(directory)).filter((name) => name.startsWith('backup-'))).toHaveLength(1)
  await first.dispose()
  await expect(readFile(first.manifestPath)).rejects.toMatchObject({ code: 'ENOENT' })
})

it('rejects manifest/base tampering and paths outside storage without opening a new session', async () => {
  const { directory, input } = await context(),
    point = await createRecoveryPoint(input)
  const read = () =>
    readRecoveryPoint({
      manifestPath: point.manifestPath,
      expectedManifestHash: point.manifestHash,
      documentUri: input.documentUri,
      allowedRoots: [directory],
    })
  const manifest = await readFile(point.manifestPath)
  await writeFile(point.manifestPath, Buffer.concat([manifest, Buffer.from(' ')]))
  await expect(read()).rejects.toMatchObject({ code: 'recovery_invalid' })
  await writeFile(point.manifestPath, manifest)
  await chmod(join(point.directory, 'base.xlsx'), 0o600)
  await writeFile(join(point.directory, 'base.xlsx'), 'tampered original')
  await expect(read()).rejects.toMatchObject({ code: 'recovery_invalid' })
  await expect(
    readRecoveryPoint({
      manifestPath: point.manifestPath,
      expectedManifestHash: point.manifestHash,
      documentUri: input.documentUri,
      allowedRoots: [await resources.directory()],
    }),
  ).rejects.toMatchObject({ code: 'recovery_invalid' })
})

it('checks budgets/cancellation before acknowledgment and acknowledges a fully published copy despite late cancellation', async () => {
  const { directory, input } = await context()
  await expect(createRecoveryPoint({ ...input, byteLimit: 10 })).rejects.toMatchObject({
    code: 'undo_budget_exceeded',
  })
  const cancelled = new AbortController()
  cancelled.abort()
  await expect(createRecoveryPoint({ ...input, signal: cancelled.signal })).rejects.toMatchObject({
    code: 'operation_cancelled',
  })
  const late = new AbortController(),
    point = await createRecoveryPoint({
      ...input,
      signal: late.signal,
      fault(phase) {
        if (phase === 'afterManifestPublish') late.abort()
      },
    })
  const restored = await readRecoveryPoint({
    manifestPath: point.manifestPath,
    expectedManifestHash: point.manifestHash,
    documentUri: input.documentUri,
    allowedRoots: [directory],
  })
  expect(restored.state).toEqual(input.state)
})
