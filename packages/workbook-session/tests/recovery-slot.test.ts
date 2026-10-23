import { expect, it } from 'vitest'
import { writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { createRecoveryPoint } from '../src/recovery'
import { publishRecoverySlot, readRecoverySlot } from '../src/recovery-slot'
import { initialState } from '../src/state'
import { hashFile } from '../src/resources'
import { useSessionTestResources } from './helpers'
const resources = useSessionTestResources()

it('resolves an old native ID to the newest confirmed root and ignores deletion of an obsolete generation', async () => {
  const directory = await resources.directory(),
    path = join(directory, 'base.xlsx')
  await writeFile(path, 'immutable base')
  const original = { path, sha256: await hashFile(path) },
    state = initialState({ sheets: [{ id: 'base', name: 'Data', rowCount: 1, columnCount: 1 }] })
  const input = {
    storageDirectory: directory,
    original,
    documentUri: 'file:///book.xlsx',
    savedStateId: state.stateId,
    baselineHash: original.sha256,
  }
  const first = await publishRecoverySlot({
    ...input,
    point: await createRecoveryPoint({ ...input, state }),
  })
  const changed = structuredClone(state)
  changed.stateId = 'changed-root'
  changed.sheets[0]!.cells.A1 = { kind: 'number', rawValue: 101 }
  const second = await publishRecoverySlot({
    ...input,
    slotPath: first.path,
    point: await createRecoveryPoint({ ...input, state: changed }),
  })
  expect(second.path).toBe(first.path)
  await first.dispose()
  const read = () =>
    readRecoverySlot({
      slotPath: first.path,
      documentUri: input.documentUri,
      allowedRoots: [directory],
    })
  expect((await read()).state).toEqual(changed)
  await expect(
    readRecoverySlot({
      slotPath: first.path,
      documentUri: 'file:///other.xlsx',
      allowedRoots: [directory],
    }),
  ).rejects.toMatchObject({ code: 'recovery_invalid' })
  await second.dispose()
  await expect(read()).rejects.toMatchObject({ code: 'recovery_invalid' })
})
