import { expect, it } from 'vitest'
import { readFile, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { FileResourceStore } from '../src/store'
import { useSessionTestResources, failAt } from './helpers'
const resources = useSessionTestResources()
it('publishes only complete manifests and detects altered chunks', async () => {
  const directory = await resources.directory(),
    store = await FileResourceStore.create(directory)
  const reference = await store.writeChunk(new TextEncoder().encode('hello'))
  const data = { schemaVersion: 1, references: [reference], payload: { stateId: 'root' } }
  await store.publishManifest('root', data)
  expect(await store.readManifest('root')).toEqual(data)
  await expect(
    store.publishManifest('root', { ...data, references: ['a'.repeat(64)] }),
  ).rejects.toThrow()
  expect(await store.readManifest('root')).toEqual(data)
  await writeFile(join(directory, 'chunks', reference), 'tampered')
  await expect(store.readChunk(reference)).rejects.toThrow('digest')
  await expect(store.readChunk('../root')).rejects.toThrow()
})
it('an interrupted root publish leaves the previous root readable', async () => {
  const directory = await resources.directory(),
    store = await FileResourceStore.create(directory)
  await store.publishManifest('root', { schemaVersion: 1, references: [], payload: 'previous' })
  const faulty = await FileResourceStore.create(directory, failAt('beforeManifestRename'))
  await expect(
    faulty.publishManifest('root', { schemaVersion: 1, references: [], payload: 'candidate' }),
  ).rejects.toThrow('Injected')
  expect((await store.readManifest('root')).payload).toBe('previous')
  expect(await readFile(join(directory, 'manifests', 'root.json'), 'utf8')).not.toContain(
    'candidate',
  )
})
