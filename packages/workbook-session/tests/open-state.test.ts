import { expect, it } from 'vitest'
import { readFile, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { openWorkbookSnapshot } from '../src/open'
import { initialState } from '../src/state'
import { useSessionTestResources } from './helpers'

const resources = useSessionTestResources()
it('opens copied bytes and retains the base until every reference is released', async () => {
  const directory = await resources.directory(),
    input = join(directory, 'input.xlsx')
  await writeFile(input, 'original')
  let openedPath = '',
    closes = 0
  const opened = await openWorkbookSnapshot({
    uri: new URL(`file://${input}`).href,
    storageDirectory: directory,
    source: {
      async openSnapshot(path) {
        openedPath = path
        await writeFile(input, 'changed')
        return { sessionId: 'sidecar-session', metadata: {} }
      },
      async archiveManifest(path) {
        expect(await readFile(path, 'utf8')).toBe('original')
        return { entries: [] }
      },
      async readBaseRange() {
        throw new Error('Not used')
      },
      async readFeatures() {
        return {}
      },
      async close() {
        closes++
      },
    },
  })
  expect(opened.original.sha256).toMatch(/^[a-f0-9]{64}$/)
  expect(openedPath).not.toBe(input)
  const release = opened.retain()
  await opened.dispose()
  expect(await readFile(openedPath, 'utf8')).toBe('original')
  await release()
  await expect(readFile(openedPath)).rejects.toThrow()
  expect(closes).toBe(1)
})

it('creates stable logical IDs and a base-relative empty overlay without loading cells', () => {
  const metadata = { sheets: [{ id: 'sheet-1', name: 'Data', rowCount: 10, columnCount: 4 }] }
  const state = initialState(metadata)
  expect(state.sheets[0]!.id).not.toBe('sheet-1')
  expect(state.sheets[0]!.originalSheetId).toBe('sheet-1')
  expect(state.sheets[0]!.cells).toEqual({})
  expect(state.structuralLog).toEqual([])
  expect(Object.isFrozen(state.sheets[0])).toBe(true)
})
