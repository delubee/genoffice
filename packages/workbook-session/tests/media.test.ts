import { expect, it, vi } from 'vitest'
import { initialState } from '../src/state'
import { readCurrentMedia } from '../src/media'
import type { WorkbookSource } from '../src/ports'

function fixture() {
  const state = structuredClone(
    initialState({ sheets: [{ id: 'base', name: 'Data', rowCount: 1, columnCount: 1 }] }),
  )
  state.features = {
    originalMetadata: {
      visuals: [
        { id: 'picture', sheetId: 'base', kind: 'image' },
        { id: 'shape', sheetId: 'base', kind: 'shape' },
        { id: 'chart', sheetId: 'base', kind: 'chart' },
        { id: 'foreign', sheetId: 'missing', kind: 'image' },
      ],
    },
  }
  const readMedia = vi.fn(async () => ({ mediaType: 'image/png', base64: 'AA==' }))
  const source = { readMedia } as unknown as WorkbookSource
  return {
    state,
    source,
    readMedia,
    read: (visualId = 'picture', stateId = state.stateId) =>
      readCurrentMedia({
        request: { stateId, visualId },
        getState: () => state,
        source,
        sourceSessionId: 'private-source',
      }),
  }
}
it('reads known current source pictures/shape fills through opaque identities', async () => {
  const f = fixture()
  expect(await f.read()).toEqual({ mediaType: 'image/png', base64: 'AA==' })
  await f.read('shape')
  expect(f.readMedia.mock.calls.map((call) => call)).toHaveLength(2)
  expect(f.readMedia).toHaveBeenCalledWith('private-source', 'picture')
})
it('rejects stale, foreign, removed, unsupported and ambiguous identities before source IO', async () => {
  const f = fixture()
  await expect(f.read('picture', 'outdated')).rejects.toThrow('no longer current')
  for (const id of ['missing', 'foreign', 'chart', '../../secret'])
    await expect(f.read(id)).rejects.toThrow('not in the current workbook')
  f.state.sheets[0]!.features = { visualEdits: { picture: { remove: true } } }
  await expect(f.read()).rejects.toThrow('not in the current workbook')
  f.state.sheets[0]!.features = {}
  ;(f.state.features!.originalMetadata as any).visuals.push({
    id: 'picture',
    sheetId: 'base',
    kind: 'image',
  })
  await expect(f.read()).rejects.toThrow('not in the current workbook')
  expect(f.readMedia).not.toHaveBeenCalled()
})
it('rejects a media read overtaken by another root and strips no unexpected source fields', async () => {
  const f = fixture()
  let finish!: (value: any) => void
  f.readMedia.mockImplementationOnce(
    () =>
      new Promise((resolve) => {
        finish = resolve
      }),
  )
  const pending = f.read()
  f.state.stateId = 'advanced'
  finish({ mediaType: 'image/png', base64: 'AA==' })
  await expect(pending).rejects.toThrow('changed while reading')
  f.readMedia.mockResolvedValueOnce({
    mediaType: 'image/png',
    base64: 'AA==',
    path: '/private/file',
  } as any)
  await expect(f.read()).rejects.toThrow()
})
