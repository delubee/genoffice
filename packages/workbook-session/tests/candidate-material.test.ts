import { expect, it, vi } from 'vitest'
import { prepareCandidateMaterial } from '../src/candidate-material'
import { initialState } from '../src/state'
import type { WorkbookSource } from '../src/ports'

it('waits for complete protection metadata before reading edit material or granting write access', async () => {
  const state = initialState({
      sheets: [{ id: 'base', name: 'Data', rowCount: 1, columnCount: 1 }],
    }),
    sheetId = state.sheets[0]!.id
  let indexed = false
  const readFeatures = vi.fn(async () => {
    if (!indexed) {
      indexed = true
      return { indexingComplete: false }
    }
    return { indexingComplete: true, sheetProtection: { protected: true } }
  })
  const source: WorkbookSource = {
    async openSnapshot() {
      throw new Error('Unused')
    },
    async archiveManifest() {
      return {}
    },
    async close() {},
    readFeatures,
    async readBaseRange() {
      expect(indexed).toBe(true)
      expect(readFeatures.mock.calls.length).toBeGreaterThanOrEqual(2)
      return { cells: [{ row: 0, column: 0, value: 7 }] }
    },
  }
  const material = await prepareCandidateMaterial({
    state,
    source,
    sourceSessionId: 'source',
    operations: [{ op: 'set_cell', sheetId, address: 'A1', value: 8 }],
  })
  expect(material.writableSheetIds).toEqual([])
  expect(material.state.sheets[0]!.protected).toBe(true)
  expect(material.state.sheets[0]!.cells.A1!.rawValue).toBe(7)
})
