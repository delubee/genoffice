import { expect, it } from 'vitest'
import { initialState } from '../src/state'
import { prepareCandidateMaterial } from '../src/candidate-material'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import type { WorkbookSource } from '../src/ports'
const fixture = () => {
  const state = structuredClone(
    initialState({ sheets: [{ id: 'base', name: 'Data', rowCount: 2, columnCount: 2 }] }),
  )
  let indexed = false
  const source: WorkbookSource = {
    async openSnapshot() {
      throw new Error('unused')
    },
    async close() {},
    async archiveManifest() {
      return {}
    },
    async readBaseRange() {
      throw new Error('view-only changes must not hydrate cells')
    },
    async readFeatures() {
      if (!indexed) {
        indexed = true
        return { indexingComplete: false }
      }
      return { indexingComplete: true, sheetProtection: { protected: true } }
    },
  }
  const run = async (operations: WorkbookOperation[]) => {
    const material = await prepareCandidateMaterial({
      state,
      operations,
      source,
      sourceSessionId: 'source',
    })
    expect(material.writableSheetIds).toEqual([])
    expect(material.viewWritableSheetIds).toEqual(
      state.sheets[0]!.readOnly ? [] : [state.sheets[0]!.id],
    )
    return computeCandidateState(
      state,
      {
        schemaVersion: 1,
        sessionId: 'view',
        transactionId: crypto.randomUUID(),
        expectedRevision: 0,
        origin: 'manual',
        summary: 'View',
        operations,
      },
      material,
    )
  }
  return { state, source, run, sheetId: state.sheets[0]!.id }
}
it('grants separately verified view-only access after protection indexing without changing cells or protection', async () => {
  const f = fixture(),
    before = structuredClone(f.state)
  const result = await f.run([
    {
      op: 'set_page_setup',
      sheetId: f.sheetId,
      showFormulas: true,
      showGridlines: false,
      showHeadings: false,
      zoomScale: 120,
    },
  ])
  expect(result.nextState.sheets[0]!.features?.pageSetupEdits).toMatchObject({
    showFormulas: true,
    showGridlines: false,
    showHeadings: false,
    zoomScale: 120,
  })
  expect(result.nextState.sheets[0]!.protected).toBe(true)
  expect(result.nextState.sheets[0]!.features?.sheetProtection).toEqual({ protected: true })
  expect(result.nextState.sheets[0]!.cells).toEqual(before.sheets[0]!.cells)
  expect(result.inverse.state).toEqual(before)
  expect(f.state).toEqual(before)
})
it('does not grant print/layout or mixed-patch access to a protected sheet', async () => {
  const f = fixture(),
    before = structuredClone(f.state)
  for (const patch of [
    { orientation: 'landscape' as const },
    { showFormulas: true, orientation: 'landscape' as const },
    { showGridlines: true, printArea: 'A1:B2' },
  ])
    await expect(f.run([{ op: 'set_page_setup', sheetId: f.sheetId, ...patch }])).rejects.toThrow(
      'unverified_writable_target',
    )
  expect(f.state).toEqual(before)
})
it('keeps readonly targets blocked even for otherwise verified view controls', async () => {
  const f = fixture()
  f.state.sheets[0]!.readOnly = true
  await expect(
    f.run([{ op: 'set_page_setup', sheetId: f.sheetId, showFormulas: true }]),
  ).rejects.toThrow('readonly_workbook')
})
it('requires an explicit verified view grant and preserves atomic rejection of content in mixed batches', async () => {
  const f = fixture(),
    before = structuredClone(f.state)
  // Metadata-only protection changes cannot make an edit plan writable.
  const material = await prepareCandidateMaterial({
    state: f.state,
    operations: [],
    source: f.source,
    sourceSessionId: 'source',
  })
  const operation = { op: 'set_page_setup' as const, sheetId: f.sheetId, showFormulas: true }
  const intent = {
    schemaVersion: 1 as const,
    sessionId: 'view',
    transactionId: crypto.randomUUID(),
    expectedRevision: 0,
    origin: 'manual' as const,
    summary: 'View',
    operations: [operation],
  }
  await expect(
    computeCandidateState(f.state, intent, { ...material, viewWritableSheetIds: [] }),
  ).rejects.toThrow('unverified_writable_target')
  const mixed = {
    ...intent,
    operations: [
      operation,
      { op: 'set_cell' as const, sheetId: f.sheetId, address: 'A1', value: 99 },
    ],
  }
  await expect(computeCandidateState(f.state, mixed, material)).rejects.toThrow(
    'unverified_writable_target',
  )
  expect(f.state).toEqual(before)
})
