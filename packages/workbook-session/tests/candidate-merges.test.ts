import { expect, it } from 'vitest'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { initialState } from '../src/state'

const before = initialState({
  sheets: [{ id: 'base', name: 'Data', rowCount: 10, columnCount: 5 }],
})
const sheetId = before.sheets[0]!.id
const material = structuredClone(before)
material.sheets[0]!.features = {
  merges: [{ startRow: 0, endRow: 1, startColumn: 0, endColumn: 2 }],
}
material.sheets[0]!.cells.A1 = { kind: 'string', rawValue: 'Heading' }
const run = (op: WorkbookOperation, origin: 'manual' | 'agent' = 'manual') =>
  computeCandidateState(
    before,
    {
      schemaVersion: 1,
      sessionId: 'merges',
      transactionId: 'edit',
      expectedRevision: 0,
      origin,
      summary: 'Merged regression',
      operations: [op],
    },
    {
      state: material,
      completeSheetIds: [sheetId],
      writableSheetIds: [sheetId],
      allocation: { stateId: 'after', sheetIds: {} },
    },
  )

it('allows anchor edits through manual and agent candidates without creating covered values', async () => {
  for (const origin of ['manual', 'agent'] as const) {
    const result = await run({ op: 'set_cell', sheetId, address: 'A1', value: 'Updated' }, origin)
    expect(result.nextState.sheets[0]!.cells.A1?.rawValue).toBe('Updated')
    expect(result.nextState.sheets[0]!.cells.B1).toBeUndefined()
  }
})
it('rejects covered-cell and multi-value writes atomically', async () => {
  for (const op of [
    { op: 'set_cell', sheetId, address: 'B1', value: 'Hidden' },
    { op: 'set_formula', sheetId, address: 'A2', formula: '=1' },
    { op: 'set_range', sheetId, range: 'A1:B1', values: [['One', 'Two']] },
    { op: 'clear_range', sheetId, range: 'B1:C2' },
  ] as WorkbookOperation[])
    await expect(run(op)).rejects.toThrow('merged_cell_target')
  expect(material.sheets[0]!.cells.A1?.rawValue).toBe('Heading')
})
it('permits clearing and formatting whole merges but rejects unsupported merge copies and structure', async () => {
  expect(
    (await run({ op: 'clear_range', sheetId, range: 'A1:C2' })).nextState.sheets[0]!.cells.A1
      ?.rawValue,
  ).toBeNull()
  await run({ op: 'format_range', sheetId, range: 'A1:C2', format: { bold: true } })
  await expect(run({ op: 'copy_range', sheetId, source: 'A1:C2', target: 'A5' })).rejects.toThrow(
    'merged_range_copy',
  )
  await expect(run({ op: 'insert_rows', sheetId, row: 1, count: 1 })).rejects.toThrow(
    'unsupported_structural_feature: merges',
  )
})

it('also blocks covered values arriving through a Python artifact import', async () => {
  const { prepareImportCandidate } = await import('../src/exchange/import-candidate')
  const source = {
    async openSnapshot() {
      throw new Error('unused')
    },
    async close() {},
    async archiveManifest() {
      return {}
    },
    async readFeatures() {
      return {
        indexingComplete: true,
        sheetProtection: null,
        styles: [],
        merges: material.sheets[0]!.features!.merges,
      }
    },
    async readBaseRange() {
      return { cells: [] }
    },
  }
  await expect(
    prepareImportCandidate({
      state: before,
      intent: {
        schemaVersion: 1,
        sessionId: 'merges',
        transactionId: 'python',
        expectedRevision: 0,
        origin: 'import',
        summary: 'Python merged regression',
        operations: [],
      },
      source,
      sourceSessionId: 'base',
      batch: {
        planId: 'python-result',
        resultHash: '0'.repeat(64),
        mode: 'values',
        target: { sheetId },
        cells: [{ address: 'B1', cell: { kind: 'string', value: 'Hidden' } }],
      },
    }),
  ).rejects.toThrow('merged_cell_target')
})
