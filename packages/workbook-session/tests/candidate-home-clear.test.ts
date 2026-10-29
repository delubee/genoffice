import { expect, it } from 'vitest'
import { initialState } from '../src/state'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import { ribbonIntents, type RibbonTarget } from '@genoffice/sheets-editor/ribbon-intents'
import { toSelectionFormat } from '../../../apps/sheets/src/renderer/selection-format'

async function clear(command: string, filteredRows: number[] = []) {
  const state = structuredClone(
    initialState({ sheets: [{ id: 's', name: 'Data', rowCount: 4, columnCount: 2 }] }),
  )
  const sheet = state.sheets[0]!
  sheet.cells = {
    A1: {
      kind: 'number',
      rawValue: 12,
      formula: '=B1*2',
      style: { bold: true, fillColor: '#FFFF00', indent: 3, numberFormat: '0.00' },
      numberFormat: '0.00',
    },
    A2: { kind: 'number', rawValue: 99, style: { italic: true } },
    B1: { kind: 'string', rawValue: 'untouched', style: { underline: true } },
  }
  sheet.features = {
    originalMetadata: {
      comments: [
        { row: 0, column: 0, text: 'original note' },
        { row: 0, column: 1, text: 'keep note' },
      ],
    },
    noteEdits: { A2: 'new note' },
    hyperlinks: [{ row: 0, column: 0, target: 'https://example.com' }],
    hyperlinkEdits: { A2: '#Data!B1' },
  }
  const target: RibbonTarget = {
    sheetId: sheet.id,
    address: 'A1:A2',
    format: toSelectionFormat({}, 'General'),
    doubleUnderline: false,
    noteAddresses: ['A1', 'A2'],
    hyperlinkAddresses: ['A1', 'A2'],
    filteredRows,
  }
  const result = await computeCandidateState(
    state,
    {
      schemaVersion: 1,
      sessionId: 'clear',
      transactionId: crypto.randomUUID(),
      expectedRevision: 0,
      origin: 'manual',
      summary: command,
      operations: ribbonIntents(command, target),
    },
    {
      state: structuredClone(state),
      completeSheetIds: [sheet.id],
      writableSheetIds: [sheet.id],
      allocation: { stateId: crypto.randomUUID(), sheetIds: {} },
    },
  )
  expect(result.inverse.state).toEqual(state)
  expect(result.nextState.sheets[0]!.cells.B1).toEqual(sheet.cells.B1)
  return { before: sheet, after: result.nextState.sheets[0]! }
}

it('clears only formatting, retains formulas/notes/links, and skips filtered rows', async () => {
  const { before, after } = await clear('clear-formats', [1])
  expect(after.cells.A1).toMatchObject({
    formula: '=B1*2',
    styleReset: true,
    numberFormat: 'General',
    style: { numberFormat: 'General' },
  })
  expect(after.cells.A1!.style?.bold).toBeUndefined()
  expect(after.cells.A1!.style?.fillColor).toBeUndefined()
  expect(after.cells.A2).toEqual(before.cells.A2)
  expect(after.features?.noteEdits).toEqual(before.features?.noteEdits)
  expect(after.features?.hyperlinkEdits).toEqual(before.features?.hyperlinkEdits)
})

it('clears notes or hyperlinks independently while retaining content and formatting', async () => {
  const notes = await clear('clear-comments')
  expect(notes.after.cells.A1!.formula).toBe('=B1*2')
  expect(notes.after.cells.A1!.style?.bold).toBe(true)
  expect(notes.after.features?.noteEdits).toEqual({ A1: null, A2: null })
  expect(notes.after.features?.hyperlinkEdits).toEqual(notes.before.features?.hyperlinkEdits)
  const links = await clear('clear-hyperlinks')
  expect(links.after.features?.hyperlinkEdits).toEqual({ A1: null, A2: null })
  expect(links.after.features?.noteEdits).toEqual(links.before.features?.noteEdits)
})

it('combines all clears into one reversible candidate without changing unselected cells', async () => {
  const { after } = await clear('clear-all')
  expect(after.cells.A1).toMatchObject({
    rawValue: null,
    styleReset: true,
    style: { numberFormat: 'General' },
  })
  expect(after.cells.A1!.formula).toBeUndefined()
  expect(after.features?.noteEdits).toEqual({ A1: null, A2: null })
  expect(after.features?.hyperlinkEdits).toEqual({ A1: null, A2: null })
})
