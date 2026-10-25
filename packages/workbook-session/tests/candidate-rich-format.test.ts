import { expect, it } from 'vitest'
import { initialState } from '../src/state'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import { ribbonIntents } from '@genoffice/sheets-editor/ribbon-intents'
import { gridCell } from '@genoffice/sheets-editor/grid-cell'
import { toSelectionFormat } from '../../../apps/sheets/src/renderer/selection-format'
import type { IStyleData } from '@univerjs/core'
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'

it('preserves rich Home formatting, outline corners, explicit clears and inverse state across later edits', async () => {
  let state = structuredClone(
    initialState({ sheets: [{ id: 'base', name: 'Data', rowCount: 3, columnCount: 3 }] }),
  )
  const sheetId = state.sheets[0]!.id
  state.sheets[0]!.cells.A1 = {
    kind: 'number',
    rawValue: 12.5,
    style: {
      underline: true,
      underlineStyle: 'single',
      vertAlign: 'subscript',
      fill: {
        gradient: {
          stops: [
            { position: 0, color: '#111111' },
            { position: 1, color: '#EEEEEE' },
          ],
        },
      },
      fillColor: '#111111',
      fontColor: '#222222',
      borderBottom: { style: 'dotted', color: '#AA0000' },
    },
  }
  const apply = async (operations: WorkbookOperation[]) => {
    const before = structuredClone(state)
    const result = await computeCandidateState(
      state,
      {
        schemaVersion: 1,
        sessionId: 'rich-format',
        transactionId: crypto.randomUUID(),
        expectedRevision: 0,
        origin: 'manual',
        summary: 'Desktop Home format',
        operations,
      },
      {
        state: structuredClone(state),
        completeSheetIds: [sheetId],
        writableSheetIds: [sheetId],
        allocation: { stateId: crypto.randomUUID(), sheetIds: {} },
      },
    )
    expect(result.inverse.state).toEqual(before)
    state = result.nextState
  }
  const originalStyle = gridCell(state.sheets[0]!.cells.A1!).s as IStyleData
  const target = {
    sheetId,
    address: 'A1:C3',
    format: toSelectionFormat(originalStyle, 'General'),
    doubleUnderline: false,
  }
  await apply(
    [
      'underline:double',
      'superscript',
      'align:distributed',
      'rotate:-45',
      'indent:2',
      'cellprot:locked-off:hidden-on',
      'border:outer:#123456:double',
      'fill:none',
      'font-color:auto',
    ].flatMap((command) => ribbonIntents(command, target)),
  )
  const styled = state.sheets[0]!.cells.A1!.style!
  expect(styled).toMatchObject({
    underline: true,
    underlineStyle: 'double',
    vertAlign: 'superscript',
    horizontalAlignment: 'distributed',
    textRotation: 135,
    indent: 2,
    protectionLocked: false,
    protectionHidden: true,
    fillColor: null,
    fontColor: null,
    borderTop: { style: 'double', color: '#123456' },
    borderLeft: { style: 'double', color: '#123456' },
    borderBottom: { style: 'dotted', color: '#AA0000' },
  })
  expect(styled.fill).toBeUndefined()
  expect(state.sheets[0]!.cells.C3!.style).toMatchObject({
    borderBottom: { style: 'double' },
    borderRight: { style: 'double' },
  })
  expect(state.sheets[0]!.cells.B2!.style!.borderTop).toBeUndefined()
  expect(gridCell(state.sheets[0]!.cells.A1!).s).toMatchObject({
    ul: { s: 1, t: 10 },
    va: 3,
    ht: 6,
    tr: { a: -45 },
    pd: { l: 16 },
    cl: { rgb: '' },
    bg: { rgb: '' },
  })
  await apply([{ op: 'format_range', sheetId, range: 'A1:C3', format: { bold: true } }])
  expect(state.sheets[0]!.cells.A1!.style).toMatchObject({ ...styled, bold: true })
  await apply([
    {
      op: 'format_range',
      sheetId,
      range: 'A1:C3',
      format: {
        vertAlign: null,
        underline: false,
        fillColor: '#00AA00',
        textRotation: 0,
      },
    },
  ])
  expect(state.sheets[0]!.cells.A1!.style).toMatchObject({
    vertAlign: null,
    underline: false,
    fillColor: '#00AA00',
    textRotation: 0,
  })
  expect(state.sheets[0]!.cells.A1!.style!.fill).toBeUndefined()
})

it('keeps number-format colons and validates a complete dialog before any submission', () => {
  const target = {
    sheetId: 'sheet',
    address: 'A1:B2',
    format: toSelectionFormat({}, 'General'),
    doubleUnderline: false,
  }
  expect(ribbonIntents('format:h:mm:ss AM/PM', target)[0]).toMatchObject({
    format: { numberFormat: 'h:mm:ss AM/PM' },
  })
  const submit = (commands: string[]) =>
    commands.flatMap((command) => ribbonIntents(command, target))
  expect(() => submit(['bold:on', 'font-size:NaN'])).toThrow()
  expect(() => submit(['border:outer:#123456:injected'])).toThrow()
  expect(() => submit(['cellprot:locked-yes:'])).toThrow()
})
