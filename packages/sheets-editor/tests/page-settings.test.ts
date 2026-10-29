import { expect, it } from 'vitest'
import { pageSettingsCommand } from '../../../apps/sheets/src/renderer/page-settings-command'
const context = {
  sheetId: 'sheet',
  address: 'C5:E40',
  dimensions: { rows: 100, columns: 20 },
  prior: {},
  breaks: { rowBreaks: [2], colBreaks: [1] },
}
it('shares desktop paper, fit, area and capped title row semantics', () => {
  expect(pageSettingsCommand('orientation:landscape', context)).toEqual({
    orientation: 'landscape',
  })
  expect(pageSettingsCommand('fit-width:2', context)).toEqual({ fitToWidth: 2, fitToHeight: 0 })
  expect(pageSettingsCommand('fit-height:3', { ...context, prior: { fitToWidth: 2 } })).toEqual({
    fitToWidth: 2,
    fitToHeight: 3,
  })
  expect(pageSettingsCommand('print-area:set', { ...context, address: 'A1' })).toEqual({
    printArea: 'A1:A1',
  })
  expect(pageSettingsCommand('print-titles:selection', context)).toEqual({ printTitles: '5:25' })
  expect(pageSettingsCommand('paper:119', context)).toBeNull()
})
it('retains original breaks and distinguishes cell, full-row and full-column selections', () => {
  expect(pageSettingsCommand('breaks:insert', context)).toEqual({
    rowBreaks: [2, 4],
    colBreaks: [1, 2],
  })
  expect(pageSettingsCommand('breaks:insert', { ...context, address: 'A5:T5' })).toEqual({
    rowBreaks: [2, 4],
    colBreaks: [1],
  })
  expect(pageSettingsCommand('breaks:insert', { ...context, address: 'C1:C100' })).toEqual({
    rowBreaks: [2],
    colBreaks: [1, 2],
  })
  expect(pageSettingsCommand('breaks:remove', { ...context, address: 'B3' })).toEqual({
    rowBreaks: [],
    colBreaks: [],
  })
  expect(pageSettingsCommand('breaks:insert', { ...context, address: 'A1' })).toBeNull()
  expect(pageSettingsCommand('breaks:remove', context)).toBeNull()
})
