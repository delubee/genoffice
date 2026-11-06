import { fromUniverColorFilter, toUniverColorFilter } from '../src/domain/univer-color-filter'
import { expect, it } from 'vitest'
import { colorFilterCriteriaSchema, colorFilterMatches } from '../src/domain/color-filter'
import { applyFilterState } from '../src/gateway/xlsx-filter'

it('normalizes native multicolor drafts and preserves no-fill separately from white', () => {
  expect(fromUniverColorFilter({ cellFillColors: ['rgb(255, 0, 0)', null, '#ffffff'] })).toEqual({
    fill: ['#FF0000', null, '#FFFFFF'],
  })
  expect(toUniverColorFilter({ font: ['#FF0000', '#000000'] })).toEqual({
    cellTextColors: ['rgb(255,0,0)', 'rgb(0,0,0)'],
  })
  expect(colorFilterMatches({ fill: [null] }, undefined)).toBe(true)
  expect(colorFilterMatches({ fill: [null] }, { fillColor: '#FFFFFF' })).toBe(false)
  expect(colorFilterMatches({ fill: ['#FF0000', null] }, { fillColor: '#ff0000' })).toBe(true)
  expect(colorFilterMatches({ font: ['#000000'] }, undefined)).toBe(true)
  expect(
    colorFilterMatches({ font: ['#FF0000'] }, { fontColor: { theme: 4 } }, [
      '#FFFFFF',
      '#000000',
      '#000000',
      '#FFFFFF',
      '#FF0000',
    ]),
  ).toBe(true)
  expect(() => colorFilterCriteriaSchema.parse({ fill: [], font: ['#000000'] })).toThrow()
})
it('writes standard single-color dxf criteria and complete multicolor extension without changing cells', () => {
  const dxfs: string[] = []
  const state = {
    sheetName: 'Data',
    filter: {
      range: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 },
      columns: [{ colId: 0, colors: { fill: ['#FF0000'] } }],
    },
    visibilityRange: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 },
    hiddenRows: [2],
  }
  const xml =
    '<worksheet><sheetData><row r="1"><c r="A1"/></row><row r="2"><c r="A2" s="7"/></row><row r="3"><c r="A3" s="8"/></row></sheetData></worksheet>'
  const saved = applyFilterState(xml, state, {
    internDxf(xml) {
      dxfs.push(xml)
      return 9
    },
  })
  expect(saved).toContain('<colorFilter dxfId="9" cellColor="1"/>')
  expect(dxfs[0]).toContain('fgColor rgb="FFFF0000"')
  expect(saved).toContain('<go:color val="#FF0000"/>')
  expect(saved).toContain('<c r="A2" s="7"/>')
  expect(saved).toContain('<row r="3" hidden="1">')
  state.filter.columns[0]!.colors.fill = ['#FF0000', '#0000FF']
  const multi = applyFilterState(xml, state)
  expect(multi).not.toContain('<colorFilter ')
  expect(multi).toContain('<go:color val="#FF0000"/><go:color val="#0000FF"/>')
})
