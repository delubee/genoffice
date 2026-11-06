import { expect, it } from 'vitest'
import JSZip from 'jszip'
import { blankXlsxBuffer } from '../src/gateway/csv-import'
import { applyCellEditsToXlsx } from '../src/gateway/xlsx-gateway'

it('replaces original styles for formatted fills and resets empty-donor cells to xf zero', async () => {
  const initial = await applyCellEditsToXlsx(await blankXlsxBuffer('Data'), [
    {
      sheetName: 'Data',
      row: 1,
      column: 0,
      writeValue: true,
      cell: { value: 99 },
      style: { bold: true, fillColor: '#CCEEFF', indent: 2, numberFormat: '0.00' },
    },
    {
      sheetName: 'Data',
      row: 1,
      column: 2,
      writeValue: true,
      cell: { value: 88 },
      style: { italic: true, fillColor: '#FF0000', indent: 5 },
    },
  ])
  const result = await applyCellEditsToXlsx(initial.buffer, [
    {
      sheetName: 'Data',
      row: 1,
      column: 0,
      writeValue: true,
      cell: { value: 12 },
      styleReset: true,
      style: { indent: 3, underline: true, numberFormat: '0.000' },
    },
    {
      sheetName: 'Data',
      row: 1,
      column: 2,
      writeValue: true,
      cell: { value: null },
      styleReset: true,
      style: {},
    },
  ])
  const zip = await JSZip.loadAsync(result.buffer)
  const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  const styles = await zip.file('xl/styles.xml')!.async('string')
  const xfs = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/
    .exec(styles)![1]!
    .match(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g)!
  const xf = (address: string) => {
    const cell = new RegExp('<c\\b[^>]*\\br="' + address + '"[^>]*>').exec(sheet)![0]
    return xfs[Number(/\bs="(\d+)"/.exec(cell)?.[1] ?? 0)]!
  }
  expect(xf('A2')).toContain('indent="3"')
  expect(xf('A2')).toContain('fillId="0"')
  expect(xf('C2')).toContain('fontId="0"')
  expect(xf('C2')).not.toContain('alignment')
})
