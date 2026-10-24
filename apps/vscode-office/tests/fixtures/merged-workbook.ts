import JSZip from 'jszip'
import { manualWorkbook } from './manual-workbook'

export const mergedRanges = [
  { startRow: 0, endRow: 0, startColumn: 0, endColumn: 2 },
  { startRow: 2, endRow: 4, startColumn: 3, endColumn: 3 },
  { startRow: 59, endRow: 61, startColumn: 49, endColumn: 51 },
]
export async function mergedWorkbook() {
  const zip = await JSZip.loadAsync(await manualWorkbook())
  let xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  xml = xml
    .replace(/<dimension[^>]*\/>/, '<dimension ref="A1:AZ62"/>')
    .replace(
      '</sheetData>',
      '<row r="60"><c r="AX60" t="inlineStr"><is><t>Offscreen heading</t></is></c></row></sheetData>',
    )
    .replace(
      '</worksheet>',
      '<mergeCells count="3"><mergeCell ref="A1:C1"/><mergeCell ref="D3:D5"/><mergeCell ref="AX60:AZ62"/></mergeCells></worksheet>',
    )
  zip.file('xl/worksheets/sheet1.xml', xml)
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}
