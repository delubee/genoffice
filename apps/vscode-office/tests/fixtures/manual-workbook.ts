import JSZip from 'jszip'
import { blankXlsxBuffer } from '@genoffice/xlsx-gateway/gateway/csv-import'
import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
export const MANUAL_WORKBOOK_SHA256 =
  '8cdb24baa6557abf1638ae667ead65639d1fd6b08da69ee4d635e7496a783844'

/** Small independent manual-editor fixture, separate from Python fixtures. */
export async function manualWorkbook() {
  const base = await blankXlsxBuffer('Data')
  const mutation = await applyCellEditsToXlsx(
    base,
    [
      { sheetName: 'Data', row: 0, column: 0, writeValue: true, cell: { value: '00123' } },
      {
        sheetName: 'Data',
        row: 1,
        column: 0,
        writeValue: true,
        cell: { value: 12.5 },
        style: { bold: true, numberFormat: '0.00', fillColor: '#CCEEFF' },
      },
      { sheetName: 'Data', row: 2, column: 0, writeValue: true, cell: { value: '' } },
      { sheetName: 'Other', row: 0, column: 0, writeValue: true, cell: { value: 7 } },
      {
        sheetName: 'Other',
        row: 0,
        column: 1,
        writeValue: true,
        cell: { value: 25, formula: '=Data!A2*2' },
      },
    ],
    [],
    [],
    { renames: [], additions: [{ name: 'Other' }], removals: [], order: ['Data', 'Other'] },
  )
  const zip = await JSZip.loadAsync(mutation.buffer)
  zip.file('customXml/manual-sentinel.xml', '<preserve id="manual-fixture">unchanged</preserve>')
  // ZIP timestamps are otherwise generated from wall-clock time.
  zip.forEach((_path, entry) => {
    entry.date = new Date('2000-01-01T00:00:00Z')
  })
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}
