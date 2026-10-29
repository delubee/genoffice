import { expect, it } from 'vitest'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { writeFile } from 'node:fs/promises'
import JSZip from 'jszip'
import { blankXlsxBuffer } from '@genoffice/xlsx-gateway/gateway/csv-import'
import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { XlsxSidecarClient } from '../src/sidecar/client'
import { SidecarWorkbookSource } from '../src/sidecar/source'
import { openWorkbookSnapshot } from '../src/open'
import { initialState } from '../src/state'
import { buildMergeStages } from '../src/merge-workbooks'
import { prepareStagedCandidate } from '../src/staged-candidate'
import { useSessionTestResources } from './helpers'
const resources = useSessionTestResources()

async function sourceWorkbook() {
  const mutation = await applyCellEditsToXlsx(
    await blankXlsxBuffer('Data'),
    [
      { sheetName: 'Data', row: 0, column: 0, writeValue: true, cell: { value: 'Region' } },
      { sheetName: 'Data', row: 0, column: 2, writeValue: true, cell: { value: 'Merged' } },
      {
        sheetName: 'Data',
        row: 1,
        column: 1,
        writeValue: true,
        cell: { value: 12.5 },
        style: { bold: true, numberFormat: '0.00', fillColor: '#CCEEFF' },
      },
      { sheetName: 'Other', row: 0, column: 0, writeValue: true, cell: { value: 7 } },
      {
        sheetName: 'Other',
        row: 0,
        column: 1,
        writeValue: true,
        cell: { value: 14, formula: '=A1*2' },
      },
    ],
    [],
    [],
    { renames: [], additions: [{ name: 'Other' }], removals: [], order: ['Data', 'Other'] },
  )
  const zip = await JSZip.loadAsync(mutation.buffer)
  const path = 'xl/worksheets/sheet1.xml'
  let xml = await zip.file(path)!.async('string')
  xml = xml
    .replace(
      '<sheetData>',
      '<cols><col min="2" max="2" width="21.5" customWidth="1"/></cols><sheetData>',
    )
    .replace(/<row r="2"([^>]*)>/, '<row r="2"$1 ht="30" customHeight="1">')
    .replace(
      '</sheetData>',
      '</sheetData><mergeCells count="1"><mergeCell ref="C1:D1"/></mergeCells>',
    )
  zip.file(path, xml)
  // Excel stores the formula result; the merge imports that cached value.
  const other = 'xl/worksheets/sheet2.xml'
  zip.file(
    other,
    (await zip.file(other)!.async('string')).replace(
      /(<f>A1\*2<\/f>)(<v>[^<]*<\/v>)?/,
      '$1<v>14</v>',
    ),
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

it('builds staged sheets with values, styles, merges and sizes from a real source workbook', async () => {
  const directory = await resources.directory()
  const targetPath = join(directory, 'target.xlsx'),
    sourcePath = join(directory, 'source.xlsx')
  await writeFile(targetPath, await blankXlsxBuffer('Data'))
  await writeFile(sourcePath, await sourceWorkbook())
  const client = new XlsxSidecarClient(
    fileURLToPath(
      new URL(
        '../../../apps/sheets/native/xlsx-engine/target/release/xlsx-sidecar',
        import.meta.url,
      ),
    ),
  )
  try {
    const source = new SidecarWorkbookSource(client)
    const target = await openWorkbookSnapshot({
      uri: new URL(`file://${targetPath}`).href,
      storageDirectory: directory,
      source,
    })
    const merged = await openWorkbookSnapshot({
      uri: new URL(`file://${sourcePath}`).href,
      storageDirectory: directory,
      source,
    })
    try {
      const state = initialState(target.metadata)
      const built = await buildMergeStages({
        target: state,
        sources: [
          {
            fileName: 'source.xlsx',
            state: initialState(merged.metadata),
            sessionId: merged.original.sidecarSessionId,
          },
        ],
        source,
      })
      expect(built.sheetNames).toEqual(['Data (2)', 'Other'])
      const changes = await prepareStagedCandidate({
        state,
        intent: {
          schemaVersion: 1,
          sessionId: 'session',
          transactionId: 'merge',
          expectedRevision: 0,
          origin: 'manual',
          summary: 'Merge workbooks',
          operations: built.stages.flat(),
        },
        stages: built.stages,
        source,
        sourceSessionId: target.original.sidecarSessionId,
      })
      const sheets = changes.nextState.sheets
      expect(sheets.map((sheet) => sheet.name)).toEqual(['Data', 'Data (2)', 'Other'])
      const data = sheets[1]!,
        other = sheets[2]!
      expect(data.cells.A1!.rawValue).toBe('Region')
      expect(data.cells.C1!.rawValue).toBe('Merged')
      expect(data.cells.B2).toMatchObject({
        rawValue: 12.5,
        numberFormat: '0.00',
        style: expect.objectContaining({ bold: true, fillColor: '#CCEEFF' }),
      })
      expect(other.cells.B1!.rawValue).toBe(14)
      expect(other.cells.B1!.formula).toBeUndefined()
      expect(data.features?.mergeEdits).toEqual([
        { startRow: 0, endRow: 0, startColumn: 2, endColumn: 3 },
      ])
      expect(data.features?.columnWidthsCharacters).toMatchObject({ B: 21.5 })
      expect(data.features?.rowHeights).toMatchObject({ 2: 30 })
      expect(sheets[0]!.cells).toEqual(state.sheets[0]!.cells)
    } finally {
      await merged.dispose()
      await target.dispose()
    }
  } finally {
    client.stop()
  }
})
