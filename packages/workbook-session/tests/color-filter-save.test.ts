import { expect, it } from 'vitest'
import { readFile, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import JSZip from 'jszip'
import { manualWorkbook } from '../../../apps/vscode-office/tests/fixtures/manual-workbook'
import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import type { ColorFilterCriteria } from '@genoffice/xlsx-gateway/domain/color-filter'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { XlsxSidecarClient } from '../src/sidecar/client'
import { SidecarWorkbookSource } from '../src/sidecar/source'
import { openWorkbookSnapshot } from '../src/open'
import { initialState } from '../src/state'
import { prepareCandidateMaterial } from '../src/candidate-material'
import { readCurrentRange } from '../src/read'
import { SessionPackageWriter } from '../src/package-writer'
import { useSessionTestResources } from './helpers'
const resources = useSessionTestResources()

it('saves single/multiple fill and font criteria and reopens them with the real sidecar and untouched source cells', async () => {
  const directory = await resources.directory(),
    path = join(directory, 'colors.xlsx')
  const fixture = await applyCellEditsToXlsx(await manualWorkbook(), [
    { sheetName: 'Data', row: 0, column: 0, writeValue: true, cell: { value: 'Color' } },
    {
      sheetName: 'Data',
      row: 1,
      column: 0,
      writeValue: true,
      cell: { value: 'Red' },
      style: { fillColor: '#FF0000', fontColor: '#0000FF' },
    },
    {
      sheetName: 'Data',
      row: 2,
      column: 0,
      writeValue: true,
      cell: { value: 'Blue' },
      style: { fillColor: '#0000FF', fontColor: '#FF0000' },
    },
    { sheetName: 'Data', row: 3, column: 0, writeValue: true, cell: { value: 'Plain' } },
    {
      sheetName: 'Data',
      row: 4,
      column: 0,
      writeValue: true,
      cell: { value: 'White' },
      style: { fillColor: '#FFFFFF', fontColor: '#0000FF' },
    },
  ])
  await writeFile(path, fixture.buffer)
  const client = new XlsxSidecarClient(
    fileURLToPath(
      new URL(
        '../../../apps/sheets/native/xlsx-engine/target/release/xlsx-sidecar' +
          (process.platform === 'win32' ? '.exe' : ''),
        import.meta.url,
      ),
    ),
  )
  try {
    const source = new SidecarWorkbookSource(client)
    const opened = await openWorkbookSnapshot({
      uri: pathToFileURL(path).href,
      storageDirectory: directory,
      source,
    })
    try {
      const state = initialState(opened.metadata),
        sheetId = state.sheets[0]!.id
      const cases: [ColorFilterCriteria, number[]][] = [
        [{ fill: ['#FF0000'] }, [2, 3, 4]],
        [{ fill: ['#FF0000', '#0000FF', null] }, [4]],
        [{ fill: [null] }, [1, 2, 4]],
        [{ font: ['#000000', '#0000FF'] }, [2]],
        [{ font: ['#FF0000'] }, [1, 3, 4]],
      ]
      for (const [index, [colors, hidden]] of cases.entries()) {
        const operations: WorkbookOperation[] = [
          { op: 'set_filter', sheetId, range: 'A1:A5' },
          { op: 'set_filter_criteria', sheetId, column: 'A', colors },
        ]
        const material = await prepareCandidateMaterial({
          state,
          operations,
          source,
          sourceSessionId: opened.original.sidecarSessionId,
        })
        const result = await computeCandidateState(
          state,
          {
            schemaVersion: 1,
            sessionId: 'colors',
            transactionId: crypto.randomUUID(),
            expectedRevision: 0,
            origin: 'manual',
            summary: 'Color filter',
            operations,
          },
          material,
        )
        expect(result.nextState.sheets[0]!.features!.filterEdits).toMatchObject({
          hiddenRows: hidden,
          filter: { columns: [{ colId: 0, colors }] },
        })
        const staged = await new SessionPackageWriter(client, opened.metadata).stage(
          opened.original.path,
          result.nextState,
          join(directory, `saved-${index}.xlsx`),
        )
        const zip = await JSZip.loadAsync(await readFile(staged.stagedPath))
        const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
        expect(xml).toContain('urn:genoffice:filter-colors:v1')
        expect(await zip.file('customXml/manual-sentinel.xml')!.async('string')).toBe(
          '<preserve id="manual-fixture">unchanged</preserve>',
        )
        // For a single color, remove our extension to verify standard Excel DXF import independently.
        if (('fill' in colors ? colors.fill : colors.font).length === 1) {
          expect(xml).toContain('<colorFilter dxfId=')
          zip.file(
            'xl/worksheets/sheet1.xml',
            xml.replace(
              /<extLst><ext uri="urn:genoffice:filter-colors:v1">[\s\S]*?<\/ext><\/extLst>/,
              '',
            ),
          )
          await writeFile(staged.stagedPath, await zip.generateAsync({ type: 'nodebuffer' }))
        }
        const reopened = await source.openSnapshot(staged.stagedPath)
        try {
          const next = initialState(reopened.metadata),
            id = next.sheets[0]!.id
          const read = await readCurrentRange({
            state: next,
            sessionId: reopened.sessionId,
            source,
            logicalSheetId: id,
            range: 'A1:A5',
            includeMerges: true,
          })
          expect(read).toMatchObject({ autoFilterColumns: [{ colId: 0, colors }] })
          expect(
            'rows' in read ? read.rows.filter((row) => row.hidden).map((row) => row.row) : [],
          ).toEqual(hidden)
          const cells = new Map(read.cells.map((cell) => [cell.address, cell]))
          expect(cells.get('A2')!.rawValue).toBe('Red')
          expect(cells.get('A2')!.style!.fillColor).toBe('#FF0000')
          expect(cells.get('A3')!.style!.fontColor).toBe('#FF0000')
          expect(cells.get('A5')!.style!.fillColor).toBe('#FFFFFF')
        } finally {
          await source.close(reopened.sessionId)
        }
      }
    } finally {
      await opened.dispose()
    }
  } finally {
    client.stop()
  }
}, 30_000)
