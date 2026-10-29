import { expect, it, vi } from 'vitest'
import { writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import JSZip from 'jszip'
import { manualWorkbook } from '../../../apps/vscode-office/tests/fixtures/manual-workbook'
import { XlsxSidecarClient } from '../src/sidecar/client'
import { SidecarWorkbookSource } from '../src/sidecar/source'
import { openWorkbookSnapshot } from '../src/open'
import { initialState } from '../src/state'
import { loadCalculationSource } from '../src/calculation/source'
import { useSessionTestResources } from './helpers'
const resources = useSessionTestResources()

it('hydrates an actual XLSX through the validated native protocol, including far styled blanks and overridden array anchors', async () => {
  const directory = await resources.directory(),
    path = join(directory, 'calculation-source.xlsx')
  const zip = await JSZip.loadAsync(await manualWorkbook())
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetData>
    <row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>A1*2</f><v>4</v></c></row>
    <row r="100001"><c r="B100001" s="1"/><c r="C100001" t="inlineStr"><is><t></t></is></c><c r="D100001" t="e"><v>#REF!</v></c></row>
    <row r="100002"><c r="A100002"><f t="array" ref="A100002:B100003">SEQUENCE(2,2)</f><v>1</v></c><c r="B100002"><v>2</v></c></row>
    <row r="100003"><c r="A100003"><v>3</v></c><c r="B100003"><v>4</v></c></row>
    </sheetData></worksheet>`,
  )
  await writeFile(path, await zip.generateAsync({ type: 'uint8array' }))
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
      const state = structuredClone(initialState(opened.metadata)),
        sheet = state.sheets[0]!
      // Also prove loader bounds do not depend on metadata having detected the far rows.
      sheet.features!.originalMetadata = {
        ...(sheet.features!.originalMetadata as object),
        rowCount: 1,
        columnCount: 1,
      }
      sheet.cells.A100002 = { kind: 'number', rawValue: 99 }
      const before = structuredClone(state)
      const result = await loadCalculationSource({
        state,
        source,
        sourceSessionId: opened.original.sidecarSessionId,
        isCurrent: () => true,
        readBatchCells: 8,
      })
      expect(state).toEqual(before)
      const loaded = result.sheets[0]!
      expect(loaded.cells.A1).toMatchObject({ rawValue: 2, kind: 'number' })
      expect(loaded.cells.B1).toMatchObject({
        formula: '=A1*2',
        rawValue: 4,
        cachedValue: 4,
        calculationState: 'pending',
      })
      expect(loaded.cells.B100001).toMatchObject({
        kind: 'blank',
        rawValue: null,
        numberFormat: '0.00',
        style: { bold: true },
      })
      expect(loaded.cells.C100001).toMatchObject({ kind: 'string', rawValue: '' })
      expect(loaded.cells.D100001).toMatchObject({ kind: 'error', rawValue: '#REF!' })
      expect(loaded.cells.A100002).toEqual({ kind: 'number', rawValue: 99 })
      expect(loaded.cells.B100003).toMatchObject({ rawValue: 4 })
      expect(loaded.sourceArrays).toEqual([
        { sourceAddress: 'A100002', sourceRange: 'A100002:B100003', logicalAddress: 'A100002' },
      ])
      expect(Object.keys(loaded.cells)).toHaveLength(9)
      expect(loaded.rows).toBe(100003)
      expect(result.sheets[1]!.cells.B1).toMatchObject({
        formula: '=Data!A2*2',
        rawValue: null,
        calculationState: 'pending',
      })
    } finally {
      await opened.dispose()
    }
  } finally {
    client.stop()
  }
})

it('loads more than one native enumeration page in bounded batches after row insertion', async () => {
  const directory = await resources.directory(),
    path = join(directory, 'paged-calculation.xlsx')
  const zip = await JSZip.loadAsync(await manualWorkbook())
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1"/><sheetData>${Array.from({ length: 10005 }, (_, index) => `<row r="${index + 1}"><c r="A${index + 1}"><v>${index}</v></c></row>`).join('')}<row r="100001"><c r="B100001" s="1"/></row></sheetData></worksheet>`,
  )
  await writeFile(path, await zip.generateAsync({ type: 'uint8array' }))
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
    const find = vi.spyOn(source, 'findBaseCells')
    const opened = await openWorkbookSnapshot({
      uri: pathToFileURL(path).href,
      storageDirectory: directory,
      source,
    })
    try {
      const state = structuredClone(initialState(opened.metadata)),
        sheet = state.sheets[0]!
      state.structuralLog.push({ op: 'insert_rows', sheetId: sheet.id, row: 1, count: 1 })
      sheet.cells.A1 = { kind: 'number', rawValue: 100 }
      sheet.cells.A2 = { kind: 'number', rawValue: -1 }
      const result = await loadCalculationSource({
        state,
        source,
        sourceSessionId: opened.original.sidecarSessionId,
        isCurrent: () => true,
      })
      const loaded = result.sheets[0]!
      expect(Object.keys(loaded.cells)).toHaveLength(10007)
      expect(loaded.cells.A1!.rawValue).toBe(100)
      expect(loaded.cells.A2!.rawValue).toBe(-1)
      expect(loaded.cells.A10006!.rawValue).toBe(10004)
      expect(loaded.cells.B100002).toMatchObject({ kind: 'blank', numberFormat: '0.00' })
      expect(
        find.mock.calls.some(
          ([request]) =>
            request.sheetId === sheet.originalSheetId && request.resumeAt !== undefined,
        ),
      ).toBe(true)
      expect(loaded.rows).toBe(100002)
    } finally {
      await opened.dispose()
    }
  } finally {
    client.stop()
  }
})
