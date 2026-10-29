import { expect, it } from 'vitest'
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

it('reads hidden/local definitions and calculation settings privately from the owning immutable source session', async () => {
  const directory = await resources.directory(),
    path = join(directory, 'private-context.xlsx')
  const zip = await JSZip.loadAsync(await manualWorkbook())
  const workbook = await zip.file('xl/workbook.xml')!.async('string')
  zip.file(
    'xl/workbook.xml',
    workbook
      .replace(/<definedNames\b[\s\S]*?<\/definedNames>/g, '')
      .replace(/<calcPr\b[^>]*\/>/g, '')
      .replace(
        '</workbook>',
        `<definedNames>
    <definedName name="Rate">Data!$A$1</definedName>
    <definedName name="Rate" localSheetId="1">Other!$A$1</definedName>
    <definedName name="Secret" hidden="1" localSheetId="0"><![CDATA[IF(1<2,"private & formula",0)]]></definedName>
    <definedName name="Empty" hidden="1"/>
    <definedName name="_xlnm.Print_Area" localSheetId="0">Data!$A$1:$B$2</definedName>
    </definedNames><calcPr calcMode="manual" fullCalcOnLoad="1" iterate="1" iterateCount="25"/></workbook>`,
      ),
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
      const state = structuredClone(initialState(opened.metadata))
      state.sheets.reverse()
      state.sheets[1]!.name = 'Renamed'
      state.structuralLog.push({
        op: 'rename_sheet',
        sheetId: state.sheets[1]!.id,
        name: 'Renamed',
      })
      const result = await loadCalculationSource({
        state,
        source,
        sourceSessionId: opened.original.sidecarSessionId,
        isCurrent: () => true,
      })
      const context = result.sourceContext
      expect(context.names).toEqual([
        { name: 'Rate', formula: 'Data!$A$1', hidden: false },
        {
          name: 'Rate',
          formula: 'Other!$A$1',
          hidden: false,
          sheetId: state.sheets[0]!.originalSheetId,
        },
        {
          name: 'Secret',
          formula: 'IF(1<2,"private & formula",0)',
          hidden: true,
          sheetId: state.sheets[1]!.originalSheetId,
        },
        { name: 'Empty', formula: '', hidden: true },
        {
          name: '_xlnm.Print_Area',
          formula: 'Data!$A$1:$B$2',
          hidden: false,
          sheetId: state.sheets[1]!.originalSheetId,
        },
      ])
      expect(context.calculationProperties).toMatchObject({
        calcMode: 'manual',
        iterate: '1',
        iterateCount: '25',
        fullCalcOnLoad: '1',
      })
      expect(JSON.stringify(opened.metadata)).not.toContain('private & formula')
      expect(JSON.stringify(state)).not.toContain('private & formula')
      expect(result.sheets[0]!.cells.B1!.formula).toBe('=Renamed!A2*2')
      await expect(source.readCalculationContext(crypto.randomUUID())).rejects.toMatchObject({
        code: 'session_mismatch',
      })
    } finally {
      await opened.dispose()
    }
    await expect(
      source.readCalculationContext(opened.original.sidecarSessionId),
    ).rejects.toMatchObject({ code: 'session_mismatch' })
  } finally {
    client.stop()
  }
})
