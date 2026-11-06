import { expect, it } from 'vitest'
import { readFile, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import JSZip from 'jszip'
import { manualWorkbook } from '../../../apps/vscode-office/tests/fixtures/manual-workbook'
import { XlsxSidecarClient } from '../src/sidecar/client'
import { SidecarWorkbookSource } from '../src/sidecar/source'
import { openWorkbookSnapshot } from '../src/open'
import { initialState } from '../src/state'
import { prepareCandidateMaterial } from '../src/candidate-material'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import { SessionPackageWriter } from '../src/package-writer'
import { useSessionTestResources } from './helpers'
const resources = useSessionTestResources()
it('saves canonical workbook/scoped names and preserves hidden/print names and source cells', async () => {
  const directory = await resources.directory(),
    path = join(directory, 'names.xlsx')
  const zip = await JSZip.loadAsync(await manualWorkbook())
  const hidden = '<definedName name="HiddenPrivate" hidden="1">Data!$A$2</definedName>'
  const builtin =
    '<definedName name="_xlnm.Print_Area" localSheetId="0">Data!$A$1:$B$3</definedName>'
  const originalWorkbook = await zip.file('xl/workbook.xml')!.async('string')
  const section = `<definedNames><definedName name="Source">Data!$A$2</definedName>${hidden}${builtin}</definedNames>`
  zip.file('xl/workbook.xml', originalWorkbook.replace('</sheets>', '</sheets>' + section))
  const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
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
      const state = initialState(opened.metadata)
      const operations = [
        { op: 'delete_defined_name' as const, name: 'Source' },
        { op: 'add_defined_name' as const, name: 'Renamed', ref: 'Other!$A$1' },
        { op: 'add_defined_name' as const, name: 'Total', ref: 'SUM(Data!A2,Other!A1)' },
        {
          op: 'add_defined_name' as const,
          name: 'Total',
          ref: 'Data!$A$2',
          scopeSheetId: state.sheets[0]!.id,
        },
      ]
      const material = await prepareCandidateMaterial({
        state,
        operations,
        source,
        sourceSessionId: opened.original.sidecarSessionId,
      })
      expect(material.referenceRanges).toEqual({})
      const result = await computeCandidateState(
        state,
        {
          schemaVersion: 1,
          sessionId: 'names',
          transactionId: crypto.randomUUID(),
          expectedRevision: 0,
          origin: 'manual',
          summary: 'Manage names',
          operations,
        },
        material,
      )
      expect(result.inverse.state).toEqual(state)
      expect(result.nextState.sheets.map((sheet) => sheet.cells)).toEqual(
        state.sheets.map((sheet) => sheet.cells),
      )
      const writer = new SessionPackageWriter(client, opened.metadata)
      const target = join(directory, 'saved.xlsx')
      const staged = await writer.stage(opened.original.path, result.nextState, target)
      const saved = await JSZip.loadAsync(await readFile(staged.stagedPath))
      const xml = await saved.file('xl/workbook.xml')!.async('string')
      expect(xml).toContain(hidden)
      expect(xml).toContain(builtin)
      expect(xml).not.toContain('name="Source"')
      expect(xml).toContain('<definedName name="Total" localSheetId="0">Data!$A$2</definedName>')
      expect(xml).toContain('<definedName name="Renamed">Other!$A$1</definedName>')
      expect(await saved.file('xl/worksheets/sheet1.xml')!.async('string')).toBe(sheet)
      expect(await saved.file('customXml/manual-sentinel.xml')!.async('string')).toBe(
        '<preserve id="manual-fixture">unchanged</preserve>',
      )
      const reopened = await source.openSnapshot(staged.stagedPath)
      try {
        expect(reopened.metadata.definedNames).toEqual(
          expect.arrayContaining([
            { name: 'Renamed', formula: 'Other!$A$1' },
            { name: 'Total', formula: 'SUM(Data!A2,Other!A1)' },
            { name: 'Total', formula: 'Data!$A$2', sheetIndex: 0 },
          ]),
        )
      } finally {
        await source.close(reopened.sessionId)
      }
    } finally {
      await opened.dispose()
    }
  } finally {
    client.stop()
  }
})
