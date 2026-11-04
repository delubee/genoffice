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
it('saves protected-sheet formula view without touching protection, source cell XML or package sentinel', async () => {
  const directory = await resources.directory(),
    path = join(directory, 'protected.xlsx')
  const zip = await JSZip.loadAsync(await manualWorkbook())
  const protection = '<sheetProtection sheet="1" objects="1" scenarios="1" password="ABCD"/>'
  const originalXml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  zip.file(
    'xl/worksheets/sheet1.xml',
    originalXml.replace('</sheetData>', '</sheetData>' + protection),
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
      const state = initialState(opened.metadata)
      const operations = [
        {
          op: 'set_page_setup' as const,
          sheetId: state.sheets[0]!.id,
          showFormulas: true,
          showGridlines: false,
          showHeadings: false,
          zoomScale: 120,
        },
      ]
      const material = await prepareCandidateMaterial({
        state,
        operations,
        source,
        sourceSessionId: opened.original.sidecarSessionId,
      })
      expect(material.writableSheetIds).not.toContain(state.sheets[0]!.id)
      expect(material.viewWritableSheetIds).toContain(state.sheets[0]!.id)
      expect(material.state.sheets[0]!.protected).toBe(true)
      expect(material.referenceRanges).toEqual({})
      const result = await computeCandidateState(
        state,
        {
          schemaVersion: 1,
          sessionId: 'protected',
          transactionId: crypto.randomUUID(),
          expectedRevision: 0,
          origin: 'manual',
          summary: 'Formula view',
          operations,
        },
        material,
      )
      expect(result.inverse.state).toEqual(state)
      expect(result.nextState.sheets[0]!.cells).toEqual(state.sheets[0]!.cells)
      const writer = new SessionPackageWriter(client, opened.metadata)
      const staged = await writer.stage(
        opened.original.path,
        result.nextState,
        join(directory, 'saved.xlsx'),
      )
      const saved = await JSZip.loadAsync(await readFile(staged.stagedPath))
      const xml = await saved.file('xl/worksheets/sheet1.xml')!.async('string')
      expect(xml).toContain(protection)
      expect(xml).toContain('showFormulas="1"')
      expect(xml.match(/<sheetData\b[^>]*>[\s\S]*?<\/sheetData>/)![0]).toBe(
        originalXml.match(/<sheetData\b[^>]*>[\s\S]*?<\/sheetData>/)![0],
      )
      expect(await saved.file('customXml/manual-sentinel.xml')!.async('string')).toBe(
        '<preserve id="manual-fixture">unchanged</preserve>',
      )
      const reopened = await source.openSnapshot(staged.stagedPath)
      try {
        expect(reopened.metadata.sheets[0]).toMatchObject({
          showFormulas: true,
          showGridLines: false,
          showRowColHeaders: false,
          zoomScale: 120,
        })
        // The reopened source indexes its features in the background: wait for them.
        let read = await source.readFeatures(reopened.sessionId, reopened.metadata.sheets[0]!.id)
        for (let attempt = 0; attempt < 50 && !read.indexingComplete; attempt++) {
          await new Promise((resolve) => setTimeout(resolve, 100))
          read = await source.readFeatures(reopened.sessionId, reopened.metadata.sheets[0]!.id)
        }
        expect(read).toMatchObject({ sheetProtection: { protected: true } })
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
