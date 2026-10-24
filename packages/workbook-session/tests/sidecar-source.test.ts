import { expect, it } from 'vitest'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { writeFile } from 'node:fs/promises'
import { XlsxSidecarClient } from '../src/sidecar/client'
import { SidecarWorkbookSource } from '../src/sidecar/source'
import { openWorkbookSnapshot } from '../src/open'
import { prepareCandidateMaterial } from '../src/candidate-material'
import { initialState } from '../src/state'
import { readCurrentRange } from '../src/read'
import { buildStructureFixture } from '../../../apps/sheets/tests/fixture-builder'
import { useSessionTestResources } from './helpers'
const resources = useSessionTestResources()
it('uses the real sidecar through the extracted client and shared schemas', async () => {
  const directory = await resources.directory(),
    path = join(directory, 'fixture.xlsx')
  await writeFile(path, await buildStructureFixture())
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
    const opened = await openWorkbookSnapshot({
      uri: new URL(`file://${path}`).href,
      storageDirectory: directory,
      source,
    })
    try {
      const state = initialState(opened.metadata)
      const read = await readCurrentRange({
        state,
        source,
        sessionId: opened.original.sidecarSessionId,
        logicalSheetId: state.sheets[0]!.id,
        range: 'A1:D2',
      })
      expect(read.cells[0]!.rawValue).toBe(1)
      expect(read.cells[7]).toMatchObject({ rawValue: 3, formula: '=SUM(A1:A2)' })
      const material = await prepareCandidateMaterial({
        state,
        source,
        sourceSessionId: opened.original.sidecarSessionId,
        operations: [{ op: 'set_cell', sheetId: state.sheets[0]!.id, address: 'A1', value: 99 }],
      })
      expect(material.state.sheets[0]!.cells.A1!.rawValue).toBe(1)
      expect(material.writableSheetIds).toContain(state.sheets[0]!.id)
      expect(material.state.sheets[0]!.features?.merges).toBeDefined()
    } finally {
      await opened.dispose()
    }
  } finally {
    client.stop()
  }
})
