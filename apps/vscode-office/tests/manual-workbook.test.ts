import { expect, it } from 'vitest'
import { createHash } from 'node:crypto'
import { writeFile, readFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import { join } from 'node:path'
import JSZip from 'jszip'
import { manualWorkbook, MANUAL_WORKBOOK_SHA256 } from './fixtures/manual-workbook'
import { withTestWorkspace } from './helpers'
import { XlsxSidecarClient } from '@genoffice/workbook-session/sidecar/client'
import { SidecarWorkbookSource } from '@genoffice/workbook-session/sidecar/source'
import { openWorkbookSnapshot } from '@genoffice/workbook-session/open'
import { initialState } from '@genoffice/workbook-session/state'
import { readCurrentRange } from '@genoffice/workbook-session/read'
import { SessionPackageWriter } from '@genoffice/workbook-session/package-writer'

it('preserves original package bytes while saving and reopening typed manual edits and styles', async () => {
  await withTestWorkspace(async (directory) => {
    const sourcePath = join(directory, 'manual.xlsx'),
      original = await manualWorkbook()
    await writeFile(sourcePath, original)
    const sourceDigest = createHash('sha256').update(original).digest('hex')
    expect(sourceDigest).toBe(MANUAL_WORKBOOK_SHA256)
    const client = new XlsxSidecarClient(
      fileURLToPath(
        new URL('../../sheets/native/xlsx-engine/target/release/xlsx-sidecar', import.meta.url),
      ),
    )
    try {
      const source = new SidecarWorkbookSource(client)
      const opened = await openWorkbookSnapshot({
        uri: new URL(`file://${sourcePath}`).href,
        source,
        storageDirectory: directory,
      })
      try {
        const state = initialState(opened.metadata),
          current = structuredClone(state)
        expect(state.sheets).toHaveLength(2)
        current.stateId = 'manual-edit'
        current.sheets[0]!.cells.A2 = {
          kind: 'number',
          rawValue: 42.5,
          style: { bold: true, fillColor: '#CCEEFF', numberFormat: '0.00' },
          numberFormat: '0.00',
        }
        current.sheets[1]!.cells.C1 = {
          kind: 'number',
          rawValue: 123,
          formula: '=Data!A2*3',
          calculationState: 'pending',
        }
        const writer = new SessionPackageWriter(client, opened.metadata)
        const target = join(directory, 'copy.xlsx')
        const staged = await writer.stage(opened.original.path, current, target)
        expect(
          createHash('sha256')
            .update(await readFile(sourcePath))
            .digest('hex'),
        ).toBe(sourceDigest)
        await writer.publish(staged.stagedPath, target, null)
        const reopened = await openWorkbookSnapshot({
          uri: new URL(`file://${target}`).href,
          source,
          storageDirectory: directory,
        })
        try {
          const loaded = initialState(reopened.metadata)
          const read = await readCurrentRange({
            state: loaded,
            source,
            sessionId: reopened.original.sidecarSessionId,
            logicalSheetId: loaded.sheets[0]!.id,
            range: 'A1:A3',
          })
          expect(read.cells.map((cell) => cell.rawValue)).toEqual(['00123', 42.5, ''])
          expect(read.cells[1]!.numberFormat).toBe('0.00')
          expect(read.cells[1]!.style?.bold).toBe(true)
          const formulas = await readCurrentRange({
            state: loaded,
            source,
            sessionId: reopened.original.sidecarSessionId,
            logicalSheetId: loaded.sheets[1]!.id,
            range: 'B1:C1',
          })
          expect(formulas.cells[0]).toMatchObject({
            formula: '=Data!A2*2',
            rawValue: null,
            calculationState: 'pending',
          })
          expect(formulas.cells[1]).toMatchObject({
            formula: '=Data!A2*3',
            rawValue: null,
            calculationState: 'pending',
          })
          const zip = await JSZip.loadAsync(await readFile(target))
          expect(await zip.file('customXml/manual-sentinel.xml')!.async('text')).toBe(
            '<preserve id="manual-fixture">unchanged</preserve>',
          )
        } finally {
          await reopened.dispose()
        }
      } finally {
        await opened.dispose()
      }
    } finally {
      client.stop()
    }
  })
})
