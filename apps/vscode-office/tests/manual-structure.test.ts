import { expect, it } from 'vitest'
import { writeFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import { join } from 'node:path'
import { manualWorkbook } from './fixtures/manual-workbook'
import { withTestWorkspace } from './helpers'
import { XlsxSidecarClient } from '@genoffice/workbook-session/sidecar/client'
import { SidecarWorkbookSource } from '@genoffice/workbook-session/sidecar/source'
import { openWorkbookSnapshot } from '@genoffice/workbook-session/open'
import { initialState } from '@genoffice/workbook-session/state'
import { prepareCandidateMaterial } from '@genoffice/workbook-session/candidate-material'
import { readCurrentRange } from '@genoffice/workbook-session/read'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { SessionPackageWriter } from '@genoffice/workbook-session/package-writer'

it('replays an ordered structural log from the original on repeated saves with stable logical IDs', async () => {
  await withTestWorkspace(async (directory) => {
    const client = new XlsxSidecarClient(
        fileURLToPath(
          new URL('../../sheets/native/xlsx-engine/target/release/xlsx-sidecar', import.meta.url),
        ),
      ),
      source = new SidecarWorkbookSource(client),
      path = join(directory, 'structure.xlsx')
    await writeFile(path, await manualWorkbook())
    const opened = await openWorkbookSnapshot({
      uri: new URL(`file://${path}`).href,
      storageDirectory: directory,
      source,
    })
    try {
      let state = initialState(opened.metadata),
        revision = 0
      const dataId = state.sheets[0]!.id,
        otherId = state.sheets[1]!.id
      const apply = async (operations: WorkbookOperation[]) => {
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
            sessionId: 'structure',
            transactionId: `structure-${revision}`,
            expectedRevision: revision++,
            origin: 'manual',
            summary: 'Structure',
            operations,
          },
          material,
        )
        state = result.nextState
      }
      await apply([{ op: 'insert_rows', sheetId: dataId, row: 2, count: 1 }])
      await apply([{ op: 'rename_sheet', sheetId: dataId, name: 'Facts' }])
      await apply([{ op: 'add_sheet', name: 'Scratch' }])
      const scratchId = state.sheets.find((sheet) => sheet.name === 'Scratch')!.id
      await apply([{ op: 'delete_sheet', sheetId: scratchId }])
      expect(state.sheets.map((sheet) => sheet.id)).toEqual([dataId, otherId])
      expect(state.sheetMap[scratchId]).toBeDefined()
      const writer = new SessionPackageWriter(client, opened.metadata)
      for (const filename of ['first.xlsx', 'again.xlsx']) {
        const target = join(directory, filename),
          staged = await writer.stage(opened.original.path, state, target)
        await writer.publish(staged.stagedPath, target, null)
        const reopened = await openWorkbookSnapshot({
          uri: new URL(`file://${target}`).href,
          storageDirectory: directory,
          source,
        })
        try {
          const loaded = initialState(reopened.metadata)
          expect(loaded.sheets.map((sheet) => sheet.name)).toEqual(['Facts', 'Other'])
          const cells = await readCurrentRange({
            state: loaded,
            source,
            sessionId: reopened.original.sidecarSessionId,
            logicalSheetId: loaded.sheets[0]!.id,
            range: 'A1:A4',
          })
          expect(cells.cells.map((cell) => cell.rawValue)).toEqual(['00123', null, 12.5, ''])
          expect(cells.cells[2]!.style?.bold).toBe(true)
          const formula = await readCurrentRange({
            state: loaded,
            source,
            sessionId: reopened.original.sidecarSessionId,
            logicalSheetId: loaded.sheets[1]!.id,
            range: 'B1',
          })
          expect(formula.cells[0]!.formula).toBe('=Facts!A3*2')
        } finally {
          await reopened.dispose()
        }
      }
    } finally {
      await opened.dispose()
      client.stop()
    }
  })
})
