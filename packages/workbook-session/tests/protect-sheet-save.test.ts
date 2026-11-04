import { expect, it } from 'vitest'
import { readFile, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import JSZip from 'jszip'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { manualWorkbook } from '../../../apps/vscode-office/tests/fixtures/manual-workbook'
import { XlsxSidecarClient } from '../src/sidecar/client'
import { SidecarWorkbookSource } from '../src/sidecar/source'
import { openWorkbookSnapshot } from '../src/open'
import { initialState } from '../src/state'
import { prepareCandidateMaterial } from '../src/candidate-material'
import { SessionPackageWriter } from '../src/package-writer'
import type { WorkbookState } from '../src/domain'
import { useSessionTestResources } from './helpers'
const resources = useSessionTestResources()

it(
  'protects with a deterministic password hash, blocks edits, saves, verifies and unprotects',
  { timeout: 180_000 },
  async () => {
    const directory = await resources.directory(),
      path = join(directory, 'protect.xlsx')
    await writeFile(path, await manualWorkbook())
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
      const run = async (
        state: WorkbookState,
        sessionId: string,
        operations: WorkbookOperation[],
        transactionId = 'protect-tx',
      ) =>
        computeCandidateState(
          state,
          {
            schemaVersion: 1,
            sessionId: 'protect',
            transactionId,
            expectedRevision: 0,
            origin: 'manual',
            summary: 'Protection',
            operations,
          },
          await prepareCandidateMaterial({ state, operations, source, sourceSessionId: sessionId }),
        )
      const opened = await openWorkbookSnapshot({
        uri: pathToFileURL(path).href,
        storageDirectory: directory,
        source,
      })
      const savedPath = join(directory, 'protected.xlsx')
      try {
        const state = initialState(opened.metadata),
          sheetId = state.sheets[0]!.id
        const protect: WorkbookOperation[] = [
          {
            op: 'protect_sheet',
            sheetId,
            protected: true,
            password: 'open sesame',
            allow: { sort: true },
          },
        ]
        const first = await run(state, opened.original.sidecarSessionId, protect)
        const again = await run(state, opened.original.sidecarSessionId, protect)
        // The view candidate and host execution must agree byte for byte.
        expect(JSON.stringify(again.nextState.sheets[0]!.features?.protectionEdits)).toBe(
          JSON.stringify(first.nextState.sheets[0]!.features?.protectionEdits),
        )
        expect(first.nextState.sheets[0]!.protected).toBe(true)
        expect(first.nextState.sheets[0]!.features?.protectionEdits).toMatchObject({
          protected: true,
          hasPassword: true,
          allow: expect.objectContaining({ sort: true, selectLockedCells: true }),
        })
        await expect(
          run(first.nextState, opened.original.sidecarSessionId, [
            { op: 'set_cell', sheetId, address: 'B2', value: 1 },
          ]),
        ).rejects.toThrow()
        // Allow Edit Ranges open cells of the protected sheet, as the desktop guard does.
        // Excel edits Allow Edit Ranges while the sheet is unprotected, then protects it.
        const open = await run(state, opened.original.sidecarSessionId, [
          { op: 'set_protected_ranges', sheetId, ranges: [{ name: 'Inputs', sqref: 'B2:C3' }] },
        ])
        const ranged = await run(open.nextState, opened.original.sidecarSessionId, protect)
        const edited = await run(ranged.nextState, opened.original.sidecarSessionId, [
          { op: 'set_cell', sheetId, address: 'C3', value: 7 },
        ])
        expect(edited.nextState.sheets[0]!.cells.C3!.rawValue).toBe(7)
        await expect(
          run(ranged.nextState, opened.original.sidecarSessionId, [
            { op: 'set_range', sheetId, start: 'C3', values: [[1, 2]] },
          ]),
        ).rejects.toThrow(/protected_target|unverified_writable_target/)
        const staged = await new SessionPackageWriter(client, opened.metadata).stage(
          opened.original.path,
          first.nextState,
          join(directory, 'stage-1.xlsx'),
        )
        const xml = await (
          await JSZip.loadAsync(await readFile(staged.stagedPath))
        )
          .file('xl/worksheets/sheet1.xml')!
          .async('string')
        expect(xml).toMatch(/<sheetProtection\b[^>]*sheet="1"/)
        expect(xml).toMatch(/hashValue="/)
        expect(xml).toMatch(/sort="0"/)
        await writeFile(savedPath, await readFile(staged.stagedPath))
      } finally {
        await opened.dispose()
      }
      const reopened = await openWorkbookSnapshot({
        uri: pathToFileURL(savedPath).href,
        storageDirectory: directory,
        source,
      })
      try {
        const state = initialState(reopened.metadata),
          sheetId = state.sheets[0]!.id
        await expect(
          run(state, reopened.original.sidecarSessionId, [
            { op: 'protect_sheet', sheetId, protected: false },
          ]),
        ).rejects.toThrow('protection:password_required')
        await expect(
          run(state, reopened.original.sidecarSessionId, [
            { op: 'protect_sheet', sheetId, protected: false, password: 'wrong' },
          ]),
        ).rejects.toThrow('protection:wrong_password')
        const open = await run(state, reopened.original.sidecarSessionId, [
          { op: 'protect_sheet', sheetId, protected: false, password: 'open sesame' },
        ])
        expect(open.nextState.sheets[0]!.protected).toBe(false)
        const edited = await run(open.nextState, reopened.original.sidecarSessionId, [
          { op: 'set_cell', sheetId, address: 'B2', value: 1 },
        ])
        expect(edited.nextState.sheets[0]!.cells.B2!.rawValue).toBe(1)
        const staged = await new SessionPackageWriter(client, reopened.metadata).stage(
          reopened.original.path,
          edited.nextState,
          join(directory, 'stage-2.xlsx'),
        )
        const xml = await (
          await JSZip.loadAsync(await readFile(staged.stagedPath))
        )
          .file('xl/worksheets/sheet1.xml')!
          .async('string')
        expect(xml).not.toMatch(/<sheetProtection\b/)
      } finally {
        await reopened.dispose()
      }
    } finally {
      client.stop()
    }
  },
)

it('locks workbook structure and saves allow-edit ranges through the session', async () => {
  const directory = await resources.directory(),
    path = join(directory, 'lock.xlsx')
  await writeFile(path, await manualWorkbook())
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
      const run = async (state: WorkbookState, operations: WorkbookOperation[]) =>
        computeCandidateState(
          state,
          {
            schemaVersion: 1,
            sessionId: 'lock',
            transactionId: crypto.randomUUID(),
            expectedRevision: 0,
            origin: 'manual',
            summary: 'Review',
            operations,
          },
          await prepareCandidateMaterial({
            state,
            operations,
            source,
            sourceSessionId: opened.original.sidecarSessionId,
          }),
        )
      const state = initialState(opened.metadata),
        sheetId = state.sheets[0]!.id
      const ranged = await run(state, [
        {
          op: 'set_protected_ranges',
          sheetId,
          ranges: [{ name: 'Inputs', sqref: 'B2:C3' }],
        },
      ])
      const locked = await run(ranged.nextState, [{ op: 'protect_workbook', locked: true }])
      expect(locked.nextState.features?.workbookProtectionEdits).toEqual({ lockStructure: true })
      await expect(run(locked.nextState, [{ op: 'add_sheet', name: 'Blocked' }])).rejects.toThrow(
        'protected_workbook_structure',
      )
      const staged = await new SessionPackageWriter(client, opened.metadata).stage(
        opened.original.path,
        locked.nextState,
        join(directory, 'stage.xlsx'),
      )
      const zip = await JSZip.loadAsync(await readFile(staged.stagedPath))
      expect(await zip.file('xl/workbook.xml')!.async('string')).toMatch(
        /<workbookProtection\b[^>]*lockStructure="1"/,
      )
      expect(await zip.file('xl/worksheets/sheet1.xml')!.async('string')).toMatch(
        /<protectedRange\b[^>]*name="Inputs"[^>]*sqref="B2:C3"|<protectedRange\b[^>]*sqref="B2:C3"[^>]*name="Inputs"/,
      )
      const reopened = await source.openSnapshot(staged.stagedPath)
      try {
        expect(
          (reopened.metadata as { workbookProtection?: { lockStructure?: boolean } })
            .workbookProtection?.lockStructure,
        ).toBe(true)
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
