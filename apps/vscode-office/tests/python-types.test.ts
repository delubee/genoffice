import { expect, it } from 'vitest'
import { execFile } from 'node:child_process'
import { promisify } from 'node:util'
import { copyFile, mkdir, readFile, writeFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import { join } from 'node:path'
import JSZip from 'jszip'
import { typesWorkbook, preservationWorkbook } from './fixtures/sales-workbook'
import { nativeSession } from './fixtures/native-session'
import { withTestWorkspace } from './helpers'
import { readTypedNdjson } from '@genoffice/workbook-session/exchange/ndjson'
import { initialState } from '@genoffice/workbook-session/state'
import { readCurrentRange } from '@genoffice/workbook-session/read'

it('round-trips 1900/1904 dates, serial 60, literal errors/text and explicit formulas through Python, import, save and Rust reopening', async () => {
  for (const date1904 of [false, true])
    await withTestWorkspace(async (directory) => {
      const path = join(directory, 'types.xlsx'),
        storage = join(directory, 'private')
      await writeFile(path, await typesWorkbook(date1904))
      const runtime = await nativeSession(path, storage),
        { engine, snapshots, imports } = runtime
      try {
        await runtime.waitForIndex()
        const originalId = engine.state.stateId,
          sheetId = engine.state.sheets[0]!.id
        const snapshot = await snapshots.export({
          view: runtime.view(),
          sheetId,
          range: 'A2:L2',
          workspaceRoot: directory,
        })
        const cells = []
        for await (const record of readTypedNdjson({
          path: snapshot.dataPath,
          shape: snapshot.manifest.shape,
          expected: snapshot.manifest.data,
        }))
          cells.push(...record.cells)
        expect(cells[0]).toMatchObject({ kind: 'string', value: '00123' })
        expect(cells[1]).toMatchObject({ kind: 'string', value: '' })
        expect(cells[2]).toMatchObject({ kind: 'blank', value: null })
        expect(cells[3]).toMatchObject({
          kind: 'date',
          serial: 60,
          dateSystem: date1904 ? '1904' : '1900',
        })
        expect(cells[4]).toMatchObject({ kind: 'date', serial: 60.5 })
        expect(cells[5]).toMatchObject({ kind: 'number', value: 0.25, numberFormat: '0%' })
        expect(cells[7]).toMatchObject({ kind: 'string', value: '=1+2' })
        expect(cells[8]).toMatchObject({ kind: 'error', excelCode: '#REF!' })
        expect(cells[9]).toMatchObject({
          kind: 'formula',
          formula: '=1+2',
          calculationState: 'ready',
          cached: { kind: 'number', value: 3 },
        })
        expect(cells[10]).toMatchObject({
          kind: 'formula',
          formula: '=NOW()',
          calculationState: 'pending',
          cached: null,
        })
        expect(cells[11]).toMatchObject({
          kind: 'formula',
          formula: '=1/0',
          calculationState: 'error',
          cached: { kind: 'error', excelCode: '#DIV/0!' },
        })
        const script = join(snapshot.directory, 'analysis.py')
        await copyFile(
          fileURLToPath(
            new URL('../skills/genoffice-workbook/scripts/analyze.py', import.meta.url),
          ),
          script,
        )
        await promisify(execFile)(
          'python3',
          [
            script,
            snapshot.manifestPath,
            '--mode',
            'copy',
            '--new-sheet-name',
            'Roundtrip',
            '--allow-formulas',
          ],
          { cwd: snapshot.directory, timeout: 30_000 },
        )
        const manifestPath = join(snapshot.directory, 'result.manifest.json'),
          manifest = JSON.parse(await readFile(manifestPath, 'utf8'))
        const prepare = () =>
          imports.prepare({
            snapshotId: snapshot.manifest.snapshotId,
            bindingId: 'types',
            resultManifestPath: manifestPath,
            privateStorage: storage,
            view: runtime.view(),
            expectedRevision: 0,
            currentDescriptor: () => engine.descriptor,
          })
        await writeFile(manifestPath, JSON.stringify({ ...manifest, mode: 'values' }))
        await expect(prepare()).rejects.toMatchObject({ code: 'type_mismatch' })
        expect(engine.state.stateId).toBe(originalId)
        await writeFile(manifestPath, JSON.stringify(manifest))
        const plan = await prepare()
        const applied = await imports.apply({
          planId: plan.planId,
          bindingId: 'types',
          cursor: 0,
          expectedRevision: 0,
          transactionId: 'types-roundtrip',
          engine,
        })
        expect(applied.status).toBe('completed')
        const roundtrip = engine.state.sheets.find((sheet) => sheet.name === 'Roundtrip')!
        expect(roundtrip.cells.I1).toMatchObject({ kind: 'error', rawValue: '#REF!' })
        expect(roundtrip.cells.H1).toMatchObject({ kind: 'string', rawValue: '=1+2' })
        expect(roundtrip.cells.J1).toMatchObject({ formula: '=1+2', calculationState: 'pending' })
        await runtime.save()
        const reopened = await runtime.source.openSnapshot(path)
        try {
          const state = initialState(reopened.metadata),
            target = state.sheets.find((sheet) => sheet.name === 'Roundtrip')!
          await runtime.waitForIndex(
            reopened.sessionId,
            state.sheets.flatMap((sheet) => (sheet.originalSheetId ? [sheet.originalSheetId] : [])),
          )
          const saved = await readCurrentRange({
            state,
            logicalSheetId: target.id,
            range: 'A1:L1',
            source: runtime.source,
            sessionId: reopened.sessionId,
          })
          expect(saved.cells[0]).toMatchObject({ kind: 'string', rawValue: '00123' })
          expect(saved.cells[1]).toMatchObject({ kind: 'string', rawValue: '' })
          expect(saved.cells[2]).toMatchObject({ kind: 'blank', rawValue: null })
          expect(saved.cells[3]).toMatchObject({
            kind: 'date',
            rawValue: 60,
            dateSystem: date1904 ? '1904' : '1900',
          })
          expect(saved.cells[4]).toMatchObject({ kind: 'date', rawValue: 60.5 })
          expect(saved.cells[5]).toMatchObject({
            kind: 'number',
            rawValue: 0.25,
            numberFormat: '0%',
          })
          expect(saved.cells[8]).toMatchObject({ kind: 'error', rawValue: '#REF!' })
          for (const index of [9, 10, 11])
            expect(saved.cells[index]).toMatchObject({
              formula: (cells[index] as { formula: string }).formula,
              calculationState: 'pending',
            })
        } finally {
          await runtime.source.close(reopened.sessionId)
        }
      } finally {
        await runtime.dispose()
      }
    })
}, 120_000)

it('preserves original protected/merged/chart/unknown OOXML parts when editing a different worksheet', async () => {
  await withTestWorkspace(async (directory) => {
    const path = join(directory, 'preservation.xlsx'),
      original = await preservationWorkbook()
    await writeFile(path, original)
    const runtime = await nativeSession(path, join(directory, 'private')),
      { engine } = runtime
    try {
      await runtime.waitForIndex()
      const editable = engine.state.sheets.find((sheet) => sheet.name === 'Editable')!,
        protectedSheet = engine.state.sheets.find((sheet) => sheet.name === 'Preserve')!
      const features = (await runtime.source.readFeatures(
        runtime.opened.original.sidecarSessionId,
        protectedSheet.originalSheetId!,
      )) as unknown as {
        sheetProtection: { protected: boolean }
        nativeVisuals: unknown[]
        merges: unknown[]
      }
      expect(features.sheetProtection.protected).toBe(true)
      expect(features.nativeVisuals.length).toBeGreaterThan(0)
      expect(features.merges.length).toBeGreaterThan(0)
      await engine.submit({
        schemaVersion: 1,
        sessionId: engine.descriptor.sessionId,
        transactionId: 'editable-only',
        expectedRevision: 0,
        origin: 'manual',
        summary: 'Edit the unprotected sheet',
        operations: [{ op: 'set_cell', sheetId: editable.id, address: 'A1', value: 9 }],
      })
      const current = engine.state
      await expect(
        engine.submit({
          schemaVersion: 1,
          sessionId: engine.descriptor.sessionId,
          transactionId: 'protected-attempt',
          expectedRevision: 1,
          origin: 'agent',
          summary: 'Must reject protected write',
          operations: [
            { op: 'set_cell', sheetId: protectedSheet.id, address: 'A2', value: 'changed' },
          ],
        }),
      ).rejects.toMatchObject({ code: 'transaction_rejected' })
      expect(engine.state).toBe(current)
      await runtime.save()
      const before = await JSZip.loadAsync(original),
        after = await JSZip.loadAsync(await readFile(path))
      for (const part of [
        'xl/worksheets/sheet1.xml',
        'xl/worksheets/_rels/sheet1.xml.rels',
        'xl/drawings/drawing1.xml',
        'xl/drawings/_rels/drawing1.xml.rels',
        'xl/charts/chart1.xml',
        'customXml/genoffice-sentinel.xml',
      ])
        expect(await after.file(part)!.async('nodebuffer')).toEqual(
          await before.file(part)!.async('nodebuffer'),
        )
      const report = fileURLToPath(
        new URL('../test-results/preservation-proof.json', import.meta.url),
      )
      await mkdir(join(report, '..'), { recursive: true })
      await writeFile(
        report,
        JSON.stringify(
          {
            passed: true,
            protectedWriteRejected: true,
            originalChartDetected: true,
            originalMergedProtectedSheetPreserved: true,
            originalPartsByteIdentical: true,
            checkedAt: new Date().toISOString(),
          },
          null,
          2,
        ),
      )
    } finally {
      await runtime.dispose()
    }
  })
}, 120_000)
