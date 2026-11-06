import { expect, it } from 'vitest'
import { writeFile, readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import JSZip from 'jszip'
import { manualWorkbook } from '../../../apps/vscode-office/tests/fixtures/manual-workbook'
import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import { XlsxSidecarClient } from '../src/sidecar/client'
import { SidecarWorkbookSource } from '../src/sidecar/source'
import { openWorkbookSnapshot } from '../src/open'
import { initialState } from '../src/state'
import { prepareCandidateMaterial } from '../src/candidate-material'
import { readCurrentRange } from '../src/read'
import { SessionPackageWriter } from '../src/package-writer'
import { useSessionTestResources } from './helpers'
const resources = useSessionTestResources()
it('replays scoped name changes for off-page original formulas while preserving later user formulas through save/reopen', async () => {
  const directory = await resources.directory(),
    path = join(directory, 'references.xlsx')
  const fixture = await applyCellEditsToXlsx(
    await manualWorkbook(),
    [
      {
        sheetName: 'Data',
        row: 100000,
        column: 2,
        writeValue: true,
        cell: { value: 999, formula: '=Revenue+SUM(Revenue,1)' },
      },
      {
        sheetName: 'Data',
        row: 100000,
        column: 3,
        writeValue: true,
        cell: { value: 'Revenue12.5', formula: '="Revenue"&Revenue' },
      },
      {
        sheetName: 'Other',
        row: 100000,
        column: 2,
        writeValue: true,
        cell: { value: 999, formula: '=Revenue+Data!Revenue' },
      },
      { sheetName: 'Private', row: 0, column: 0, writeValue: true, cell: { value: 9 } },
      {
        sheetName: 'Private',
        row: 100000,
        column: 2,
        writeValue: true,
        cell: { value: 999, formula: '=Revenue+Data!Revenue' },
      },
    ],
    [],
    [],
    {
      renames: [],
      additions: [{ name: 'Private' }],
      removals: [],
      order: ['Data', 'Other', 'Private'],
    },
  )
  const zip = await JSZip.loadAsync(fixture.buffer)
  const hidden = '<definedName name="HiddenPrivate" hidden="1">Data!$A$2</definedName>'
  const hiddenLocal =
    '<definedName name="Revenue" hidden="1" localSheetId="2">Private!$A$1</definedName>'
  const print = '<definedName name="_xlnm.Print_Area" localSheetId="0">Data!$A$1:$D$3</definedName>'
  const workbook = await zip.file('xl/workbook.xml')!.async('string')
  zip.file(
    'xl/workbook.xml',
    workbook.replace(
      '</sheets>',
      '</sheets><definedNames><definedName name="Revenue">Data!$A$2</definedName><definedName name="Revenue" localSheetId="1">Other!$A$1</definedName>' +
        hidden +
        hiddenLocal +
        print +
        '</definedNames>',
    ),
  )
  const foreign =
    '<extLst><ext uri="vendor"><f>Revenue+private</f><v>private-cache</v></ext></extLst>'
  const originalSheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  zip.file(
    'xl/worksheets/sheet1.xml',
    originalSheet
      .replace('<f>Revenue+SUM(Revenue,1)</f>', '<f><![CDATA[Revenue+SUM(Revenue,1)]]></f>')
      .replace('</c>', foreign + '</c>')
      .replace('</worksheet>', foreign + '</worksheet>'),
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
      let state = structuredClone(initialState(opened.metadata))
      expect(opened.metadata).toMatchObject({
        nameBindingsComplete: true,
        reservedDefinedNames: expect.arrayContaining([
          expect.objectContaining({ name: 'HiddenPrivate' }),
          expect.objectContaining({ name: 'Revenue', sheetIndex: 2 }),
        ]),
      })
      const data = state.sheets[0]!.id,
        other = state.sheets[1]!.id
      const step = async (operations: WorkbookOperation[]) => {
        const before = structuredClone(state)
        const material = await prepareCandidateMaterial({
          state,
          operations,
          source,
          sourceSessionId: opened.original.sidecarSessionId,
        })
        if (operations.every((op) => op.op === 'manage_defined_name'))
          expect(material.referenceRanges).toEqual({})
        const result = await computeCandidateState(
          state,
          {
            schemaVersion: 1,
            sessionId: 'references',
            transactionId: crypto.randomUUID(),
            expectedRevision: 0,
            origin: 'manual',
            summary: 'Manage names',
            operations,
          },
          material,
        )
        expect(result.inverse.state).toEqual(before)
        state = result.nextState
      }
      await step([{ op: 'set_formula', sheetId: data, address: 'C5', formula: '=Revenue+1' }])
      await step([
        {
          op: 'manage_defined_name',
          name: 'Revenue',
          replacement: { name: 'Sales', ref: 'Data!$A$2' },
        },
      ])
      await step([{ op: 'set_formula', sheetId: data, address: 'D5', formula: '=Revenue+2' }])
      await step([{ op: 'add_defined_name', name: 'Revenue', ref: '1' }])
      await step([
        { op: 'manage_defined_name', name: 'Revenue', scopeSheetId: other, replacement: null },
      ])
      await step([
        {
          op: 'manage_defined_name',
          name: 'Sales',
          replacement: { name: 'Net', ref: 'Data!$A$2' },
        },
      ])
      expect(Object.keys(state.sheets[0]!.cells).sort()).toEqual(['C5', 'D5'])
      expect(state.sheets[0]!.cells.C5!.formula).toBe('=Net+1')
      expect(state.sheets[0]!.cells.D5!.formula).toBe('=Revenue+2')
      const current = await readCurrentRange({
        state,
        source,
        sessionId: opened.original.sidecarSessionId,
        logicalSheetId: data,
        range: 'C100001:C100001',
      })
      expect(current.cells[0]).toMatchObject({
        formula: '=Net+SUM(Net,1)',
        calculationState: 'pending',
      })
      const local = await readCurrentRange({
        state,
        source,
        sessionId: opened.original.sidecarSessionId,
        logicalSheetId: other,
        range: 'C100001:C100001',
      })
      expect(local.cells[0]!.formula).toBe('=#REF!+Data!Net')
      const protectedView = await readCurrentRange({
        state,
        source,
        sessionId: opened.original.sidecarSessionId,
        logicalSheetId: state.sheets[2]!.id,
        range: 'C100001:C100001',
      })
      expect(protectedView.cells[0]!.formula).toBe('=Revenue+Data!Net')
      const writer = new SessionPackageWriter(client, opened.metadata)
      const saved = await writer.stage(opened.original.path, state, join(directory, 'saved.xlsx'))
      const result = await JSZip.loadAsync(await readFile(saved.stagedPath))
      const dataXml = await result.file('xl/worksheets/sheet1.xml')!.async('string')
      expect(dataXml).toContain('<f>Net+SUM(Net,1)</f>')
      expect(dataXml.split(foreign)).toHaveLength(3)
      expect(/<c\b[^>]*>[\s\S]*?<\/c>/.exec(dataXml)![0]).toBe(
        /<c\b[^>]*>[\s\S]*?<\/c>/.exec(originalSheet)![0].replace('</c>', foreign + '</c>'),
      )
      expect(dataXml).toContain('<f>Net+1</f>')
      expect(dataXml).toContain('<f>Revenue+2</f>')
      expect(dataXml).toContain('&amp;Net</f>')
      expect(await result.file('xl/worksheets/sheet2.xml')!.async('string')).toContain(
        '<f>#REF!+Data!Net</f>',
      )
      const names = await result.file('xl/workbook.xml')!.async('string')
      expect(names).toContain(hidden)
      expect(names).toContain(hiddenLocal)
      expect(names).toContain(print)
      expect(await result.file('customXml/manual-sentinel.xml')!.async('string')).toBe(
        '<preserve id="manual-fixture">unchanged</preserve>',
      )
      const reopened = await source.openSnapshot(saved.stagedPath)
      try {
        const root = initialState(reopened.metadata)
        const view = await readCurrentRange({
          state: root,
          source,
          sessionId: reopened.sessionId,
          logicalSheetId: root.sheets[0]!.id,
          range: 'C100001:D100001',
        })
        expect(view.cells.map((cell) => cell.formula)).toEqual([
          '=Net+SUM(Net,1)',
          '="Revenue"&Net',
        ])
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
