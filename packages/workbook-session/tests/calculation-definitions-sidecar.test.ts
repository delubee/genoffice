import { afterAll, beforeAll, expect, it } from 'vitest'
import { build } from 'esbuild'
import { createRequire } from 'node:module'
import { mkdtemp, rm, writeFile, readFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import JSZip from 'jszip'
import type { IWorkbookData } from '@univerjs/core'
import type * as Calculation from '../src/calculation/univer'
import type { CalculationSourceData } from '../src/calculation/source'
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import { manualWorkbook } from '../../../apps/vscode-office/tests/fixtures/manual-workbook'
import { XlsxSidecarClient } from '../src/sidecar/client'
import { SidecarWorkbookSource } from '../src/sidecar/source'
import { openWorkbookSnapshot } from '../src/open'
import { initialState } from '../src/state'
import { prepareCandidateMaterial } from '../src/candidate-material'
import { SessionPackageWriter } from '../src/package-writer'
import { loadCalculationSource } from '../src/calculation/source'
import { parseAddress } from '@genoffice/xlsx-gateway/domain/cell-address'
import { useSessionTestResources } from './helpers'
const resources = useSessionTestResources()
let directory: string, calculate: typeof Calculation.calculateUniverWorkbook
beforeAll(async () => {
  directory = await mkdtemp(join(tmpdir(), 'genoffice-definition-runtime-'))
  const outfile = join(directory, 'calculation.cjs')
  await build({
    entryPoints: [fileURLToPath(new URL('../src/calculation/univer.ts', import.meta.url))],
    outfile,
    bundle: true,
    platform: 'node',
    format: 'cjs',
    target: 'node22',
    logLevel: 'silent',
  })
  calculate = (createRequire(import.meta.url)(outfile) as typeof Calculation)
    .calculateUniverWorkbook
})
afterAll(async () => {
  if (directory) await rm(directory, { recursive: true, force: true })
})

// Only the numeric/string/name fixture is adapted here; this is not the pending
// production snapshot/resource builder for tables, arrays, dates or external links.
function snapshot(loaded: CalculationSourceData): Partial<IWorkbookData> {
  return {
    id: loaded.stateId,
    sheetOrder: loaded.sheets.map((sheet) => sheet.id),
    sheets: Object.fromEntries(
      loaded.sheets.map((sheet) => {
        const cellData: Record<
          number,
          Record<number, { f?: string; v?: string | number | boolean }>
        > = {}
        for (const [address, cell] of Object.entries(sheet.cells)) {
          const { row, column } = parseAddress(address)
          cellData[row] ??= {}
          cellData[row]![column] = {
            ...(cell.rawValue === null ? {} : { v: cell.rawValue }),
            ...(cell.formula ? { f: cell.formula } : {}),
          }
        }
        return [
          sheet.id,
          {
            id: sheet.id,
            name: sheet.name,
            rowCount: sheet.rows,
            columnCount: sheet.columns,
            cellData,
          },
        ]
      }),
    ) as IWorkbookData['sheets'],
  }
}

it('calculates edited source/journal/hidden name dependencies consistently before and after real save/reopen', async () => {
  const work = await resources.directory(),
    path = join(work, 'definition-history.xlsx')
  const zip = await JSZip.loadAsync(await manualWorkbook())
  const workbook = await zip.file('xl/workbook.xml')!.async('string')
  zip.file(
    'xl/workbook.xml',
    workbook.replace(/<definedNames\b[^>]*(?:\/>|>[\s\S]*?<\/definedNames>)/g, '').replace(
      '</sheets>',
      `</sheets><definedNames>
    <definedName name="Revenue">Data!$A$2</definedName>
    <definedName name="Dependent">Revenue*2</definedName>
    <definedName name="HiddenDependent" hidden="1">Revenue*3</definedName>
    <definedName name="HiddenText" hidden="1"><![CDATA[IF(Revenue>0,"&amp; Revenue",0)]]></definedName>
    <definedName name="Revenue" hidden="1" localSheetId="1">Other!$A$1</definedName>
    <definedName name="LocalDependent" localSheetId="1">Revenue*2</definedName>
    </definedNames>`,
    ),
  )
  const dataXml = await zip.file('xl/worksheets/sheet1.xml')!.async('string')
  zip.file(
    'xl/worksheets/sheet1.xml',
    dataXml.replace(
      '</sheetData>',
      '<row r="100001"><c r="C100001"><f>Dependent+HiddenDependent+After</f><v>999</v></c><c r="D100001"><f>HiddenText</f><v>999</v></c></row></sheetData>',
    ),
  )
  const otherXml = await zip.file('xl/worksheets/sheet2.xml')!.async('string')
  zip.file(
    'xl/worksheets/sheet2.xml',
    otherXml.replace(
      '</sheetData>',
      '<row r="100001"><c r="C100001"><f>LocalDependent+Data!Revenue</f><v>999</v></c></row></sheetData>',
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
      storageDirectory: work,
      source,
    })
    try {
      let state = initialState(opened.metadata)
      const step = async (operations: WorkbookOperation[]) => {
        const before = structuredClone(state)
        const material = await prepareCandidateMaterial({
          state,
          source,
          sourceSessionId: opened.original.sidecarSessionId,
          operations,
        })
        const result = await computeCandidateState(
          state,
          {
            schemaVersion: 1,
            sessionId: 'names',
            transactionId: crypto.randomUUID(),
            expectedRevision: 0,
            origin: 'manual',
            summary: 'Manage calculation definitions',
            operations,
          },
          material,
        )
        expect(result.inverse.state).toEqual(before)
        state = result.nextState
      }
      await step([
        {
          op: 'manage_defined_name',
          name: 'Revenue',
          replacement: { name: 'Sales', ref: 'Data!$A$2' },
        },
      ])
      await step([
        { op: 'add_defined_name', name: 'Revenue', ref: '1' },
        { op: 'add_defined_name', name: 'After', ref: 'Revenue*5' },
      ])
      await step([
        {
          op: 'manage_defined_name',
          name: 'Sales',
          replacement: { name: 'Net', ref: 'Data!$A$2' },
        },
      ])
      const before = structuredClone(state)
      const loaded = await loadCalculationSource({
        state,
        source,
        sourceSessionId: opened.original.sidecarSessionId,
        isCurrent: () => true,
      })
      expect(loaded.definedNames.find((name) => name.name === 'Dependent')!.formula).toBe('Net*2')
      expect(loaded.definedNames.find((name) => name.name === 'HiddenDependent')!.formula).toBe(
        'Net*3',
      )
      expect(loaded.definedNames.find((name) => name.name === 'After')!.formula).toBe('Revenue*5')
      expect(loaded.definedNames.find((name) => name.name === 'LocalDependent')!.sheetId).toBe(
        state.sheets[1]!.id,
      )
      const evaluated = await calculate(snapshot(loaded), { definedNames: loaded.definedNames })
      const far = evaluated.formulas
        .filter((cell) => cell.row === 100000)
        .map((cell) => cell.values[0]![0]!.value)
      expect(far).toEqual([67.5, '&amp; Revenue', 26.5])
      expect(state).toEqual(before)
      expect(JSON.stringify(state)).not.toContain('Revenue*3')
      const staged = await new SessionPackageWriter(client, opened.metadata).stage(
        opened.original.path,
        state,
        join(work, 'saved.xlsx'),
      )
      const saved = await JSZip.loadAsync(await readFile(staged.stagedPath))
      // Export only this synthetic fixture for comparison with desktop Excel.
      if (process.env.GENOFFICE_CALC_ACCEPTANCE_ARTIFACT)
        await writeFile(
          process.env.GENOFFICE_CALC_ACCEPTANCE_ARTIFACT,
          await readFile(staged.stagedPath),
        )
      expect(await saved.file('customXml/manual-sentinel.xml')!.async('string')).toBe(
        '<preserve id="manual-fixture">unchanged</preserve>',
      )
      const reopened = await source.openSnapshot(staged.stagedPath)
      try {
        const root = initialState(reopened.metadata)
        const reloaded = await loadCalculationSource({
          state: root,
          source,
          sourceSessionId: reopened.sessionId,
          isCurrent: () => true,
        })
        const recalculated = await calculate(snapshot(reloaded), {
          definedNames: reloaded.definedNames,
        })
        expect(
          recalculated.formulas
            .filter((cell) => cell.row === 100000)
            .map((cell) => cell.values[0]![0]!.value),
        ).toEqual(far)
        expect(reloaded.definedNames.find((name) => name.name === 'HiddenDependent')!.formula).toBe(
          'Net*3',
        )
        expect(reloaded.definedNames.find((name) => name.name === 'HiddenText')!.formula).toBe(
          'IF(Net>0,"&amp; Revenue",0)',
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
