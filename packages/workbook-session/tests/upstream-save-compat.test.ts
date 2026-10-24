import { expect, it } from 'vitest'
import { readFile, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { blankXlsxBuffer } from '@genoffice/xlsx-gateway/gateway/csv-import'
import { applyCellEditsToXlsx } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { saveWorkbookViaSidecar } from '@genoffice/xlsx-gateway/gateway/xlsx-package-io'
import {
  REDACTION_PART_PATH,
  parseRedactionPart,
} from '@genoffice/xlsx-gateway/gateway/xlsx-redaction'
import { redactionStatesSchema } from '@genoffice/xlsx-gateway/shared/workbook-protocol'
import { XlsxSidecarClient } from '../src/sidecar/client'
import { useSessionTestResources } from './helpers'

const resources = useSessionTestResources()

it('streams upstream redaction marks and fork formula-cache invalidation in the same save', async () => {
  const directory = await resources.directory()
  const sourcePath = join(directory, 'source.xlsx'),
    targetPath = join(directory, 'saved.xlsx')
  const fixture = await applyCellEditsToXlsx(await blankXlsxBuffer('Data'), [
    {
      sheetName: 'Data',
      row: 0,
      column: 0,
      writeValue: true,
      cell: { formula: '=1+1', value: 999 },
    },
  ])
  await writeFile(sourcePath, fixture.buffer)
  const states = redactionStatesSchema.parse([
    {
      sheetName: 'Data',
      marks: [
        {
          startRow: 1,
          endRow: 1,
          startColumn: 1,
          endColumn: 1,
          label: 'Hidden data',
          previousFill: '#CCEEFF',
        },
      ],
    },
  ])
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
    await saveWorkbookViaSidecar({
      client,
      sourcePath,
      targetPath,
      edits: [],
      redactionStates: states,
      invalidateFormulaCaches: true,
    })
    const saved = await JSZip.loadAsync(await readFile(targetPath))
    expect(parseRedactionPart(await saved.file(REDACTION_PART_PATH)!.async('string'))).toEqual(
      states,
    )
    const sheet = await saved.file('xl/worksheets/sheet1.xml')!.async('string')
    expect(sheet).toContain('<f>1+1</f>')
    expect(sheet).not.toContain('<v>999</v>')
  } finally {
    client.stop()
  }
})
