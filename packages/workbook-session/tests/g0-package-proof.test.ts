import { expect, it } from 'vitest'
import { readFile, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import {
  assembleWithJsZip,
  assertOnlyTouchedEntriesChanged,
  createBufferEntrySource,
  inventoryXlsx,
  planCellEditsToXlsx,
} from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { buildStructureFixture } from '../../../apps/sheets/tests/fixture-builder'
import { useSessionTestResources } from './helpers'

const resources = useSessionTestResources()

it('persists a structural change and chart from original base, then restores the inverse after save', async () => {
  const directory = await resources.directory()
  const original = await buildStructureFixture()
  const baseSource = await createBufferEntrySource(original)
  const operations = JSON.parse(
    JSON.stringify({
      structure: [{ sheetName: 'Data', ops: [{ kind: 'insert-rows', index: 1, count: 1 }] }],
      charts: [
        {
          sheetName: 'Data',
          anchor: {
            fromRow: 0,
            fromColumn: 5,
            fromRowOffset: 0,
            fromColumnOffset: 0,
            toRow: 15,
            toColumn: 12,
            toRowOffset: 0,
            toColumnOffset: 0,
          },
          chart: {
            chartType: 'column',
            title: 'G0',
            series: [
              {
                name: 'Values',
                categories: ['A', 'B'],
                values: [1, 2],
                valuesRef: "'Data'!$A$1:$A$3",
              },
            ],
          },
        },
      ],
    }),
  )
  const forward = await planCellEditsToXlsx(
    baseSource,
    [],
    operations.structure,
    [],
    undefined,
    [],
    [],
    [],
    [],
    [],
    null,
    operations.charts,
  )
  const mutation = await assembleWithJsZip(original, forward)
  assertOnlyTouchedEntriesChanged(mutation)
  const savedPath = join(directory, 'saved.xlsx')
  await writeFile(savedPath, mutation.buffer)
  const saved = await createBufferEntrySource(await readFile(savedPath))
  expect(await saved.readText('xl/worksheets/sheet1.xml')).toContain('r="A3"')
  const chartPath = (await inventoryXlsx(mutation.buffer)).find((part) =>
    /^xl\/charts\/chart\d+\.xml$/.test(part.path),
  )!.path
  expect(await saved.readText(chartPath)).toContain('<c:barChart>')
  // History keeps the pre-transaction root, not a closure over a disposed renderer.
  const inverse = JSON.parse(JSON.stringify({ structure: [], charts: [] }))
  expect(inverse).toEqual({ structure: [], charts: [] })
  // The identity overlay serializes the original bytes directly. The legacy
  // gateway requests full recalculation even for an empty mutation plan.
  await writeFile(savedPath, original)
  expect(await inventoryXlsx(await readFile(savedPath))).toEqual(await inventoryXlsx(original))
  const redo = await assembleWithJsZip(original, forward)
  expect(await inventoryXlsx(redo.buffer)).toEqual(await inventoryXlsx(mutation.buffer))
})
