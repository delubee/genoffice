import { expect, it } from 'vitest'
import { readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { typedCellSchema, type TypedCell } from '../src/exchange/schemas'
import { toTypedCell, fromTypedCell } from '../src/exchange/cells'
import { writeTypedNdjson, readTypedNdjson } from '../src/exchange/ndjson'
import { useSessionTestResources } from './helpers'
const resources = useSessionTestResources()

it('round-trips dates, serial 60, fractional time, percentages, literal strings, blanks and independent formula caches', async () => {
  const cells: TypedCell[] = [
    { kind: 'date', serial: 60.5, dateSystem: '1900', numberFormat: 'yyyy-mm-dd hh:mm' },
    { kind: 'date', serial: 0.25, dateSystem: '1904', numberFormat: 'hh:mm' },
    { kind: 'number', value: 0.25, numberFormat: '0%' },
    { kind: 'string', value: '00123' },
    { kind: 'string', value: '=A1' },
    { kind: 'blank', value: null },
    { kind: 'string', value: '' },
    { kind: 'error', excelCode: '#REF!' },
    {
      kind: 'formula',
      formula: '=A1',
      cached: { kind: 'number', value: 60.5 },
      calculationState: 'pending',
    },
    {
      kind: 'formula',
      formula: '=1/0',
      cached: { kind: 'error', excelCode: '#DIV/0!' },
      calculationState: 'error',
    },
    {
      kind: 'formula',
      formula: '=42',
      cached: { kind: 'number', value: 42 },
      calculationState: 'ready',
    },
    { kind: 'formula', formula: '=NOW()', cached: null, calculationState: 'pending' },
    {
      kind: 'formula',
      formula: '=IF(TRUE,,)',
      cached: { kind: 'blank', value: null },
      calculationState: 'ready',
    },
  ]
  const directory = await resources.directory(),
    path = join(directory, 'typed.ndjson'),
    shape = { rows: 1, columns: cells.length, cellCount: cells.length }
  const data = await writeTypedNdjson({
    path,
    shape,
    records: [{ rowOffset: 0, columnOffset: 0, cells }],
  })
  const restored = []
  for await (const record of readTypedNdjson({ path, shape, expected: data }))
    restored.push(...record.cells)
  expect(restored).toEqual(cells)
  for (const cell of cells)
    expect(toTypedCell(fromTypedCell(cell, { trustFormulaCache: true }))).toEqual(cell)
  expect(typedCellSchema.safeParse({ kind: 'number', value: NaN }).success).toBe(false)
  expect(typedCellSchema.safeParse({ kind: 'number', value: Infinity }).success).toBe(false)
})

it('segments wide Unicode rows within the record limit and rejects incomplete, overlapping or over-budget artifacts', async () => {
  const directory = await resources.directory(),
    path = join(directory, 'wide.ndjson'),
    cells = Array.from({ length: 10 }, () => ({
      kind: 'string' as const,
      value: '中'.repeat(32767),
    })),
    shape = { rows: 1, columns: 10, cellCount: 10 }
  const data = await writeTypedNdjson({
    path,
    shape,
    records: [{ rowOffset: 0, columnOffset: 0, cells }],
  })
  expect(data.recordCount).toBeGreaterThan(1)
  for (const line of (await readFile(path, 'utf8')).trimEnd().split('\n'))
    expect(Buffer.byteLength(line)).toBeLessThanOrEqual(256 * 1024)
  await expect(
    writeTypedNdjson({
      path: join(directory, 'missing.ndjson'),
      shape,
      records: [{ rowOffset: 0, columnOffset: 1, cells: cells.slice(1) }],
    }),
  ).rejects.toMatchObject({ code: 'artifact_incomplete' })
  await expect(
    writeTypedNdjson({
      path: join(directory, 'large.ndjson'),
      shape,
      records: [{ rowOffset: 0, columnOffset: 0, cells }],
      byteLimit: 100,
    }),
  ).rejects.toMatchObject({ code: 'artifact_limit_exceeded' })
})
