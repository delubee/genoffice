import { afterAll, beforeAll, expect, it } from 'vitest'
import { build } from 'esbuild'
import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createRequire } from 'node:module'
import { fileURLToPath } from 'node:url'
import type { IWorkbookData } from '@univerjs/core'
import type * as Calculation from '../src/calculation/univer'

let directory: string
let calculate: typeof Calculation.calculateUniverWorkbook
beforeAll(async () => {
  directory = await mkdtemp(join(tmpdir(), 'genoffice-calculation-'))
  const outfile = join(directory, 'calculation.cjs')
  // Exercise the same bundled Node/CommonJS boundary as the VS Code host,
  // including its transitive ESM/CommonJS dependencies; no browser or DOM.
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
function workbook(
  cells: Record<number, Record<number, { v?: number | string; f?: string }>>,
): Partial<IWorkbookData> {
  return {
    id: 'calculation-test',
    sheetOrder: ['data', 'other'],
    sheets: {
      data: { id: 'data', name: 'Data', rowCount: 2000, columnCount: 10, cellData: cells },
      other: {
        id: 'other',
        name: 'Other',
        rowCount: 20,
        columnCount: 10,
        cellData: { 0: { 0: { v: 7 } }, 1: { 0: { v: 8 } } },
      },
    } as IWorkbookData['sheets'],
  }
}
it('computes exact numeric, boolean, error and error-shaped text results without a DOM or input mutation', async () => {
  expect('document' in globalThis).toBe(false)
  const input = workbook({
    0: {
      0: { v: 2 },
      1: { f: '=A1*2' },
      2: { f: '=1=1' },
      3: { f: '=1/0' },
      4: { f: '="#DIV/0!"' },
      5: { f: '=Other!A1+A1' },
    },
  })
  const before = structuredClone(input)
  const result = await calculate(input)
  expect(result.scope).toEqual({ kind: 'workbook' })
  expect(result.workbookId).toBe(input.id)
  const values = new Map(result.formulas.map((cell) => [cell.column, cell.values]))
  expect(values.get(1)).toEqual([[{ type: 'number', value: 4 }]])
  expect(values.get(2)).toEqual([[{ type: 'boolean', value: true }]])
  expect(values.get(3)).toEqual([[{ type: 'error', value: '#DIV/0!' }]])
  expect(values.get(4)).toEqual([[{ type: 'string', value: '#DIV/0!' }]])
  expect(values.get(5)).toEqual([[{ type: 'number', value: 9 }]])
  expect(input).toEqual(before)
})
it('captures the full dynamic-array matrix and reference spills instead of only worksheet anchor caches', async () => {
  const result = await calculate(
    workbook({ 0: { 0: { f: '=SEQUENCE(2,2)' } }, 4: { 2: { f: '=Other!A1:A2' } } }),
  )
  expect(
    result.formulas
      .find((cell) => cell.row === 0)!
      .values.map((row) => row.map((cell) => cell.value)),
  ).toEqual([
    [1, 2],
    [3, 4],
  ])
  expect(result.formulas.find((cell) => cell.row === 4)!.values).toEqual([
    [{ type: 'number', value: 7 }],
    [{ type: 'number', value: 8 }],
  ])
})
it('reports a blocked spill as an error without replacing the blocking literal or publishing spill children', async () => {
  const input = workbook({ 0: { 0: { f: '=SEQUENCE(2,2)' } }, 1: { 1: { v: 99 } } }),
    before = structuredClone(input)
  const result = await calculate(input)
  expect(result.formulas[0]!.values).toEqual([[{ type: 'error', value: '#SPILL!' }]])
  expect(input).toEqual(before)
})
it('keeps an accepted array whose first value merely spells SPILL as text', async () => {
  const result = await calculate(workbook({ 0: { 0: { f: '={"#SPILL!",2;3,4}' } } }))
  expect(result.formulas[0]!.values).toEqual([
    [
      { type: 'string', value: '#SPILL!' },
      { type: 'number', value: 2 },
    ],
    [
      { type: 'number', value: 3 },
      { type: 'number', value: 4 },
    ],
  ])
})
it('returns no partial results when the output budget is exceeded and permits a later independent calculation', async () => {
  await expect(
    calculate(workbook({ 0: { 0: { f: '=SEQUENCE(2,2)' } } }), { maxResultCells: 1 }),
  ).rejects.toMatchObject({ code: 'calculation_result_limit' })
  expect((await calculate(workbook({ 0: { 0: { f: '=2+3' } } }))).formulas[0]!.values).toEqual([
    [{ type: 'number', value: 5 }],
  ])
})
it('supports empty inputs and cancellation before calculation without creating results', async () => {
  const abort = new AbortController()
  abort.abort()
  await expect(calculate(workbook({}), { signal: abort.signal })).rejects.toMatchObject({
    code: 'calculation_cancelled',
  })
  expect((await calculate(workbook({}))).formulas).toEqual([])
})
it('stops an in-progress calculation, disposes it and leaves the next workbook independent', async () => {
  const abort = new AbortController()
  const input = workbook(
    Object.fromEntries(Array.from({ length: 1000 }, (_, row) => [row, { 0: { f: '=RAND()' } }])),
  )
  let observed = false
  await expect(
    calculate(input, {
      signal: abort.signal,
      onProgress(progress) {
        if (progress.completedFormulasCount > 0) {
          observed = true
          abort.abort()
        }
      },
    }),
  ).rejects.toMatchObject({ code: 'calculation_cancelled' })
  expect(observed).toBe(true)
  expect((await calculate(workbook({ 0: { 0: { f: '=4+5' } } }))).formulas[0]!.values).toEqual([
    [{ type: 'number', value: 9 }],
  ])
})

it('keeps per-element error provenance inside mixed array results', async () => {
  const result = await calculate(
    workbook({ 0: { 0: { f: '=IF(SEQUENCE(2,2)=1,1/0,"#DIV/0!")' } } }),
  )
  expect(result.formulas[0]!.values).toEqual([
    [
      { type: 'error', value: '#DIV/0!' },
      { type: 'string', value: '#DIV/0!' },
    ],
    [
      { type: 'string', value: '#DIV/0!' },
      { type: 'string', value: '#DIV/0!' },
    ],
  ])
})
it('isolates concurrent workbooks with the same caller IDs and different values', async () => {
  const make = (value: number) =>
    workbook({ 0: { 0: { v: value }, 1: { f: '=SUMIF(A1:A2,">0",A1:A2)' } } })
  const results = await Promise.all([calculate(make(2)), calculate(make(17)), calculate(make(31))])
  expect(results.map((result) => result.formulas[0]!.values)).toEqual(
    [2, 17, 31].map((value) => [[{ type: 'number', value }]]),
  )
})

it('retains cancellation requested from the first START progress notification across the engine reset', async () => {
  const abort = new AbortController()
  let progressCalls = 0
  await expect(
    calculate(workbook({ 0: { 0: { f: '=1+2' } } }), {
      signal: abort.signal,
      onProgress() {
        if (++progressCalls === 1) abort.abort()
      },
    }),
  ).rejects.toMatchObject({ code: 'calculation_cancelled' })
  expect(progressCalls).toBeGreaterThan(0)
})

it('retains computed date-format hints, hyperlinks and image metadata without rewriting the input styles', async () => {
  const input = workbook({
    0: {
      0: { f: '=DATE(2026,10,27)' },
      1: { f: '=HYPERLINK("https://example.invalid/", "label")' },
      2: { f: '=IMAGE("https://example.invalid/image.png","caption",3,40,60)' },
    },
  })
  const before = structuredClone(input)
  const result = await calculate(input)
  const date = result.formulas.find((cell) => cell.column === 0)!.values[0]![0]!
  expect(date.type).toBe('number')
  expect(date.value).toBe((Date.UTC(2026, 9, 27) - Date.UTC(1899, 11, 30)) / 86400000)
  expect(date.numberFormat).toMatch(/[ymd]/i)
  expect(result.formulas.find((cell) => cell.column === 1)!.values).toEqual([
    [{ type: 'string', value: 'label', hyperlink: 'https://example.invalid/' }],
  ])
  expect(result.formulas.find((cell) => cell.column === 2)!.values).toEqual([
    [
      {
        type: 'string',
        value: '',
        image: {
          source: 'https://example.invalid/image.png',
          altText: 'caption',
          sizing: 3,
          height: 40,
          width: 60,
        },
      },
    ],
  ])
  expect(input).toEqual(before)
})

it('uses desktop Calculate Sheet dirtiness, updates off-sheet dependents and keeps unrelated cached formulas outside the result', async () => {
  const input = workbook({ 0: { 0: { v: 2 }, 1: { f: '=A1*2', v: 0 } } })
  input.sheets!.other!.cellData = { 0: { 0: { f: '=Data!B1+1', v: 0 }, 2: { f: '=1+2', v: 777 } } }
  const before = structuredClone(input)
  const result = await calculate(input, { sheetId: 'data' })
  expect(result.scope).toEqual({ kind: 'sheet', sheetId: 'data' })
  expect(
    result.formulas.map((cell) => [cell.sheetId, cell.row, cell.column, cell.values[0]![0]!.value]),
  ).toEqual([
    ['data', 0, 1, 4],
    ['other', 0, 0, 5],
  ])
  expect(input).toEqual(before)
  await expect(calculate(input, { sheetId: 'missing' })).rejects.toMatchObject({
    code: 'calculation_unknown_sheet',
  })
})

it('matches engine blank-reference semantics for a single cell and a spilled range', async () => {
  const result = await calculate(
    workbook({ 0: { 0: { f: '=B10' } }, 2: { 2: { f: '=Other!B1:B2' } } }),
  )
  expect(result.formulas.find((cell) => cell.row === 0)!.values).toEqual([
    [{ type: 'number', value: 0 }],
  ])
  expect(result.formulas.find((cell) => cell.row === 2)!.values).toEqual([
    [{ type: 'blank', value: null }],
    [{ type: 'blank', value: null }],
  ])
})

it('binds identical global/local names independently across sheets and qualified references without altering authored formulas', async () => {
  const input = workbook({
    0: { 0: { f: '=Rate*3' }, 1: { f: '=Other!Rate*3' }, 2: { f: '="Rate"' }, 3: { f: '=Rate+1' } },
  })
  input.sheets!.other!.cellData = { 0: { 0: { f: '=Rate*3' }, 1: { f: '=Data!Rate*3' } } }
  const before = structuredClone(input)
  const result = await calculate(input, {
    definedNames: [
      { name: 'Rate', formula: '10' },
      { name: 'Rate', formula: '1+2', sheetId: 'data', hidden: true },
      { name: 'Rate', formula: '4+5', sheetId: 'other' },
    ],
  })
  expect(input).toEqual(before)
  expect(
    result.formulas.map((c) => [c.sheetId, c.column, c.formula, c.values[0]![0]!.value]),
  ).toEqual([
    ['data', 0, '=Rate*3', 9],
    ['data', 1, '=Other!Rate*3', 27],
    ['data', 2, '="Rate"', 'Rate'],
    ['data', 3, '=Rate+1', 4],
    ['other', 0, '=Rate*3', 27],
    ['other', 1, '=Data!Rate*3', 9],
  ])
})

it('resolves nested and Unicode names with global fallback and respects LET/LAMBDA variables and built-in functions', async () => {
  const input = workbook({
    0: {
      0: { f: '=Rate*3' },
      1: { f: '=LET(Rate,2,Rate+Bonus)' },
      2: { f: '=LAMBDA(Rate,Rate+Bonus)(2)' },
      3: { f: '=SUM(1,2)+SUM' },
      4: { f: '=样本+rate' },
    },
  })
  input.sheets!.other!.cellData = { 0: { 0: { f: '=Rate*3' } } }
  const result = await calculate(input, {
    definedNames: [
      { name: 'Rate', formula: '10' },
      { name: 'Rate', formula: 'Bonus+1', sheetId: 'data' },
      { name: 'Bonus', formula: '3' },
      { name: 'SUM', formula: '100' },
      { name: '样本', formula: '2' },
    ],
  })
  expect(result.formulas.map((c) => c.values[0]![0]!.value)).toEqual([12, 5, 5, 103, 6, 30])
})

it('resolves dynamic INDIRECT name strings using the current sheet and case-insensitive qualified name scope', async () => {
  const input = workbook({
    0: { 0: { v: 2 }, 1: { f: '=INDIRECT("rate")' }, 2: { f: '=INDIRECT("Other!Rate")' } },
  })
  input.sheets!.other!.cellData = { 0: { 0: { v: 7 }, 1: { f: '=INDIRECT("Rate")' } } }
  const result = await calculate(input, {
    definedNames: [
      { name: 'Rate', formula: 'Data!$A$1', sheetId: 'data' },
      { name: 'Rate', formula: 'Other!$A$1', sheetId: 'other' },
    ],
  })
  expect(result.formulas.map((c) => c.values[0]![0]!.value)).toEqual([2, 7, 7])
})

it('rejects duplicate, foreign-scope and circular name bindings with no successful partial result', async () => {
  const input = workbook({ 0: { 0: { f: '=Rate' } } })
  await expect(
    calculate(input, {
      definedNames: [
        { name: 'Rate', formula: '1' },
        { name: 'rate', formula: '2' },
      ],
    }),
  ).rejects.toMatchObject({ code: 'calculation_ambiguous_name' })
  await expect(
    calculate(input, { definedNames: [{ name: 'Rate', formula: '1', sheetId: 'missing' }] }),
  ).rejects.toMatchObject({ code: 'calculation_invalid_name' })
  await expect(
    calculate(input, {
      definedNames: [
        { name: 'Rate', formula: 'OtherName' },
        { name: 'OtherName', formula: 'Rate' },
      ],
    }),
  ).rejects.toMatchObject({ code: 'calculation_circular_name' })
})

it('binds local range/array definitions to their own worksheet, quoted scopes and workbook-qualified global names', async () => {
  const input = workbook({
    0: {
      0: { v: 2 },
      1: { f: '=Window' },
      3: { f: "='O''Brien'!Window" },
      5: { f: "='Book.xlsx'!Rate" },
    },
    1: { 0: { v: 3 } },
  })
  input.name = 'Book.xlsx'
  input.sheets!.other!.name = "O'Brien"
  input.sheets!.other!.cellData = {
    0: { 0: { v: 7 }, 1: { f: '=SUM(Data!Window)' } },
    1: { 0: { v: 8 } },
  }
  const result = await calculate(input, {
    definedNames: [
      { name: 'Window', formula: '$A$1:$A$2', sheetId: 'data' },
      { name: 'Window', formula: '$A$1:$A$2', sheetId: 'other' },
      { name: 'Rate', formula: '10' },
      { name: 'Rate', formula: '2', sheetId: 'data' },
    ],
  })
  expect(result.formulas[0]!.values.map((row) => row[0]!.value)).toEqual([2, 3])
  expect(result.formulas[1]!.values.map((row) => row[0]!.value)).toEqual([7, 8])
  expect(result.formulas[2]!.values[0]![0]!.value).toBe(10)
  expect(result.formulas[3]!.values[0]![0]!.value).toBe(5)
})

it('evaluates named LAMBDA calls and sequential LET shadowing without rewriting string literals', async () => {
  const input = workbook({
    0: {
      0: { f: '=Bump(2)' },
      1: { f: '=LET( Rate,Rate+1,Rate)' },
      2: { f: '=LET(Rate,2,Bonus,Rate+1,Bonus+Rate)' },
      3: { f: '=LAMBDA(Bonus,Bonus+Rate)(2)' },
    },
  })
  const result = await calculate(input, {
    definedNames: [
      { name: 'Bump', formula: 'LAMBDA(x,x+Bonus)' },
      { name: 'Rate', formula: '10' },
      { name: 'Bonus', formula: '3' },
    ],
  })
  expect(result.formulas.map((c) => c.values[0]![0]!.value)).toEqual([5, 11, 5, 12])
})

it('isolates concurrent name bindings and freezes definitions before progress callbacks can mutate caller data', async () => {
  const input = workbook({ 0: { 0: { f: '=Rate+INDIRECT("Value")' }, 1: { v: 2 } } })
  const names = [
    { name: 'Rate', formula: '3' },
    { name: 'Value', formula: 'Data!$B$1' },
  ]
  const results = await Promise.all([
    calculate(input, {
      definedNames: names,
      onProgress() {
        names[1]!.formula = 'Data!$Z$99'
      },
    }),
    calculate(input, {
      definedNames: [
        { name: 'Rate', formula: '10' },
        { name: 'Value', formula: 'Data!$B$1' },
      ],
    }),
  ])
  expect(results.map((r) => r.formulas[0]!.values[0]![0]!.value)).toEqual([5, 12])
})

it('expands shared formulas before binding names and returns original logical formula text without private aliases', async () => {
  const input = workbook({ 0: { 0: { v: 2 } }, 1: { 0: { v: 3 } } })
  input.sheets!.data!.cellData = {
    0: { 0: { v: 2 }, 1: { f: '=Rate+A1', si: 'shared' } },
    1: { 0: { v: 3 }, 1: { si: 'shared' } },
  }
  const before = structuredClone(input)
  const result = await calculate(input, { definedNames: [{ name: 'Rate', formula: '1' }] })
  expect(result.formulas.map((c) => [c.formula, c.values[0]![0]!.value])).toEqual([
    ['=Rate+A1', 3],
    ['=Rate+A2', 4],
  ])
  expect(JSON.stringify(result)).not.toContain('_genoffice_')
  expect(input).toEqual(before)
})

it('preserves unary operators, percentages, array LET values and lexical variable scope while binding names', async () => {
  const result = await calculate(
    workbook({
      0: {
        0: { f: '=-Rate' },
        1: { f: '=+Rate' },
        2: { f: '=Rate%' },
        3: { f: '=SUM(LET(Rate,{1,2},Rate))+Bonus' },
        4: { f: '=LET(Rate,2,Rate)+Rate' },
        5: { f: '=IFERROR(1/0,Rate)' },
      },
    }),
    {
      definedNames: [
        { name: 'Rate', formula: '10' },
        { name: 'Bonus', formula: '3' },
      ],
    },
  )
  expect(result.formulas.map((c) => c.values[0]![0]!.value)).toEqual([-10, 10, 0.1, 6, 12, 10])
})

it('rejects missing and ambiguous shared masters before name binding and leaves the input intact', async () => {
  const input = workbook({})
  input.sheets!.data!.cellData = { 0: { 0: { si: 'missing' } } }
  const before = structuredClone(input)
  await expect(calculate(input)).rejects.toMatchObject({
    code: 'calculation_missing_shared_formula',
  })
  expect(input).toEqual(before)
  input.sheets!.data!.cellData = {
    0: { 0: { si: 'duplicate', f: '=1' }, 1: { si: 'duplicate', f: '=2' } },
  }
  await expect(calculate(input)).rejects.toMatchObject({
    code: 'calculation_ambiguous_shared_formula',
  })
})

it('matches Excel 16 global-definition lexical scope even when the consuming sheet has a local shadow', async () => {
  const input = workbook({ 0: { 0: { f: '=DoubleRate' }, 1: { f: '=OtherDouble' } } })
  input.sheets!.other!.cellData = {
    0: { 0: { f: '=DoubleRate' }, 1: { f: '=OtherDouble' }, 2: { f: '=Rate' } },
  }
  const result = await calculate(input, {
    definedNames: [
      { name: 'Rate', formula: '10' },
      { name: 'Rate', formula: '20', sheetId: 'other' },
      { name: 'DoubleRate', formula: 'Rate*2' },
      { name: 'OtherDouble', formula: 'Other!Rate*2' },
    ],
  })
  expect(result.formulas.map((c) => c.values[0]![0]!.value)).toEqual([20, 40, 20, 40, 20])
})
