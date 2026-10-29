import { expect, it } from 'vitest'
import { symbolIntents } from '../src/symbol-intent'

it('appends to raw text, numbers, booleans and blanks at only the chosen anchor', () => {
  const target = (
    rawValue: string | number | boolean | null,
    kind: 'string' | 'number' | 'boolean' | 'blank',
  ) => symbolIntents('sheet', 'B2', { address: 'B2', kind, rawValue }, 'Ω')
  expect(target('00123', 'string')).toEqual([
    { op: 'set_cell', sheetId: 'sheet', address: 'B2', value: '00123Ω' },
  ])
  expect(target(12.5, 'number')[0]).toMatchObject({ value: '12.5Ω' })
  expect(target(true, 'boolean')[0]).toMatchObject({ value: 'trueΩ' })
  expect(target(null, 'blank')[0]).toMatchObject({ value: 'Ω' })
})

it('uses the desktop facade conversion for formula-looking symbol output and parsed number formats', () => {
  expect(
    symbolIntents('sheet', 'A1', { address: 'A1', kind: 'string', rawValue: '=1+' }, 'π'),
  ).toEqual([{ op: 'set_formula', sheetId: 'sheet', address: 'A1', formula: '=1+π' }])
  // The facade recognizes 12.5€ as a currency. The accompanying number
  // format must share the value transaction rather than disappear on save.
  expect(
    symbolIntents('sheet', 'A1', { address: 'A1', kind: 'number', rawValue: 12.5 }, '€'),
  ).toEqual([
    { op: 'set_cell', sheetId: 'sheet', address: 'A1', value: 12.5 },
    { op: 'format_range', sheetId: 'sheet', range: 'A1', format: { numberFormat: '#,##0.00€' } },
  ])
  expect(
    symbolIntents(
      'sheet',
      'A1',
      { address: 'A1', kind: 'string', rawValue: '0'.repeat(32767) },
      '€',
    )[0],
  ).toMatchObject({ op: 'set_cell', value: 0 })
})

it('uses ready/error cached results and rejects pending results or oversized output before producing writes', () => {
  const formula = {
    address: 'A1',
    kind: 'number' as const,
    rawValue: 25,
    formula: '=A2*2',
    cachedValue: 30,
  }
  expect(
    symbolIntents('sheet', 'A1', { ...formula, calculationState: 'ready' }, 'Ω')[0],
  ).toMatchObject({ op: 'set_cell', value: '30Ω' })
  expect(
    symbolIntents(
      'sheet',
      'A1',
      { ...formula, cachedValue: null, calculationState: 'ready' },
      'Ω',
    )[0],
  ).toMatchObject({ value: 'Ω' })
  expect(
    symbolIntents(
      'sheet',
      'A1',
      { ...formula, cachedValue: '#DIV/0!', calculationState: 'error' },
      'Ω',
    )[0],
  ).toMatchObject({ value: '#DIV/0!Ω' })
  expect(() =>
    symbolIntents('sheet', 'A1', { ...formula, calculationState: 'pending' }, 'Ω'),
  ).toThrow('formula result is pending')
  expect(() =>
    symbolIntents(
      'sheet',
      'A1',
      { address: 'A1', kind: 'string', rawValue: 'x'.repeat(32767) },
      'Ω',
    ),
  ).toThrow('text limit')
})
