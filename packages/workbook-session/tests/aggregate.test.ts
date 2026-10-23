import { expect, it } from 'vitest'
import { aggregateCurrentRange } from '../src/aggregate'
import { parseRange, formatAddress } from '@genoffice/xlsx-gateway/domain/cell-address'

it('aggregates typed raw values without writing cells or conflating blank, empty string and numeric text', async () => {
  const cells = [
    { address: 'A1', kind: 'number', rawValue: 10, displayValue: '$10' },
    { address: 'A2', kind: 'number', rawValue: 20, displayValue: '$20' },
    { address: 'A3', kind: 'string', rawValue: '00123', displayValue: '00123' },
    { address: 'A4', kind: 'blank', rawValue: null, displayValue: '' },
    { address: 'A5', kind: 'string', rawValue: '', displayValue: '' },
  ] as const
  const result = await aggregateCurrentRange({ range: 'A1:A5', read: async () => [...cells] })
  expect(result).toMatchObject({
    cellCount: 5,
    nonEmptyCount: 4,
    distinctCount: 4,
    numericCount: 2,
    sum: 30,
    mean: 15,
    min: 10,
    max: 20,
  })
  expect(result.topValues).toContainEqual({ kind: 'string', rawValue: '', count: 1 })
  expect(result.topValues).not.toContainEqual(expect.objectContaining({ kind: 'blank' }))
})

it('pages a complete 100k-cell current view and bounds the frequency result without helper cells', async () => {
  let calls = 0
  const result = await aggregateCurrentRange({
    range: 'A1:J10000',
    read: async (range) => {
      calls++
      const bounds = parseRange(range),
        cells = []
      for (let row = bounds.startRow; row <= bounds.endRow; row++)
        for (let column = bounds.startColumn; column <= bounds.endColumn; column++)
          cells.push({
            address: formatAddress(row, column),
            kind: 'number' as const,
            rawValue: row * 10 + column,
          })
      expect(cells.length).toBeLessThanOrEqual(2000)
      return cells
    },
  })
  expect(result).toMatchObject({
    cellCount: 100000,
    nonEmptyCount: 100000,
    numericCount: 100000,
    distinctCount: 100000,
    sum: (100000 * 99999) / 2,
    min: 0,
    max: 99999,
  })
  expect(result.topValues).toHaveLength(50)
  expect(calls).toBe(50)
})

it('rejects oversized ranges before reading and never treats a pending formula cache as an authoritative aggregate', async () => {
  let reads = 0
  const read = async () => {
    reads++
    return [
      {
        address: 'A1',
        kind: 'number' as const,
        rawValue: 25,
        formula: '=A2*2',
        calculationState: 'pending' as const,
        displayValue: '25',
      },
    ]
  }
  await expect(aggregateCurrentRange({ range: 'A1:B600000', read })).rejects.toMatchObject({
    code: 'range_limit',
  })
  expect(reads).toBe(0)
  await expect(aggregateCurrentRange({ range: 'A1', read })).rejects.toMatchObject({
    code: 'calculation_pending',
  })
})
