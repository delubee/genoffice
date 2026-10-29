import { expect, it } from 'vitest'
import { smartFilterRange } from '../src/filter-region'
import type { ViewCell } from '../src/grid-cell'
import { parseAddress, rangeCellCount } from '@genoffice/xlsx-gateway/domain/cell-address'
import type { GridRange } from '../src/ports'

function source(values: Record<string, unknown>, merges: GridRange[] = []) {
  const calls: GridRange[] = []
  return {
    calls,
    async read(range: GridRange) {
      calls.push(range)
      expect(rangeCellCount(range)).toBeLessThanOrEqual(2000)
      return {
        merges,
        cells: Object.entries(values)
          .filter(([address]) => {
            const cell = parseAddress(address)
            return (
              cell.row >= range.startRow &&
              cell.row <= range.endRow &&
              cell.column >= range.startColumn &&
              cell.column <= range.endColumn
            )
          })
          .map(
            ([address, rawValue]) =>
              ({
                address,
                rawValue,
                kind: typeof rawValue === 'number' ? 'number' : 'string',
              }) as ViewCell,
          ),
      }
    },
  }
}
it('uses desktop expansion over complete source halos outside the visible grid and stops at blank separators', async () => {
  const file = source({
    B50001: 'Header',
    C50001: 'Total',
    B50002: 'North',
    C50002: 10,
    B50003: 'South',
    C50003: 20,
    B50005: 'Separate',
  })
  const result = await smartFilterRange(
    { startRow: 50001, endRow: 50001, startColumn: 2, endColumn: 2 },
    file.read,
    () => true,
  )
  expect(result).toEqual({ startRow: 50000, endRow: 50002, startColumn: 1, endColumn: 2 })
  expect(file.calls.length).toBeGreaterThan(0)
})
it('expands header-only selections downward while keeping their width and preserves explicit multi-row ranges', async () => {
  const file = source({
    A1: 'Outside',
    B1: 'Header',
    C1: 'Total',
    B2: 'North',
    C2: 10,
    D2: 'Outside',
    B3: 'South',
    C3: 20,
  })
  expect(
    await smartFilterRange(
      { startRow: 0, endRow: 0, startColumn: 1, endColumn: 2 },
      file.read,
      () => true,
    ),
  ).toEqual({ startRow: 0, endRow: 2, startColumn: 1, endColumn: 2 })
  file.calls.length = 0
  const explicit = { startRow: 0, endRow: 1, startColumn: 1, endColumn: 2 }
  expect(await smartFilterRange(explicit, file.read, () => true)).toEqual(explicit)
  expect(file.calls).toHaveLength(0)
})
it('uses the desktop merge-anchor geometry without authoring source cells', async () => {
  const file = source({ B2: 'Merged', B3: 'South', C3: 20 }, [
    { startRow: 1, endRow: 1, startColumn: 1, endColumn: 2 },
  ])
  expect(
    await smartFilterRange(
      { startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 },
      file.read,
      () => true,
    ),
  ).toEqual({ startRow: 1, endRow: 2, startColumn: 1, endColumn: 2 })
})
it('rejects an overtaken read or a region beyond the current canonical filter budget', async () => {
  const file = source({ A1: 'Header' })
  let current = true
  await expect(
    smartFilterRange(
      { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
      async (range) => {
        const result = await file.read(range)
        current = false
        return result
      },
      () => current,
    ),
  ).rejects.toThrow('document changed')
  await expect(
    smartFilterRange(
      { startRow: 0, endRow: 0, startColumn: 0, endColumn: 2000 },
      file.read,
      () => true,
    ),
  ).rejects.toThrow('2000')
})
