import {
  formatAddress,
  parseRange,
  rangeCellCount,
} from '@genoffice/xlsx-gateway/domain/cell-address'
import { MAX_GRID_COLUMNS, MAX_GRID_ROWS } from '@genoffice/xlsx-gateway/shared/grid-bounds'
import { LIMITS, WorkbookError, assertNotCancelled, type CellRecord } from './domain'

export async function aggregateCurrentRange(input: {
  range: string
  read(range: string): Promise<readonly (CellRecord & { address: string })[]>
  signal?: AbortSignal
}) {
  const bounds = parseRange(input.range),
    cellCount = rangeCellCount(bounds)
  if (
    cellCount > LIMITS.aggregateCells ||
    bounds.endRow >= MAX_GRID_ROWS ||
    bounds.endColumn >= MAX_GRID_COLUMNS
  )
    throw new WorkbookError(
      'range_limit',
      'Aggregate at most 1,000,000 cells inside the Excel grid.',
    )
  const frequencies = new Map<
    string,
    {
      kind: CellRecord['kind']
      rawValue: CellRecord['rawValue']
      count: number
      dateSystem?: CellRecord['dateSystem']
    }
  >()
  let nonEmptyCount = 0,
    numericCount = 0,
    sum = 0,
    correction = 0,
    min = Infinity,
    max = -Infinity
  // Bound every source page, including very wide rows.
  for (let row = bounds.startRow; row <= bounds.endRow;) {
    const width = bounds.endColumn - bounds.startColumn + 1
    const rowEnd = Math.min(
      bounds.endRow,
      row + Math.max(1, Math.floor(LIMITS.readCells / width)) - 1,
    )
    for (let column = bounds.startColumn; column <= bounds.endColumn; column += LIMITS.readCells) {
      assertNotCancelled(input.signal)
      const columnEnd = Math.min(bounds.endColumn, column + LIMITS.readCells - 1)
      const cells = await input.read(
        `${formatAddress(row, column)}:${formatAddress(rowEnd, columnEnd)}`,
      )
      if (cells.length !== (rowEnd - row + 1) * (columnEnd - column + 1))
        throw new WorkbookError('invalid_document', 'The current read page is incomplete.')
      for (const cell of cells) {
        if (cell.formula && cell.calculationState !== 'ready')
          throw new WorkbookError(
            'calculation_pending',
            'Formula caches are not authoritative; export and calculate explicitly before relying on their aggregates.',
            { address: cell.address },
            true,
          )
        if (cell.kind === 'blank') continue
        nonEmptyCount++
        const key = JSON.stringify([cell.kind, cell.rawValue, cell.dateSystem ?? null])
        const item = frequencies.get(key)
        if (item) item.count++
        else
          frequencies.set(key, {
            kind: cell.kind,
            rawValue: cell.rawValue,
            count: 1,
            ...(cell.dateSystem ? { dateSystem: cell.dateSystem } : {}),
          })
        if (cell.kind === 'number' && typeof cell.rawValue === 'number') {
          numericCount++
          const adjusted = cell.rawValue - correction,
            total = sum + adjusted
          correction = total - sum - adjusted
          sum = total
          if (!Number.isFinite(sum))
            throw new WorkbookError(
              'numeric_overflow',
              'The aggregate exceeds finite numeric precision.',
            )
          min = Math.min(min, cell.rawValue)
          max = Math.max(max, cell.rawValue)
        }
      }
    }
    row = rowEnd + 1
  }
  assertNotCancelled(input.signal)
  return {
    range: input.range,
    cellCount,
    nonEmptyCount,
    distinctCount: frequencies.size,
    numericCount,
    sum,
    mean: numericCount ? sum / numericCount : null,
    min: numericCount ? min : null,
    max: numericCount ? max : null,
    numericKinds: ['number'],
    topValues: [...frequencies.entries()]
      .sort((a, b) => b[1].count - a[1].count || a[0].localeCompare(b[0]))
      .slice(0, 50)
      .map(([, item]) => item),
  }
}
