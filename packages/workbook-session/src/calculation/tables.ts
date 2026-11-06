import { z } from 'zod'
import type { ISuperTable } from '@univerjs/engine-formula'
import { MAX_GRID_COLUMNS, MAX_GRID_ROWS } from '@genoffice/xlsx-gateway/shared/grid-bounds'
import { WorkbookError } from '../domain'

const tableFeatureSchema = z
  .object({
    range: z
      .object({
        startRow: z.number().int().nonnegative(),
        startColumn: z.number().int().nonnegative(),
        endRow: z.number().int().nonnegative(),
        endColumn: z.number().int().nonnegative(),
      })
      .strict(),
    headerRowCount: z.number().int().nonnegative(),
    totalsRowCount: z.number().int().nonnegative().optional(),
    name: z.string().min(1).optional(),
    columns: z.array(z.string()).optional(),
  })
  .passthrough()

export interface CalculationTable extends ISuperTable {
  name: string
  totalsRowCount: number
}

/**
 * Bind the source table metadata accepted by the installed formula engine.
 * That engine models one header row and no totals-row count, so other layouts
 * fail closed instead of returning plausible but Excel-inaccurate results.
 */
export function calculationTablesFromFeatures(raw: unknown, sheetId: string): CalculationTable[] {
  const features = z
    .object({ tables: z.array(tableFeatureSchema).default([]) })
    .passthrough()
    .parse(raw)
  return features.tables.map((table) => {
    const { range, headerRowCount, totalsRowCount = 0, name, columns } = table
    if (
      !name ||
      headerRowCount !== 1 ||
      totalsRowCount !== 0 ||
      !columns ||
      range.endRow >= MAX_GRID_ROWS ||
      range.endColumn >= MAX_GRID_COLUMNS ||
      range.startRow > range.endRow ||
      range.startColumn > range.endColumn ||
      columns.length !== range.endColumn - range.startColumn + 1 ||
      columns.some((column) => !column.trim()) ||
      new Set(columns.map((column) => column.toLocaleLowerCase('en-US'))).size !== columns.length
    )
      throw new WorkbookError(
        'calculation_unavailable',
        'This table layout cannot be represented exactly by the installed formula engine.',
      )
    return {
      name,
      sheetId,
      range,
      totalsRowCount,
      titleMap: new Map(columns.map((column, index) => [column, index])),
    }
  })
}

/** Reject #Totals references: the engine assumes the final row is a totals row. */
export function assertCalculationTableReferences(
  formula: string,
  tables: readonly CalculationTable[],
) {
  // Keep offsets stable while excluding formula string literals such as
  // `="Table[#Totals]"`, which are text rather than structured references.
  const searchable = formula.split('')
  for (let index = 0; index < searchable.length; index++) {
    if (searchable[index] !== '"') continue
    searchable[index] = ' '
    for (index++; index < searchable.length; index++) {
      if (searchable[index] !== '"') searchable[index] = ' '
      else if (searchable[index + 1] === '"') {
        searchable[index] = ' '
        searchable[++index] = ' '
      } else {
        searchable[index] = ' '
        break
      }
    }
  }
  const text = searchable.join('')
  for (const table of tables) {
    const escaped = table.name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
    const start = new RegExp(`(?:^|[^\\p{L}\\p{N}_.])${escaped}\\s*\\[`, 'giu')
    for (const match of text.matchAll(start)) {
      const open = text.indexOf('[', match.index! + match[0].length - 1)
      let depth = 0
      for (let index = open; index < text.length; index++) {
        if (text[index] === '[') depth++
        else if (text[index] === ']' && --depth === 0) {
          if (
            text
              .slice(open, index + 1)
              .toLowerCase()
              .includes('#totals')
          )
            throw new WorkbookError(
              'calculation_unavailable',
              'Structured references to table totals are not supported by the calculation engine.',
            )
          break
        }
      }
    }
  }
}
