import type { ICellData, IWorkbookData } from '@univerjs/core'
import { LexerTreeBuilder } from '@univerjs/engine-formula'
import { CalculationBindingError } from './names'

/** Expand shared followers in the private snapshot BEFORE name binding. This
 * keeps authored formula provenance and relative offsets independent of aliases.
 */
export function expandCalculationSharedFormulas(snapshot: Partial<IWorkbookData>) {
  let lexer: LexerTreeBuilder | undefined
  try {
    for (const sheet of Object.values(snapshot.sheets ?? {})) {
      const masters = new Map<string, { row: number; column: number; formula: string }>()
      const shared: { row: number; column: number; cell: ICellData }[] = []
      for (const [row, cells] of Object.entries(sheet.cellData ?? {}))
        for (const [column, raw] of Object.entries(cells ?? {})) {
          const cell = raw as ICellData | null | undefined
          if (!cell?.si) continue
          shared.push({ row: Number(row), column: Number(column), cell })
          if (cell.f) {
            if (masters.has(cell.si))
              throw new CalculationBindingError('calculation_ambiguous_shared_formula')
            masters.set(cell.si, { row: Number(row), column: Number(column), formula: cell.f })
          }
        }
      for (const { row, column, cell } of shared) {
        const master = masters.get(cell.si!)
        if (!master) throw new CalculationBindingError('calculation_missing_shared_formula')
        if (!cell.f) {
          lexer ??= new LexerTreeBuilder()
          cell.f = lexer.moveFormulaRefOffset(
            master.formula,
            column - master.column,
            row - master.row,
          )
        }
        delete cell.si
      }
    }
  } finally {
    lexer?.dispose()
  }
}
