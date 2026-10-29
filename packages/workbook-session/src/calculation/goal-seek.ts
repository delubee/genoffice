import { solveGoalSeekSamples } from '@genoffice/xlsx-gateway/domain/goal-seek'
import { parseAddress } from '@genoffice/xlsx-gateway/domain/cell-address'
import { addressSchema } from '@genoffice/xlsx-gateway/shared/session-state'
import { WorkbookError, assertNotCancelled } from '../domain'
import type { CalculationSourceData } from './source'
import { calculateSourceWorkbook } from './evaluate-source'

/** Solve on a private source copy. The caller must validate permissions and the
 * returned root again before committing the final scalar as a normal transaction. */
export async function seekCalculationSource(
  source: CalculationSourceData,
  input: { sheetId: string; setCell: string; byCell: string; toValue: number },
  options: { isCurrent(): boolean; signal?: AbortSignal; maxResultCells?: number },
) {
  const frozen = structuredClone(source)
  const signal = AbortSignal.any([
    ...(options.signal ? [options.signal] : []),
    AbortSignal.timeout(30_000),
  ])
  const check = () => {
    assertNotCancelled(signal)
    if (!options.isCurrent())
      throw new WorkbookError('revision_conflict', 'The document changed during Goal Seek.')
  }
  check()
  const setCell = addressSchema.parse(input.setCell.toUpperCase())
  const byCell = addressSchema.parse(input.byCell.toUpperCase())
  const sheet = frozen.sheets.find((item) => item.id === input.sheetId)
  if (!sheet) throw new WorkbookError('unknown_sheet', 'The Goal Seek worksheet no longer exists.')
  if (!sheet.cells[setCell]?.formula)
    throw new WorkbookError('invalid_operation', 'The Goal Seek target must contain a formula.')
  const original = sheet.cells[byCell]
  if (original?.formula || (original?.rawValue != null && typeof original.rawValue !== 'number'))
    throw new WorkbookError(
      'invalid_operation',
      'The changing cell must contain a number or be blank.',
    )
  const target = parseAddress(setCell)
  const changing = parseAddress(byCell)
  sheet.rows = Math.max(sheet.rows, changing.row + 1)
  sheet.columns = Math.max(sheet.columns, changing.column + 1)
  const result = await solveGoalSeekSamples(
    async (guess) => {
      check()
      sheet.cells[byCell] = { ...original, kind: 'number', rawValue: guess }
      const calculated = await calculateSourceWorkbook(frozen, {
        signal,
        ...(options.maxResultCells === undefined ? {} : { maxResultCells: options.maxResultCells }),
      })
      check()
      const formula = calculated.formulas.find(
        (cell) =>
          cell.sheetId === sheet.id && cell.row === target.row && cell.column === target.column,
      )
      if (!formula || formula.values.length !== 1 || formula.values[0]?.length !== 1)
        throw new WorkbookError(
          'calculation_unavailable',
          'Goal Seek requires a scalar formula result.',
        )
      const value = formula.values[0][0]!
      return value.type === 'number' ? value.value : Number.NaN
    },
    typeof original?.rawValue === 'number' ? original.rawValue : 0,
    input.toValue,
  )
  check()
  return { ...result, stateId: frozen.stateId, sheetId: sheet.id, address: byCell }
}
