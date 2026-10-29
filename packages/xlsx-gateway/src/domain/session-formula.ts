import { columnIndex } from './cell-address'
import { renameSheetInFormula } from '../gateway/xlsx-sheets'
import { shiftFormulaText } from '../gateway/xlsx-structure'
import type { CandidateState as WorkbookState } from '../shared/session-changeset'

export function shiftedSessionFormula(
  state: Readonly<WorkbookState>,
  sheetId: string,
  formula: string,
) {
  const names = new Map(
    state.sheets.map((sheet) => [
      sheet.id,
      (sheet.features?.originalMetadata as { name?: string } | undefined)?.name ?? sheet.name,
    ]),
  )
  for (const op of state.structuralLog) {
    if (!('sheetId' in op)) continue
    const target = state.sheets.find((sheet) => sheet.id === op.sheetId)
    if (!target) continue
    if (op.op === 'rename_sheet') {
      formula = renameSheetInFormula(formula, names.get(op.sheetId)!, op.name)
      names.set(op.sheetId, op.name)
      continue
    }
    const rows = op.op === 'insert_rows' || op.op === 'delete_rows'
    const columns = op.op === 'insert_cols' || op.op === 'delete_cols'
    if (!rows && !columns) continue
    const index = 'row' in op ? op.row - 1 : 'column' in op ? columnIndex(op.column) : 0
    const count = 'count' in op ? op.count : 0,
      deleting = op.op.startsWith('delete_')
    formula = shiftFormulaText(
      formula,
      names.get(op.sheetId)!,
      {
        boundary: deleting ? index + count : index,
        delta: deleting ? -count : count,
        deleted: deleting ? { start: index, end: index + count - 1 } : null,
      },
      rows ? 'row' : 'column',
      sheetId !== op.sheetId,
    )
  }
  return formula
}
