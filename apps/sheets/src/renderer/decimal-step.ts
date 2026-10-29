/**
 * Excel-parity Increase/Decrease Decimal for the ribbon buttons.
 *
 * Univer's Add/SubtractDecimalCommand derive a General-format cell's current
 * decimal count with `getDecimalFromPattern`, which asks the numfmt engine for
 * the pattern's `maxDecimals` — and `getFormatInfo('General').maxDecimals` is
 * 9. A cell with no numfmt record takes the raw-value branch (0.77 → 2), but a
 * cell carrying an explicit "General" record (what the file gateway writes for
 * numFmtId 0) reads as 9, so one Increase Decimal click jumps 341504.77 to
 * 341504.7700000000. This reimplements both commands verbatim except that an
 * explicit default-format record goes through the same raw-value branch as a
 * missing one.
 */
import { Range } from '@univerjs/core'
import { INumfmtService, SheetsSelectionsService } from '@univerjs/sheets'
import { decimalStepValues, type DecimalCell } from './decimal-step-model'
export { rawValueDecimals } from './decimal-step-model'

import type { UniverRuntime } from './univer-state'

export function stepSelectionDecimals(runtime: UniverRuntime, step: 1 | -1): void {
  const injector = runtime.univer.__getInjector()
  const selectionsService = injector.get(SheetsSelectionsService)
  const numfmtService = injector.get(INumfmtService)
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const facadeSheet = workbook?.getActiveSheet()
  if (!workbook || !facadeSheet) return
  const unitId = workbook.getId()
  const subUnitId = facadeSheet.getSheetId()
  const worksheet = facadeSheet.getSheet()
  const selections = selectionsService.getCurrentSelections()
  if (!selections || selections.length === 0) return

  const cells: DecimalCell[] = []
  for (const selection of selections)
    Range.foreach(selection.range, (row, col) => {
      cells.push({
        row,
        col,
        pattern: numfmtService.getValue(unitId, subUnitId, row, col)?.pattern,
        raw: worksheet.getCellRaw(row, col) ?? null,
      })
    })
  const values = decimalStepValues(cells, step)
  if (values.length === 0) return
  void runtime.univerAPI.executeCommand('sheet.command.numfmt.set.numfmt', {
    unitId,
    subUnitId,
    values,
  })
}
