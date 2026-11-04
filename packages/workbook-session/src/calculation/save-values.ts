import { formatAddress } from '@genoffice/xlsx-gateway/domain/cell-address'
import type { WorkbookState } from '../domain'
import type { WorkbookSource } from '../ports'
import { loadCalculationSource } from './source'
import { calculateSourceWorkbook } from './evaluate-source'

/** One freshly calculated formula result, at the cell's final logical position. */
export interface SaveCalculationValue {
  sheetId: string
  row: number
  column: number
  value: string | number | boolean | null | { error: string }
}

/**
 * Fresh results for every scalar formula of the frozen root, so the saved file's cached values
 * agree with its inputs for readers without a formula engine. Returns null when the root cannot
 * be calculated exactly (unsupported resources, budget, cancellation): the save then keeps its
 * conservative behavior of leaving the caches empty for the recalculating application.
 */
export async function calculateSaveValues(input: {
  state: Readonly<WorkbookState>
  source: WorkbookSource
  sourceSessionId: string
  signal?: AbortSignal
}): Promise<SaveCalculationValue[] | null> {
  const { signal } = input
  try {
    const loaded = await loadCalculationSource({
      state: input.state,
      source: input.source,
      sourceSessionId: input.sourceSessionId,
      isCurrent: () => !signal?.aborted,
      ...(signal ? { signal } : {}),
    })
    if (!loaded.sheets.some((sheet) => Object.values(sheet.cells).some((cell) => cell.formula)))
      return []
    const calculated = await calculateSourceWorkbook(loaded, signal ? { signal } : {})
    const values: SaveCalculationValue[] = []
    for (const formula of calculated.formulas) {
      // Arrays write their spill through ordinary cells; only scalar formulas refresh a cached value.
      if (formula.values.length !== 1 || formula.values[0]?.length !== 1) continue
      const scalar = formula.values[0][0]!
      // An unknown function evaluates to #NAME?; that is the engine's gap, not the document's result.
      if (scalar.type === 'error' && scalar.value === '#NAME?') continue
      values.push({
        sheetId: formula.sheetId,
        row: formula.row,
        column: formula.column,
        value:
          scalar.type === 'error'
            ? { error: scalar.value }
            : scalar.type === 'blank'
              ? null
              : scalar.value,
      })
    }
    return values
  } catch {
    return null
  }
}

/** The same results keyed by `${sheetId}!${address}`, for reads of the frozen root. */
export async function calculateLiveValues(input: {
  state: Readonly<WorkbookState>
  source: WorkbookSource
  sourceSessionId: string
  signal?: AbortSignal
}): Promise<Map<string, SaveCalculationValue['value']> | null> {
  const values = await calculateSaveValues(input)
  if (!values) return null
  return new Map(
    values.map((item) => [`${item.sheetId}!${formatAddress(item.row, item.column)}`, item.value]),
  )
}
