import type { IExecutionInProgressParams } from '@univerjs/engine-formula'
import type { CalculationSourceData } from './source'
import { createCalculationSnapshot } from './snapshot'
import { calculateUniverWorkbook } from './univer'

/** Evaluate one fully loaded immutable/current logical root without publishing it. */
export async function calculateSourceWorkbook(
  source: CalculationSourceData,
  options: {
    signal?: AbortSignal
    sheetId?: string
    maxResultCells?: number
    onProgress?(progress: IExecutionInProgressParams): void
  } = {},
) {
  const snapshot = createCalculationSnapshot(source)
  return calculateUniverWorkbook(snapshot.workbook, {
    ...options,
    definedNames: snapshot.definedNames,
    tables: snapshot.tables,
  })
}
