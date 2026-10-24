import {
  IUniverInstanceService,
  LocaleType,
  Univer,
  UniverInstanceType,
  type IWorkbookData,
  type Workbook,
} from '@univerjs/core'
import { UniverSheetsPlugin } from '@univerjs/sheets'
import { parseAddress } from '@genoffice/xlsx-gateway/domain/cell-address'
import type { WorkbookSnapshot } from '@genoffice/xlsx-gateway/domain/workbook.types'
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import {
  transactionInputSchema,
  workbookStateSchema,
} from '@genoffice/xlsx-gateway/shared/session-state'
import {
  computeCandidateState,
  type CandidateMaterial,
} from '@genoffice/xlsx-gateway/domain/session-candidate'
import type { CandidateState } from './changeset'
export type { CandidateMaterial } from '@genoffice/xlsx-gateway/domain/session-candidate'

export interface CandidateContext {
  hydrate(
    state: Readonly<CandidateState>,
    operations: readonly WorkbookOperation[],
    signal?: AbortSignal,
  ): Promise<CandidateMaterial>
  fault?(point: string): void
}

function univerData(book: WorkbookSnapshot): Partial<IWorkbookData> {
  return {
    id: 'disposable-candidate',
    sheetOrder: book.sheets.map((sheet) => sheet.id),
    sheets: Object.fromEntries(
      book.sheets.map((sheet) => {
        const cellData: NonNullable<IWorkbookData['sheets'][string]['cellData']> = {}
        for (const [address, cell] of Object.entries(sheet.cells)) {
          const { row, column } = parseAddress(address)
          ;(cellData[row] ??= {})[column] = {
            ...(cell.value === null ? {} : { v: cell.value }),
            ...(cell.formula ? { f: cell.formula } : {}),
          }
        }
        return [
          sheet.id,
          {
            id: sheet.id,
            name: sheet.name,
            rowCount: sheet.gridRows ?? 1000,
            columnCount: sheet.gridColumns ?? 20,
            cellData,
          },
        ]
      }),
    ),
  }
}

export class IsolatedCandidateInterpreter {
  private closed = false
  private active = new Set<Univer>()
  constructor(private readonly context: CandidateContext) {}
  get activeRuntimeCount() {
    return this.active.size
  }
  async prepare(inputState: Readonly<CandidateState>, rawIntent: unknown, signal?: AbortSignal) {
    if (this.closed) throw new Error('candidate_disposed')
    const state = workbookStateSchema.parse(inputState),
      intent = transactionInputSchema.parse(rawIntent)
    const material = await this.context.hydrate(state, intent.operations, signal)
    if (this.closed) throw new Error('candidate_disposed')
    const runtime = new Univer({
      locale: LocaleType.EN_US,
      locales: { [LocaleType.EN_US]: { sheets: {} } },
    })
    this.active.add(runtime)
    try {
      runtime.registerPlugin(UniverSheetsPlugin)
      let model: Workbook | undefined
      return await computeCandidateState(state, intent, material, {
        signal,
        fault: this.context.fault,
        afterStep: (book) => {
          if (this.closed) throw new Error('candidate_disposed')
          if (model)
            runtime.__getInjector().get(IUniverInstanceService).disposeUnit(model.getUnitId())
          model = runtime.createUnit<IWorkbookData, Workbook>(
            UniverInstanceType.UNIVER_SHEET,
            univerData(book),
          )
          if (model.getSnapshot().sheetOrder.join() !== book.sheets.map((sheet) => sheet.id).join())
            throw new Error('candidate_projection_mismatch')
        },
      })
    } finally {
      if (this.active.delete(runtime)) runtime.dispose()
    }
  }
  dispose() {
    this.closed = true
    for (const runtime of this.active) runtime.dispose()
    this.active.clear()
  }
}
