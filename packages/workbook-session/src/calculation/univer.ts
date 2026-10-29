import { expandCalculationSharedFormulas } from './shared-formulas'
import { bindCalculationNames, type CalculationNameDefinition } from './names'
import { randomUUID } from 'node:crypto'
import {
  ICommandService,
  LocaleType,
  Univer,
  UniverInstanceType,
  type IWorkbookData,
  type ICellData,
  type Workbook,
  type Nullable,
} from '@univerjs/core'
import { UniverSheetsPlugin } from '@univerjs/sheets'
import {
  FormulaDataModel,
  FormulaExecutedStateType,
  ICalculateFormulaService,
  IFormulaRuntimeService,
  IDefinedNamesService,
  IFunctionService,
  ISuperTableService,
  SetFormulaCalculationStartMutation,
  UniverFormulaEnginePlugin,
  type BaseValueObject,
  type BaseReferenceObject,
  type ArrayValueObject,
  type StringValueObject,
  type FunctionVariantType,
  type IExecutionInProgressParams,
} from '@univerjs/engine-formula'
import {
  CalculationMode,
  UniverSheetsFormulaPlugin,
  ImageFormulaCellInterceptorController,
} from '@univerjs/sheets-formula'
import type { CalculationTable } from './tables'
import { assertCalculationTableReferences } from './tables'

export type CalculationScalar = (
  | { type: 'number'; value: number }
  | { type: 'string' | 'error'; value: string }
  | { type: 'boolean'; value: boolean }
  | { type: 'blank'; value: null }
) & {
  /** Computed formatting hint; it does not authorize rewriting authored styles. */
  numberFormat?: string
  hyperlink?: string
  image?: NonNullable<ReturnType<StringValueObject['getImageInfo']>>
}
export interface CalculatedFormula {
  sheetId: string
  row: number
  column: number
  formula: string
  values: CalculationScalar[][]
}
export interface UniverCalculationResult {
  workbookId: string
  scope: { kind: 'workbook' } | { kind: 'sheet'; sheetId: string }
  formulas: CalculatedFormula[]
}
export class UniverCalculationError extends Error {
  constructor(
    readonly code: string,
    message = code,
  ) {
    super(message)
    this.name = 'UniverCalculationError'
  }
}
function scalarValue(value: Nullable<BaseValueObject>): CalculationScalar {
  if (!value || value.isNull()) return { type: 'blank', value: null }
  if (value.isCube() || value.isLambda())
    throw new UniverCalculationError('calculation_unsupported_result')
  const raw = value.getValue()
  if (value.isError()) return { type: 'error', value: String(raw) }
  if (value.isBoolean() && typeof raw === 'boolean') return { type: 'boolean', value: raw }
  if (value.isNumber() && typeof raw === 'number' && Number.isFinite(raw))
    return { type: 'number', value: raw }
  if (value.isString() && typeof raw === 'string') return { type: 'string', value: raw }
  throw new UniverCalculationError('calculation_unsupported_result')
}
function scalar(value: Nullable<BaseValueObject>): CalculationScalar {
  const result = scalarValue(value)
  const pattern = value?.getPattern()
  if (pattern) result.numberFormat = pattern
  if (value?.isString()) {
    const text = value as StringValueObject
    if (text.isHyperlink()) result.hyperlink = text.getHyperlinkUrl()
    if (text.isImage()) {
      const image = text.getImageInfo()
      if (!image) throw new UniverCalculationError('calculation_image_result_missing')
      result.image = structuredClone(image)
    }
  }
  return result
}
function matrix(value: FunctionVariantType, limit: number): CalculationScalar[][] {
  if (!value.isReferenceObject() && !value.isArray()) return [[scalar(value as BaseValueObject)]]
  const array = value as BaseReferenceObject | ArrayValueObject
  const range = array.getRangePosition()
  const rows = range.endRow - range.startRow + 1,
    columns = range.endColumn - range.startColumn + 1
  if (rows < 1 || columns < 1) throw new UniverCalculationError('calculation_empty_result')
  const first = array.getFirstCell()
  if ((rows === 1 && columns === 1) || (first.isError() && first.getValue() === '#SPILL!'))
    return [[scalar(first)]]
  if (rows * columns > limit) throw new UniverCalculationError('calculation_result_limit')
  const values = Array.from({ length: rows }, () =>
    Array.from({ length: columns }, (): CalculationScalar => ({ type: 'blank', value: null })),
  )
  // Use the same public iterator as the runtime writer. Converting a reference
  // to an array first changes its single-empty-cell semantics (zero vs blank).
  array.iterator((cell, row, column) => {
    values[row - range.startRow]![column - range.startColumn] = scalar(cell)
  })
  return values
}

/** Host-only engine adapter over an already authoritative, hydrated snapshot.
 * Observes rich runtime results during the engine's own evaluation; it never
 * re-evaluates a formula to guess error types or spill values. Input is cloned.
 * Source hydration, canonical cache commits and Calculation UI are separate work. */
export async function calculateUniverWorkbook(
  input: Partial<IWorkbookData>,
  options: {
    signal?: AbortSignal
    definedNames?: readonly CalculationNameDefinition[]
    /** Complete logical table metadata from the same frozen source root. */
    tables?: readonly CalculationTable[]
    /** Desktop Calculate Sheet semantics, including its dependent formulas. */
    sheetId?: string
    maxResultCells?: number
    onProgress?(progress: IExecutionInProgressParams): void
  } = {},
): Promise<UniverCalculationResult> {
  if (options.signal?.aborted) throw new UniverCalculationError('calculation_cancelled')
  const budget = options.maxResultCells ?? 100_000
  if (!Number.isSafeInteger(budget) || budget < 1)
    throw new UniverCalculationError('calculation_invalid_limit')
  const snapshot = structuredClone(input)
  const engineUnitId = `calculation-${randomUUID()}`
  const workbookId = snapshot.id ?? engineUnitId
  // Univer also maintains module-level indexes keyed by unit ID. Concurrent
  // or later jobs must not reuse a caller's logical document ID there.
  snapshot.id = engineUnitId
  const sheetOrder = new Map(
    (snapshot.sheetOrder ?? Object.keys(snapshot.sheets ?? {})).map((id, index) => [id, index]),
  )
  expandCalculationSharedFormulas(snapshot)
  const authoredFormulas = new Map<string, string>()
  for (const [sheetId, sheet] of Object.entries(snapshot.sheets ?? {}))
    for (const [row, cells] of Object.entries(sheet?.cellData ?? {}))
      for (const [column, rawCell] of Object.entries(cells ?? {})) {
        const cell = rawCell as ICellData | null | undefined
        if (typeof cell?.f === 'string' && cell.f)
          authoredFormulas.set(`${sheetId}:${row}:${column}`, cell.f)
      }

  const nameBindings = options.definedNames?.length
    ? bindCalculationNames(snapshot, options.definedNames)
    : undefined
  const tables = options.tables ?? []
  const seenTables = new Set<string>()
  for (const table of tables) {
    const key = table.name.toLocaleLowerCase('en-US')
    if (seenTables.has(key)) throw new UniverCalculationError('calculation_ambiguous_table')
    seenTables.add(key)
  }
  for (const formula of authoredFormulas.values()) assertCalculationTableReferences(formula, tables)
  for (const definition of options.definedNames ?? [])
    assertCalculationTableReferences(definition.formula, tables)
  const univer = new Univer({
    locale: LocaleType.EN_US,
    locales: { [LocaleType.EN_US]: { sheets: {} } },
  })
  let restore: (() => void) | undefined
  const cleanup: (() => void)[] = []
  try {
    univer.registerPlugin(UniverSheetsPlugin)
    univer.registerPlugin(UniverFormulaEnginePlugin)
    // No warm-up calculation: NOW/RAND and arrays must not be evaluated once
    // for initialization and again merely to collect their results.
    univer.registerPlugin(UniverSheetsFormulaPlugin, {
      initialFormulaComputing: CalculationMode.NO_CALCULATION,
    })
    const workbook = univer.createUnit<IWorkbookData, Workbook>(
      UniverInstanceType.UNIVER_SHEET,
      snapshot,
    )
    const scopedSheet =
      options.sheetId === undefined ? undefined : workbook.getSheetBySheetId(options.sheetId)
    if (options.sheetId !== undefined && !scopedSheet)
      throw new UniverCalculationError('calculation_unknown_sheet')
    const injector = univer.__getInjector()
    // Node evaluation captures IMAGE metadata from engine results; it must not
    // run the browser renderer's image-size loader after result publication.
    injector.get(ImageFormulaCellInterceptorController).dispose()
    const superTables = injector.get(ISuperTableService)
    for (const table of tables)
      superTables.registerTable(engineUnitId, table.name, {
        sheetId: table.sheetId,
        range: table.range,
        titleMap: table.titleMap,
      })
    if (nameBindings)
      injector
        .get(IDefinedNamesService)
        .registerDefinedNames(
          engineUnitId,
          Object.fromEntries(nameBindings.definitions.map((name) => [name.id, name])),
        )
    const runtime = injector.get(IFormulaRuntimeService)
    if (nameBindings) {
      for (const executor of injector.get(IFunctionService).getExecutors().values()) {
        if (!executor.isAddress()) continue
        const originalLookup = executor.getDefinedName
        executor.getDefinedName = function (name) {
          const definition = nameBindings.resolve(name, runtime.currentSubUnitId)
          return definition
            ? {
                id: definition.name,
                name: definition.name,
                formulaOrRefString: definition.formula,
                ...(definition.sheetId ? { localSheetId: definition.sheetId } : {}),
              }
            : null
        }
        cleanup.push(() => {
          executor.getDefinedName = originalLookup
        })
      }
    }
    const calculator = injector.get(ICalculateFormulaService)
    const formulas = injector.get(FormulaDataModel)
    const commands = injector.get(ICommandService)
    const captured = new Map<string, CalculatedFormula>()
    let size = 0
    let failure: unknown
    let cancelled = false
    const original = runtime.setRuntimeData
    runtime.setRuntimeData = function (value) {
      const sheetId = runtime.currentSubUnitId,
        row = runtime.currentRow,
        column = runtime.currentColumn
      const beforeClear = runtime
        .getAllRuntimeData()
        .clearArrayFormulaCellData[runtime.currentUnitId]?.[sheetId]?.getValue(row, column)
      try {
        // Preserve the actual engine behavior, including blocked-spill handling.
        original.call(runtime, value)
        if (runtime.currentUnitId !== engineUnitId)
          throw new UniverCalculationError('calculation_unit_mismatch')
        const key = `${sheetId}:${row}:${column}`
        const priorSize = captured.get(key)?.values.reduce((n, row) => n + row.length, 0) ?? 0
        const afterClear = runtime
          .getAllRuntimeData()
          .clearArrayFormulaCellData[runtime.currentUnitId]?.[sheetId]?.getValue(row, column)
        const blocked =
          afterClear !== beforeClear &&
          afterClear?.v === '#SPILL!' &&
          (value.isArray() || value.isReferenceObject())
        const values: CalculationScalar[][] = blocked
          ? [[{ type: 'error', value: '#SPILL!' }]]
          : matrix(value, budget - size + priorSize)
        const formula =
          authoredFormulas.get(`${sheetId}:${row}:${column}`) ??
          formulas.getFormulaStringByCell(row, column, sheetId, runtime.currentUnitId)
        if (!formula) throw new UniverCalculationError('calculation_formula_missing')
        const nextSize = values.reduce((n, row) => n + row.length, 0)
        if (size - priorSize + nextSize > budget)
          throw new UniverCalculationError('calculation_result_limit')
        size += nextSize - priorSize
        captured.set(key, { sheetId, row, column, formula, values })
      } catch (error) {
        failure ??= error
        calculator.stopFormulaExecution()
      }
    }
    restore = () => {
      runtime.setRuntimeData = original
    }
    return await new Promise<UniverCalculationResult>((resolve, reject) => {
      let complete = false
      const finish = (error?: unknown) => {
        if (complete) return
        complete = true
        if (error) reject(error)
        else
          resolve({
            workbookId,
            scope: scopedSheet
              ? { kind: 'sheet', sheetId: scopedSheet.getSheetId() }
              : { kind: 'workbook' },
            formulas: [...captured.values()].sort(
              (a, b) =>
                (sheetOrder.get(a.sheetId) ?? Number.MAX_SAFE_INTEGER) -
                  (sheetOrder.get(b.sheetId) ?? Number.MAX_SAFE_INTEGER) ||
                a.row - b.row ||
                a.column - b.column,
            ),
          })
      }
      const completed = calculator.executionCompleteListener$.subscribe((result) => {
        if (result.functionsExecutedState === FormulaExecutedStateType.INITIAL) return
        if (failure) finish(failure)
        else if (
          cancelled ||
          result.functionsExecutedState === FormulaExecutedStateType.STOP_EXECUTION
        )
          finish(new UniverCalculationError('calculation_cancelled'))
        else finish()
      })
      const progress = calculator.executionInProgressListener$.subscribe((value) => {
        // The engine resets its stop flag just after START. Reassert a cancel
        // or observer failure at the following stages so that reset cannot
        // turn a cancelled job back into work.
        if (cancelled || failure) calculator.stopFormulaExecution()
        try {
          options.onProgress?.({ ...value })
        } catch (error) {
          failure ??= error
          calculator.stopFormulaExecution()
        }
      })
      cleanup.push(() => {
        completed.unsubscribe()
        progress.unsubscribe()
      })
      const cancel = () => {
        cancelled = true
        calculator.stopFormulaExecution()
      }
      options.signal?.addEventListener('abort', cancel, { once: true })
      cleanup.push(() => options.signal?.removeEventListener('abort', cancel))
      const deadline = setTimeout(() => {
        failure ??= new UniverCalculationError('calculation_timeout')
        calculator.stopFormulaExecution()
        // Stopping normally publishes completion on the engine's next yield.
        // This fallback also releases a failed engine that cannot publish it.
        const fallback = setTimeout(() => finish(failure), 1000)
        cleanup.push(() => clearTimeout(fallback))
      }, 20_000)
      cleanup.push(() => clearTimeout(deadline))
      try {
        commands.syncExecuteCommand(
          SetFormulaCalculationStartMutation.id,
          scopedSheet
            ? {
                commands: [],
                dirtyRanges: [
                  {
                    unitId: engineUnitId,
                    sheetId: scopedSheet.getSheetId(),
                    range: {
                      startRow: 0,
                      endRow: scopedSheet.getRowCount() - 1,
                      startColumn: 0,
                      endColumn: scopedSheet.getColumnCount() - 1,
                    },
                  },
                ],
                dirtyNameMap: {
                  [engineUnitId]: { [scopedSheet.getSheetId()]: scopedSheet.getName() },
                },
              }
            : { commands: [], forceCalculation: true },
        )
      } catch (error) {
        finish(error)
      }
    })
  } finally {
    // Engine completion publishes async command notifications. Drain their
    // microtasks before disposing the trigger controller/locale service, so
    // its pending progress timer can be cleared by the terminal notification.
    await new Promise<void>((resolve) => setImmediate(resolve))
    for (const dispose of cleanup.reverse()) dispose()
    restore?.()
    univer.dispose()
  }
}
