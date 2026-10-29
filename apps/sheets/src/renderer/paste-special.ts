import {
  type CellGrid,
  type PasteSpecialOptions,
  type PasteSpecialSource,
  type SourceCell,
} from './paste-special-model'
export {
  PASTE_SPECIAL_OPEN_SHELL_COMMAND,
  PASTE_SPECIAL_OPERATIONS,
  PASTE_SPECIAL_TYPES,
  UNSUPPORTED_TYPES,
  pasteSpecialAvailability,
  type CellGrid,
  type PasteSpecialAvailability,
  type PasteSpecialType,
  type PasteSpecialOperation,
  type PasteSpecialOptions,
  type PasteSpecialSource,
  type SourceCell,
} from './paste-special-model'
/**
 * Excel's Paste Special dialog. Univer already ships the plain paste
 * variants (values / formulas / formats / column widths / all except
 * borders) as commands that read the system clipboard; the dialog delegates
 * to them whenever no extra switch is on. Operation, Skip blanks, Transpose,
 * Paste Link and the "+ number formats" variants are ours: the copy source is
 * resolved from Univer's copy cache (or the clipboard's plain text when the
 * copy came from another app), cells are transformed here, and one
 * SetRangeValuesCommand writes the block, so it undoes in one step and
 * reaches the edit journal like any paste.
 *
 * Limits: Operation works on values (a formula in the source or target
 * contributes its cached value; Excel would wrap the formulas). A transposed
 * paste writes a formula cell's value too, since relative references would
 * need rotating, not shifting. Merged cells in the source are pasted as
 * their anchor cell. A block past the sheet edge inserts the missing
 * rows/columns first, so that paste undoes in two steps.
 */
import {
  Direction,
  ICommandService,
  IUniverInstanceService,
  type ICellData,
  type Workbook,
  type Worksheet,
} from '@univerjs/core'
import { LexerTreeBuilder } from '@univerjs/engine-formula'
import {
  InsertColCommand,
  InsertRowCommand,
  SelectionMoveType,
  SetRangeValuesCommand,
  SetSelectionsOperation,
} from '@univerjs/sheets'
import { COPY_TYPE, ISheetClipboardService } from '@univerjs/sheets-ui'

import { sharedFormulaResolverFor } from './shared-formula-journal'
import type { UniverRuntime } from './univer-state'

import {
  UNIVER_DELEGATE,
  needsOwnLayer,
  buildPasteGrid,
  parseClipboardText,
  textCell,
  type SourceGridCell,
  type RelocateFormula,
} from './paste-special-transform'
export * from './paste-special-transform'

interface CopyCache {
  get(id: string):
    | {
        unitId: string
        subUnitId: string
        range: { rows: number[]; cols: number[] }
        copyType: unknown
      }
    | undefined
  getLastCopyId(): string | null
}

function copyCache(runtime: UniverRuntime): CopyCache | undefined {
  // Univer keeps the cache private on its clipboard service singleton.
  return (
    runtime.univer.__getInjector().get(ISheetClipboardService) as unknown as {
      _copyContentCache?: CopyCache
    }
  )._copyContentCache
}

function internalSource(
  cache: CopyCache | undefined,
  copyId: string | null,
): PasteSpecialSource | null {
  const entry = copyId ? cache?.get(copyId) : undefined
  if (!entry || !copyId) return null
  return {
    kind: 'internal',
    unitId: entry.unitId,
    subUnitId: entry.subUnitId,
    rows: entry.range.rows,
    cols: entry.range.cols,
    cut: entry.copyType === COPY_TYPE.CUT,
  }
}

/// The copy the dialog will paste: Univer's cache when the clipboard still
/// carries its copy id, the clipboard text when another app wrote it, else
/// the last in-app copy (the Clipboard API may be unreadable or hold only
/// an image while Univer's cache still has the copy).
export async function readPasteSpecialSource(runtime: UniverRuntime): Promise<PasteSpecialSource> {
  const cache = copyCache(runtime)
  const lastCopy = internalSource(cache, cache?.getLastCopyId() ?? null)
  let html = ''
  let text = ''
  try {
    const items = await navigator.clipboard.read()
    const item = items[0]
    if (item) {
      if (item.types.includes('text/html')) html = await (await item.getType('text/html')).text()
      if (item.types.includes('text/plain')) text = await (await item.getType('text/plain')).text()
    }
  } catch {
    return lastCopy ?? { kind: 'none' }
  }
  const copyId = /data-copy-id="([^\s"]+)"/.exec(html)?.[1] ?? null
  const internal = internalSource(cache, copyId)
  if (internal) return internal
  if (text.trim() !== '') return { kind: 'text', rows: parseClipboardText(text) }
  return lastCopy ?? { kind: 'none' }
}

function readSourceGrid(
  runtime: UniverRuntime,
  source: Extract<PasteSpecialSource, { kind: 'internal' }>,
  workbook: Workbook,
): CellGrid<SourceGridCell> | null {
  const worksheet = workbook.getSheetBySheetId(source.subUnitId)
  if (!worksheet) return null
  const resolveShared = sharedFormulaResolverFor(runtime, source.subUnitId)
  return source.rows.map((row) =>
    source.cols.map((column) => {
      const raw = worksheet.getCellRaw(row, column)
      const shown = worksheet.getCell(row, column)
      const f = raw?.f ?? (raw?.si ? resolveShared(row, column, raw.si) : null) ?? null
      const style = worksheet.getComposedCellStyleByCellData(row, column, raw)
      const cell: SourceCell =
        raw || shown
          ? {
              v: shown?.v ?? raw?.v ?? null,
              t: shown?.t ?? raw?.t ?? null,
              f,
              p: raw?.p ?? null,
              s: style && Object.keys(style).length > 0 ? style : null,
            }
          : null
      return { cell, row, column }
    }),
  )
}

export type PasteSpecialResult =
  { ok: true } | { ok: false; reason: 'no-selection' | 'out-of-bounds' | 'nothing' }

/// Runs the dialog's choice; Univer's own command for plain variants.
export async function applyPasteSpecial(
  runtime: UniverRuntime,
  source: PasteSpecialSource,
  options: PasteSpecialOptions,
): Promise<PasteSpecialResult> {
  if (source.kind === 'none') return { ok: false, reason: 'nothing' }
  if (!needsOwnLayer(options) || (source.kind === 'internal' && source.cut)) {
    const id = UNIVER_DELEGATE[options.type] ?? UNIVER_DELEGATE.all!
    await runtime.univerAPI.executeCommand(id)
    return { ok: true }
  }
  const injector = runtime.univer.__getInjector()
  const fWorkbook = runtime.univerAPI.getActiveWorkbook()
  const fSheet = fWorkbook?.getActiveSheet()
  const selection = fWorkbook?.getActiveRange()
  if (!fWorkbook || !fSheet || !selection) return { ok: false, reason: 'no-selection' }
  const targetSheet: Worksheet = fSheet.getSheet()
  let grid: CellGrid<SourceGridCell> | null
  let linkSheetName: string | null = null
  if (source.kind === 'internal') {
    const sourceWorkbook = injector.get(IUniverInstanceService).getUnit<Workbook>(source.unitId)
    if (!sourceWorkbook) return { ok: false, reason: 'nothing' }
    grid = readSourceGrid(runtime, source, sourceWorkbook)
    const sourceSheet = sourceWorkbook.getSheetBySheetId(source.subUnitId)
    linkSheetName =
      sourceSheet && source.subUnitId !== targetSheet.getSheetId() ? sourceSheet.getName() : null
  } else {
    grid = source.rows.map((row) =>
      row.map((text) => ({ cell: textCell(text), row: null, column: null })),
    )
  }
  if (!grid) return { ok: false, reason: 'nothing' }
  const lexer = injector.get(LexerTreeBuilder)
  const relocate: RelocateFormula = (formula, rowDelta, columnDelta) =>
    lexer.moveFormulaRefOffset(formula, columnDelta, rowDelta)
  const startRow = selection.getRow()
  const startColumn = selection.getColumn()
  const built = buildPasteGrid(
    grid,
    options,
    {
      startRow,
      startColumn,
      rows: selection.getHeight(),
      columns: selection.getWidth(),
      cellAt: (row, column) => targetSheet.getCell(row, column),
    },
    relocate,
    linkSheetName,
  )
  if (built.rows === 0) return { ok: false, reason: 'nothing' }
  const endRow = startRow + built.rows - 1
  const endColumn = startColumn + built.columns - 1
  const unitId = fWorkbook.getId()
  const subUnitId = targetSheet.getSheetId()
  // Like Univer's paste, grow the sheet instead of refusing the block.
  const rowCount = targetSheet.getRowCount()
  if (endRow >= rowCount) {
    await runtime.univerAPI.executeCommand(InsertRowCommand.id, {
      unitId,
      subUnitId,
      direction: Direction.DOWN,
      range: { startRow: rowCount, endRow, startColumn: 0, endColumn: 0 },
    })
  }
  const columnCount = targetSheet.getColumnCount()
  if (endColumn >= columnCount) {
    await runtime.univerAPI.executeCommand(InsertColCommand.id, {
      unitId,
      subUnitId,
      direction: Direction.RIGHT,
      range: { startRow: 0, endRow: 0, startColumn: columnCount, endColumn },
    })
  }
  if (endRow >= targetSheet.getRowCount() || endColumn >= targetSheet.getColumnCount()) {
    return { ok: false, reason: 'out-of-bounds' }
  }
  const value: Record<number, Record<number, ICellData>> = {}
  built.cells.forEach((line, r) => {
    line.forEach((cell, c) => {
      if (!cell) return
      ;(value[startRow + r] ??= {})[startColumn + c] = cell
    })
  })
  const range = { startRow, startColumn, endRow, endColumn }
  if (Object.keys(value).length > 0) {
    await runtime.univerAPI.executeCommand(SetRangeValuesCommand.id, {
      unitId,
      subUnitId,
      range,
      value,
    })
  }
  injector.get(ICommandService).syncExecuteCommand(SetSelectionsOperation.id, {
    unitId,
    subUnitId,
    type: SelectionMoveType.MOVE_END,
    selections: [{ range, primary: null }],
  })
  return { ok: true }
}
