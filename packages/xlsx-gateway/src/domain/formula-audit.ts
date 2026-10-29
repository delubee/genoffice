import { qualifierMatches } from '../gateway/xlsx-structure'
import { parseFormulaReferences } from './formula-references'

export interface CellArea {
  readonly startRow: number
  readonly endRow: number
  readonly startColumn: number
  readonly endColumn: number
}

export interface CellPoint {
  readonly row: number
  readonly column: number
}

export interface SheetBounds {
  readonly lastRow: number
  readonly lastColumn: number
  readonly maxRows: number
  readonly maxColumns: number
}

export interface PrecedentResult {
  /// Same-sheet referenced ranges, deduped, in formula order.
  readonly areas: CellArea[]
  /// References that resolve to another sheet — counted, not drawn.
  readonly offSheet: number
}

export function collectPrecedents(
  formula: string,
  sheetName: string,
  bounds: SheetBounds,
): PrecedentResult {
  const areas: CellArea[] = []
  const seen = new Set<string>()
  let offSheet = 0
  for (const reference of parseFormulaReferences(formula)) {
    if (reference.qualifier !== undefined && !qualifierMatches(reference.qualifier, sheetName)) {
      offSheet += 1
      continue
    }
    // Whole row/column references clamp to the used range so the overlay
    // stays finite; explicit refs clamp only to the sheet extent.
    const area = {
      startRow: Math.max(reference.token.startRow ?? 0, 0),
      endRow: Math.min(reference.token.endRow ?? Math.max(bounds.lastRow, 0), bounds.maxRows - 1),
      startColumn: Math.max(reference.token.startColumn ?? 0, 0),
      endColumn: Math.min(
        reference.token.endColumn ?? Math.max(bounds.lastColumn, 0),
        bounds.maxColumns - 1,
      ),
    }
    if (area.endRow < area.startRow || area.endColumn < area.startColumn) continue
    const key = `${area.startRow}:${area.startColumn}:${area.endRow}:${area.endColumn}`
    if (seen.has(key)) continue
    seen.add(key)
    areas.push(area)
  }
  return { areas, offSheet }
}

type CellDataMatrix = Record<
  number,
  | Record<
      number,
      | {
          readonly f?: unknown
        }
      | null
      | undefined
    >
  | null
  | undefined
>

export interface DependentsSource {
  readonly sheets: Record<
    string,
    | {
        readonly name?: string | undefined
        readonly cellData?: CellDataMatrix | undefined
      }
    | undefined
  >
}

export interface DependentResult {
  /// Formula cells on the active sheet that read the target cell.
  readonly cells: CellPoint[]
  /// Dependent formulas living on other sheets — counted, not drawn.
  readonly offSheet: number
}

export function collectDependents(
  source: DependentsSource,
  activeSheetName: string,
  target: CellPoint,
): DependentResult {
  const cells: CellPoint[] = []
  let offSheet = 0
  for (const sheet of Object.values(source.sheets)) {
    if (!sheet?.cellData) continue
    const onActiveSheet = sheet.name === activeSheetName
    for (const [rowKey, rowCells] of Object.entries(sheet.cellData)) {
      for (const [columnKey, cell] of Object.entries(rowCells ?? {})) {
        const formula = cell?.f
        if (typeof formula !== 'string' || !formula.startsWith('=')) continue
        const row = Number(rowKey)
        const column = Number(columnKey)
        if (onActiveSheet && row === target.row && column === target.column) continue
        if (!formulaReadsCell(formula, onActiveSheet, activeSheetName, target)) continue
        if (onActiveSheet) cells.push({ row, column })
        else offSheet += 1
      }
    }
  }
  return { cells, offSheet }
}

export function formulaReadsCell(
  formula: string,
  onActiveSheet: boolean,
  activeSheetName: string,
  target: CellPoint,
): boolean {
  for (const reference of parseFormulaReferences(formula)) {
    const targetsActiveSheet =
      reference.qualifier === undefined
        ? onActiveSheet
        : qualifierMatches(reference.qualifier, activeSheetName)
    if (!targetsActiveSheet) continue
    const { token } = reference
    if (
      (token.startRow ?? 0) <= target.row &&
      target.row <= (token.endRow ?? Number.MAX_SAFE_INTEGER) &&
      (token.startColumn ?? 0) <= target.column &&
      target.column <= (token.endColumn ?? Number.MAX_SAFE_INTEGER)
    ) {
      return true
    }
  }
  return false
}

export interface TraceArrow {
  readonly from: CellPoint
  readonly to: CellPoint
}
