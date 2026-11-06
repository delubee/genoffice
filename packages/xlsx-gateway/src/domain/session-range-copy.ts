import { formatAddress, parseRange } from './cell-address'
import { offsetFormulaRefs } from './formula-shift'
import {
  copyTargetBounds,
  expandToPrimitiveOps,
  type CopyRangeOperation,
  type FillRangeOperation,
} from './workbook-dsl'
import type { WorkbookSnapshot, CellState } from './workbook.types'
import type { cellRecordSchema } from '../shared/session-state'
import type { z } from 'zod'

type TypedCell = z.infer<typeof cellRecordSchema>

/** Copy values/formulas; ribbon fills explicitly request replacement donor formatting. */
export function copySessionRange(
  book: WorkbookSnapshot,
  op: CopyRangeOperation | FillRangeOperation,
  templates: Map<string, Record<string, TypedCell>>,
  signal?: AbortSignal,
) {
  expandToPrimitiveOps([op])
  if (op.op === 'copy_range') {
    if (op.filterColumn !== undefined || op.filterValues !== undefined)
      throw new Error('unsupported_filtered_copy')
  }
  const source = parseRange(op.source),
    target = op.op === 'copy_range' ? copyTargetBounds(op) : parseRange(op.target),
    sourceId = op.sourceSheetId ?? op.sheetId,
    sourceSheet = book.sheets.find((sheet) => sheet.id === sourceId),
    targetSheet = book.sheets.find((sheet) => sheet.id === op.sheetId)
  if (!sourceSheet || !targetSheet) throw new Error('unknown_sheet')
  const sourceTemplates = structuredClone(templates.get(sourceId) ?? {}),
    destinationTemplates = templates.get(op.sheetId) ?? {},
    cells: Record<string, CellState> = { ...targetSheet.cells },
    styles = { ...targetSheet.styles },
    sourceStyles = structuredClone(sourceSheet.styles ?? {}),
    copyFormat = op.op === 'fill_range' && op.copyFormat === true,
    height = source.endRow - source.startRow + 1,
    width = source.endColumn - source.startColumn + 1
  for (let row = target.startRow; row <= target.endRow; row++) {
    if (signal?.aborted) throw new Error('operation_cancelled')
    for (let column = target.startColumn; column <= target.endColumn; column++) {
      const sourceRow = source.startRow + ((row - target.startRow) % height),
        sourceColumn = source.startColumn + ((column - target.startColumn) % width),
        from = formatAddress(sourceRow, sourceColumn),
        to = formatAddress(row, column),
        donor = sourceSheet.cells[from] ?? { value: null },
        template = sourceTemplates[from]
      const value = donor.formula ? null : donor.value
      cells[to] = {
        value,
        ...(donor.formula
          ? { formula: offsetFormulaRefs(donor.formula, row - sourceRow, column - sourceColumn) }
          : {}),
      }
      if (copyFormat) {
        // Read from immutable source snapshots, including when the target tiles
        // the donor itself. Replacement removes destination-only formatting.
        styles[to] = structuredClone(sourceStyles[from] ?? {})
      }
      destinationTemplates[to] = {
        ...(destinationTemplates[to] ?? {}),
        kind:
          value === null
            ? 'blank'
            : template?.rawValue === value
              ? template.kind
              : typeof value === 'number'
                ? 'number'
                : typeof value === 'boolean'
                  ? 'boolean'
                  : 'string',
        rawValue: value,
        ...(template?.dateSystem ? { dateSystem: template.dateSystem } : {}),
        ...(copyFormat ? { styleReset: true, style: structuredClone(template?.style ?? {}) } : {}),
      }
      if (copyFormat) {
        delete destinationTemplates[to]!.numberFormat
        if (template?.numberFormat) destinationTemplates[to]!.numberFormat = template.numberFormat
      }
    }
  }
  templates.set(op.sheetId, destinationTemplates)
  return {
    target,
    snapshot: {
      ...book,
      sheets: book.sheets.map((sheet) =>
        sheet.id === op.sheetId ? { ...sheet, cells, ...(copyFormat ? { styles } : {}) } : sheet,
      ),
    },
  }
}
