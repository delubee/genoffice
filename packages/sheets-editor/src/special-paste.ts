import { CellValueType } from '@univerjs/core'
import {
  buildPasteGrid,
  pasteTypeCarriesValues,
  transposeGrid,
  textCell,
} from '../../../apps/sheets/src/renderer/paste-special-transform'
import {
  pasteSpecialAvailability,
  type PasteSpecialOptions,
  type PasteSpecialSource,
} from '../../../apps/sheets/src/renderer/paste-special-model'
import { formatAddress, parseRange } from '@genoffice/xlsx-gateway/domain/cell-address'
import { offsetFormulaRefs } from '@genoffice/xlsx-gateway/domain/formula-shift'
import {
  workbookOperationSchema,
  type WorkbookOperation,
} from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { cellRecordSchema } from '@genoffice/xlsx-gateway/shared/cell-record'
import { t } from '../../../apps/sheets/src/renderer/i18n/locale'
import { gridCell, type ViewCell } from './grid-cell'
import type { CopiedBlock } from './clipboard-copy'
import type { z } from 'zod'

type Cell = z.infer<typeof cellRecordSchema>
export interface SpecialPasteSnapshot {
  source: PasteSpecialSource
  block?: CopiedBlock
  sheetId?: string
  widths?: number[]
  html: string
  text: string
}
export const defaultPasteSpecial: PasteSpecialOptions = {
  type: 'all',
  operation: 'none',
  skipBlanks: false,
  transpose: false,
  link: false,
}

function sourceCells(snapshot: SpecialPasteSnapshot): Cell[][] {
  if (snapshot.block) return snapshot.block.cells
  if (snapshot.source.kind !== 'text') return []
  return snapshot.source.rows.map((row) =>
    row.map((text) => {
      const value = textCell(text)
      return value?.t === CellValueType.BOOLEAN
        ? { kind: 'boolean', rawValue: value.v === 1 }
        : typeof value?.v === 'number'
          ? { kind: 'number', rawValue: value.v }
          : typeof value?.v === 'string'
            ? { kind: 'string', rawValue: value.v }
            : { kind: 'blank', rawValue: null }
    }),
  )
}
export function specialPasteBounds(
  snapshot: SpecialPasteSnapshot,
  options: PasteSpecialOptions,
  address: string,
) {
  const source = sourceCells(snapshot),
    selected = parseRange(address)
  let height = source.length,
    width = Math.max(0, ...source.map((row) => row.length))
  if (options.type !== 'col-widths' && options.transpose && !options.link)
    [height, width] = [width, height]
  if (!height || !width) throw new Error(t('appClipboardUnsupported'))
  const repeat =
    (selected.endRow - selected.startRow + 1) % height === 0 &&
    (selected.endColumn - selected.startColumn + 1) % width === 0
  if (repeat) {
    height = selected.endRow - selected.startRow + 1
    width = selected.endColumn - selected.startColumn + 1
  }
  const result = {
    ...selected,
    endRow: selected.startRow + height - 1,
    endColumn: selected.startColumn + width - 1,
  }
  if (height * width > 2000 || result.endRow >= 1048576 || result.endColumn >= 16384)
    throw new Error(t('appClipboardUnsupported'))
  return result
}
function nativeCell(cell: Cell) {
  return {
    s: {},
    ...gridCell({
      ...cell,
      rawValue:
        cell.formula && cell.calculationState === 'ready'
          ? (cell.cachedValue ?? cell.rawValue)
          : cell.rawValue,
    }),
    ...(cell.kind === 'boolean' ? { v: cell.rawValue ? 1 : 0, t: CellValueType.BOOLEAN } : {}),
  }
}
export function specialPasteOperations(
  snapshot: SpecialPasteSnapshot,
  options: PasteSpecialOptions,
  target: { sheetId: string; address: string; cells: ViewCell[]; linkSheetName: string | null },
): WorkbookOperation[] {
  const availability = pasteSpecialAvailability(snapshot.source)
  if (!availability.types.has(options.type) || (options.link && !availability.link))
    throw new Error(t('appClipboardUnsupported'))
  const bounds = specialPasteBounds(snapshot, options, target.address),
    cells = sourceCells(snapshot)
  if (options.type === 'col-widths' && !options.link) {
    if (!snapshot.widths?.length) throw new Error(t('appClipboardUnsupported'))
    const operations: WorkbookOperation[] = []
    const columns = bounds.endColumn - bounds.startColumn + 1
    for (let c = 0; c < columns;) {
      const widthPx = snapshot.widths[c % snapshot.widths.length]!
      let count = 1
      while (
        c + count < columns &&
        count < 100 &&
        snapshot.widths[(c + count) % snapshot.widths.length] === widthPx
      )
        count++
      operations.push(
        workbookOperationSchema.parse({
          op: 'set_col_width',
          sheetId: target.sheetId,
          column: formatAddress(0, bounds.startColumn + c).replace(/1$/, ''),
          count,
          widthPx,
        }),
      )
      c += count
    }
    return operations
  }
  const index = new Map(target.cells.map((cell) => [cell.address, cell]))
  const carriesValues = !options.link && pasteTypeCarriesValues(options.type)
  const requiresValues =
    carriesValues &&
    (options.operation !== 'none' ||
      options.transpose ||
      options.type === 'values' ||
      options.type === 'values-numfmt')
  if (
    !options.link &&
    requiresValues &&
    cells.some((row) =>
      row.some(
        (cell) =>
          cell.formula && cell.calculationState !== 'ready' && cell.calculationState !== 'error',
      ),
    )
  )
    throw new Error(t('appClipboardUnsupported'))
  if (
    carriesValues &&
    options.operation !== 'none' &&
    target.cells.some(
      (cell) =>
        cell.formula && cell.calculationState !== 'ready' && cell.calculationState !== 'error',
    )
  )
    throw new Error(t('appClipboardUnsupported'))
  const source = cells.map((row, r) =>
    row.map((cell, c) => ({
      cell: nativeCell(cell),
      row: snapshot.block?.rows[r] ?? null,
      column: snapshot.block?.columns[c] ?? null,
    })),
  )
  const selected = parseRange(target.address)
  const built = buildPasteGrid(
    source,
    options,
    {
      startRow: bounds.startRow,
      startColumn: bounds.startColumn,
      rows: selected.endRow - selected.startRow + 1,
      columns: selected.endColumn - selected.startColumn + 1,
      cellAt: (r, c) => {
        const cell = index.get(formatAddress(r, c))
        return cell ? nativeCell(cell) : null
      },
    },
    offsetFormulaRefs,
    target.linkSheetName,
  )
  const donors = options.transpose && !options.link ? transposeGrid(cells) : cells
  const skipped: string[] = []
  const values = built.cells.map((row, r) =>
    row.map((data, c): Cell => {
      if (!data) {
        skipped.push(formatAddress(bounds.startRow + r, bounds.startColumn + c))
        return { kind: 'blank', rawValue: null }
      }
      const donor = donors[r % donors.length]?.[c % donors[0]!.length]
      const value = data.v ?? null
      const kind: Cell['kind'] =
        data.f || value === null
          ? 'blank'
          : data.t === CellValueType.BOOLEAN
            ? 'boolean'
            : typeof value === 'number'
              ? donor?.kind === 'date' && options.operation === 'none'
                ? 'date'
                : 'number'
              : typeof value === 'boolean'
                ? 'boolean'
                : (donor?.kind === 'error' && donor.rawValue === value) ||
                    (options.operation === 'divide' && value === '#DIV/0!')
                  ? 'error'
                  : 'string'
      return {
        kind,
        rawValue: data.f ? null : kind === 'boolean' ? !!value : value,
        ...(data.f ? { formula: data.f, calculationState: 'pending' } : {}),
        ...(kind === 'date' ? { dateSystem: donor!.dateSystem ?? '1900' } : {}),
        ...(donor?.style ? { style: donor.style } : {}),
        ...(donor?.numberFormat ? { numberFormat: donor.numberFormat } : {}),
      }
    }),
  )
  if (skipped.length === built.rows * built.columns) return []
  return [
    workbookOperationSchema.parse({
      op: 'paste_cells',
      sheetId: target.sheetId,
      start: formatAddress(bounds.startRow, bounds.startColumn),
      cells: values,
      skipCells: skipped,
      writeValues: options.link || options.type !== 'formats',
      format:
        snapshot.source.kind === 'text' ||
        options.link ||
        options.type === 'formulas' ||
        options.type === 'values'
          ? 'none'
          : options.type === 'formulas-numfmt' || options.type === 'values-numfmt'
            ? 'number'
            : options.type === 'all-except-borders'
              ? 'except-borders'
              : 'all',
    }),
  ]
}
