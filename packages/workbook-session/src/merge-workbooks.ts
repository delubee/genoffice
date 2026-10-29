import { columnLabel, formatAddress } from '@genoffice/xlsx-gateway/domain/cell-address'
import { dedupeSheetName } from '@genoffice/xlsx-gateway/domain/sheet-name-dedupe'
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import {
  measureNormalFontMdw,
  type NormalFontDescriptor,
} from '@genoffice/xlsx-gateway/domain/normal-font-metrics'
import { WorkbookError, assertNotCancelled, type CellRecord, type WorkbookState } from './domain'
import type { WorkbookSource } from './ports'
import { readCurrentRange } from './read'
import { NEW_SHEET_PREFIX, STAGED_WRITE_CELLS } from './staged-candidate'

/** An opened source workbook: its own initial root and private sidecar session. */
export interface MergeSource {
  readonly fileName: string
  readonly state: WorkbookState
  readonly sessionId: string
}

const BLOCK_CELLS = 2000
const OPS_PER_STAGE = 500

type ReadCell = CellRecord & { address: string; displayValue?: string }

/** Formulas import as their cached values, as the desktop merge does; dates follow the target system. */
function valueRecord(cell: ReadCell, dateSystem: WorkbookState['dateSystem']): CellRecord {
  const {
    address: _address,
    displayValue: _display,
    formula: _formula,
    cachedValue: _cached,
    calculationState: _state,
    ...record
  } = cell
  if (
    record.kind === 'date' &&
    typeof record.rawValue === 'number' &&
    record.dateSystem !== dateSystem
  )
    return {
      ...record,
      rawValue: record.rawValue + (record.dateSystem === '1904' ? 1462 : -1462),
      dateSystem,
    }
  return record
}

/** Excel character units to pixels at the target workbook's maximum digit width. */
function widthPx(characters: number, mdw: number) {
  if (characters === 0) return 0
  return Math.trunc(((256 * characters + Math.trunc(128 / mdw)) / 256) * mdw)
}

/**
 * Data → Merge Workbooks: every source sheet becomes a new, deduplicated sheet
 * with values, styles, merges, custom row heights and column widths, as bounded
 * stages for one staged transaction (see staged-candidate).
 */
export async function buildMergeStages(input: {
  target: WorkbookState
  sources: readonly MergeSource[]
  source: WorkbookSource
  signal?: AbortSignal
}) {
  const metadata = input.target.features?.originalMetadata as
    { workbookProtection?: { lockStructure?: boolean } } | undefined
  if (metadata?.workbookProtection?.lockStructure)
    throw new WorkbookError('protected_workbook_structure', 'The workbook structure is protected.')
  const cells = input.sources.reduce(
    (total, file) =>
      total + file.state.sheets.reduce((sum, sheet) => sum + sheet.rows * sheet.columns, 0),
    0,
  )
  if (cells > STAGED_WRITE_CELLS)
    throw new WorkbookError(
      'expanded_write_limit',
      `Merged sheets may contain at most ${STAGED_WRITE_CELLS} cells.`,
    )
  const mdw = measureNormalFontMdw(
    (input.target.features?.originalMetadata ?? {}) as NormalFontDescriptor,
  )
  const taken = new Set(input.target.sheets.map((sheet) => sheet.name))
  const additions: WorkbookOperation[] = [],
    blocks: { cells: number; op: WorkbookOperation }[] = [],
    layout: WorkbookOperation[] = [],
    sheetNames: string[] = []
  for (const file of input.sources)
    for (const sheet of file.state.sheets) {
      const name = dedupeSheetName(sheet.name, taken)
      taken.add(name)
      sheetNames.push(name)
      const sheetId = `${NEW_SHEET_PREFIX}${name}`
      additions.push({
        op: 'add_sheet',
        name,
        rows: Math.max(1000, sheet.rows + 100),
        columns: Math.max(26, sheet.columns + 5),
      })
      if (!sheet.rows || !sheet.columns) continue
      const width = Math.min(sheet.columns, BLOCK_CELLS),
        rowsPerBlock = Math.max(1, Math.floor(BLOCK_CELLS / width))
      let mergesRead = false
      for (let startColumn = 0; startColumn < sheet.columns; startColumn += width)
        for (let startRow = 0; startRow < sheet.rows; startRow += rowsPerBlock) {
          assertNotCancelled(input.signal)
          const endRow = Math.min(sheet.rows, startRow + rowsPerBlock) - 1,
            endColumn = Math.min(sheet.columns, startColumn + width) - 1
          const page = await readCurrentRange({
            state: file.state,
            sessionId: file.sessionId,
            source: input.source,
            logicalSheetId: sheet.id,
            range: `${formatAddress(startRow, startColumn)}:${formatAddress(endRow, endColumn)}`,
            includeMerges: true,
            signal: input.signal,
          })
          if (!('merges' in page)) throw new Error('A merge source page is missing its layout.')
          const byAddress = new Map(page.cells.map((cell) => [cell.address, cell as ReadCell]))
          let content = false
          const matrix: CellRecord[][] = []
          for (let row = startRow; row <= endRow; row++) {
            const line: CellRecord[] = []
            for (let column = startColumn; column <= endColumn; column++) {
              const cell = byAddress.get(formatAddress(row, column))
              const record = cell
                ? valueRecord(cell, input.target.dateSystem)
                : ({ kind: 'blank', rawValue: null } as CellRecord)
              if (record.rawValue !== null || record.style) content = true
              line.push(record)
            }
            matrix.push(line)
          }
          if (content)
            blocks.push({
              cells: matrix.length * matrix[0]!.length,
              op: {
                op: 'paste_cells',
                sheetId,
                start: formatAddress(startRow, startColumn),
                cells: matrix,
              },
            })
          if (startColumn === 0)
            for (const row of page.rows)
              if (row.customHeight && row.height !== undefined && !row.hidden)
                layout.push({
                  op: 'set_row_height',
                  sheetId,
                  row: row.row + 1,
                  count: 1,
                  heightPoints: Math.min(409.5, row.height),
                })
          if (!mergesRead) {
            mergesRead = true
            for (const merge of page.merges)
              layout.push({
                op: 'merge_cells',
                sheetId,
                range: `${formatAddress(merge.startRow, merge.startColumn)}:${formatAddress(merge.endRow, merge.endColumn)}`,
              })
          }
        }
      const widths = ((sheet.features?.originalMetadata as { columnWidths?: unknown } | undefined)
        ?.columnWidths ?? []) as {
        startColumn: number
        endColumn: number
        width?: number
        hidden: boolean
      }[]
      for (const span of widths) {
        if (span.hidden || span.width === undefined) continue
        const last = Math.min(span.endColumn, sheet.columns + 4)
        for (let column = span.startColumn; column <= last; column += 100)
          layout.push({
            op: 'set_col_width',
            sheetId,
            column: columnLabel(column),
            count: Math.min(100, last - column + 1),
            widthPx: widthPx(span.width, mdw),
            widthCharacters: span.width,
          })
      }
    }
  const stages: WorkbookOperation[][] = [additions]
  let stage: WorkbookOperation[] = [],
    budget = 0
  for (const block of blocks) {
    if (stage.length && budget + block.cells > BLOCK_CELLS) {
      stages.push(stage)
      stage = []
      budget = 0
    }
    stage.push(block.op)
    budget += block.cells
  }
  if (stage.length) stages.push(stage)
  for (let index = 0; index < layout.length; index += OPS_PER_STAGE)
    stages.push(layout.slice(index, index + OPS_PER_STAGE))
  return { stages, sheetNames, files: input.sources.length }
}
