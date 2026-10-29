import {
  formatAddress,
  parseAddress,
  parseRange,
  rangeAddresses,
  rangeCellCount,
} from './cell-address'
import { moveFormulaRefs, type FormulaMoveSpec } from './formula-shift'
import { sessionMerges } from './session-merges'
import type { WorkbookOperation } from './workbook-dsl'
import type { WorkbookSnapshot, CellState } from './workbook.types'
import type { cellRecordSchema } from '../shared/cell-record'
import type { z } from 'zod'

type Cell = z.infer<typeof cellRecordSchema>
type Area = ReturnType<typeof parseRange>
const intersects = (a: Area, b: Area) =>
  a.startRow <= b.endRow &&
  a.endRow >= b.startRow &&
  a.startColumn <= b.endColumn &&
  a.endColumn >= b.startColumn
const contains = (a: Area, b: Area) =>
  a.startRow <= b.startRow &&
  a.endRow >= b.endRow &&
  a.startColumn <= b.startColumn &&
  a.endColumn >= b.endColumn
const inArea = (address: string, area: Area) => {
  const p = parseAddress(address)
  return (
    p.row >= area.startRow &&
    p.row <= area.endRow &&
    p.column >= area.startColumn &&
    p.column <= area.endColumn
  )
}

/** The source is captured before clearing, so overlapping moves are atomic.
 * Formula dependents are read from complete workbook material, never a viewport. */
export function moveSessionRange(
  input: WorkbookSnapshot,
  op: Extract<WorkbookOperation, { op: 'relocate_range' }>,
  templates: Map<string, Record<string, Cell>>,
  features: Map<string, Record<string, unknown>>,
  writable: readonly string[],
  signal?: AbortSignal,
) {
  const book = { ...structuredClone(input), sheets: structuredClone([...input.sheets]) }
  const source = parseRange(op.source),
    start = parseAddress(op.start)
  const target: Area = {
    startRow: start.row,
    endRow: start.row + source.endRow - source.startRow,
    startColumn: start.column,
    endColumn: start.column + source.endColumn - source.startColumn,
  }
  const fromSheet = book.sheets.find((s) => s.id === op.sourceSheetId),
    toSheet = book.sheets.find((s) => s.id === op.sheetId)
  if (!fromSheet || !toSheet) throw new Error('unknown_sheet')
  if (!writable.includes(fromSheet.id) || !writable.includes(toSheet.id))
    throw new Error('protected_target')
  const spec: FormulaMoveSpec = {
    sourceSheetName: fromSheet.name,
    targetSheetName: toSheet.name,
    source,
    target,
  }
  const dr = target.startRow - source.startRow,
    dc = target.startColumn - source.startColumn
  const translate = (m: Area): Area => ({
    startRow: m.startRow + dr,
    endRow: m.endRow + dr,
    startColumn: m.startColumn + dc,
    endColumn: m.endColumn + dc,
  })
  const sourceFeatures = features.get(fromSheet.id)!,
    targetFeatures = features.get(toSheet.id)!
  const sourceMerges = sessionMerges(sourceFeatures),
    targetMerges = sessionMerges(targetFeatures)
  if (
    sourceMerges.some((m) => intersects(source, m) && !contains(source, m)) ||
    targetMerges.some((m) => intersects(target, m) && !contains(target, m))
  )
    throw new Error('merge_partial_overlap')
  for (const data of new Set([sourceFeatures, targetFeatures])) {
    const metadata = data.originalMetadata as Record<string, unknown> | undefined
    for (const key of ['tables', 'pivotRanges', 'pivotTables', 'sparklines'])
      if (Array.isArray(metadata?.[key]) && metadata[key].length)
        throw new Error(`unsupported_move_feature: ${key}`)
    for (const key of ['conditionalRules', 'dataValidations'])
      if (Array.isArray(data[key]) && (data[key] as unknown[]).length)
        throw new Error(`unsupported_move_feature: ${key}`)
    if (data.filterEdits || data.autoFilter) throw new Error('unsupported_move_feature: filter')
  }
  const fromTemplates = templates.get(fromSheet.id)!,
    toTemplates = templates.get(toSheet.id)!
  const donor = rangeAddresses(source).map((address) => ({
    address,
    cell: structuredClone(fromSheet.cells[address] ?? { value: null }),
    style: structuredClone(fromSheet.styles?.[address] ?? {}),
    typed: structuredClone(fromTemplates[address] ?? { kind: 'blank' as const, rawValue: null }),
  }))
  const affected = new Set([fromSheet.id, toSheet.id])
  const destinations = new Set<string>()
  const fromCells: Record<string, CellState> = { ...fromSheet.cells },
    fromStyles = { ...fromSheet.styles }
  const toCells = fromSheet === toSheet ? fromCells : { ...toSheet.cells },
    toStyles = fromSheet === toSheet ? fromStyles : { ...toSheet.styles }
  for (const entry of donor) {
    fromCells[entry.address] = { value: null }
    fromStyles[entry.address] = {}
    fromTemplates[entry.address] = { kind: 'blank', rawValue: null, styleReset: true, style: {} }
  }
  for (const entry of donor) {
    const point = parseAddress(entry.address),
      address = formatAddress(point.row + dr, point.column + dc)
    destinations.add(address)
    const formula = entry.cell.formula
      ? moveFormulaRefs(entry.cell.formula, spec, fromSheet.name, toSheet.name)
      : undefined
    toCells[address] = formula ? { value: null, formula } : entry.cell
    toStyles[address] = entry.style
    toTemplates[address] = { ...entry.typed, styleReset: true }
    if (formula) {
      toTemplates[address]!.kind = 'blank'
      toTemplates[address]!.rawValue = null
      toTemplates[address]!.formula = formula
      toTemplates[address]!.calculationState = 'pending'
      delete toTemplates[address]!.cachedValue
    }
  }
  let writes = rangeCellCount(source) + rangeCellCount(target)
  for (let index = 0; index < book.sheets.length; index++) {
    const sheet = book.sheets[index]!
    const cells = sheet === fromSheet ? fromCells : sheet === toSheet ? toCells : { ...sheet.cells }
    const styles = sheet === fromSheet ? fromStyles : sheet === toSheet ? toStyles : sheet.styles
    for (const [address, cell] of Object.entries(cells)) {
      if (signal?.aborted) throw new Error('operation_cancelled')
      if (!cell.formula || (sheet === toSheet && destinations.has(address))) continue
      const formula = moveFormulaRefs(cell.formula, spec, sheet.name)
      if (formula === cell.formula) continue
      if (!writable.includes(sheet.id)) throw new Error('protected_formula_dependency')
      cells[address] = { value: null, formula }
      writes++
      affected.add(sheet.id)
    }
    for (const visual of sheet.visuals ?? [])
      if (visual.kind === 'chart')
        for (const series of visual.chart.series)
          for (const key of ['valuesRef', 'categoriesRef', 'nameRef'] as const) {
            if (series[key]) {
              const reference = moveFormulaRefs(series[key]!, spec, sheet.name)
              if (reference !== series[key]) {
                if (!writable.includes(sheet.id)) throw new Error('protected_formula_dependency')
                affected.add(sheet.id)
                series[key] = reference
              }
            }
          }
    book.sheets[index] = { ...sheet, cells, ...(styles ? { styles } : {}) }
  }
  if (writes > 2000) throw new Error('expanded_write_limit: ordinary batches allow 2000 cells')
  sourceFeatures.mergeEdits = sourceMerges.filter(
    (m) => !intersects(source, m) && !(fromSheet === toSheet && intersects(target, m)),
  )
  targetFeatures.mergeEdits = [
    ...(fromSheet === toSheet
      ? (sourceFeatures.mergeEdits as Area[])
      : targetMerges.filter((m) => !intersects(target, m))),
    ...sourceMerges.filter((m) => contains(source, m)).map(translate),
  ]
  // Notes and hyperlinks follow cells; retain explicit null tombstones on source.
  for (const [key, originalKey] of [
    ['noteEdits', 'comments'],
    ['hyperlinkEdits', 'hyperlinks'],
  ] as const) {
    const records = (data: Record<string, unknown>) => {
      const metadata = data.originalMetadata as Record<string, unknown> | undefined
      const original = (data[originalKey] ?? metadata?.[originalKey] ?? []) as {
        row: number
        column: number
        text?: string
        target?: string
        image?: unknown
        thread?: unknown
      }[]
      const result: Record<string, string | null> = {}
      for (const item of original) {
        const address = formatAddress(item.row, item.column)
        if (
          ((data === sourceFeatures && inArea(address, source)) ||
            (data === targetFeatures && inArea(address, target))) &&
          (item.image || item.thread)
        )
          throw new Error('unsupported_move_feature: rich_note')
        result[address] = (key === 'noteEdits' ? item.text : item.target) ?? null
      }
      return { ...result, ...((data[key] ?? {}) as Record<string, string | null>) }
    }
    const sourceRecords = records(sourceFeatures),
      targetRecords = records(targetFeatures)
    const sourceEdits = { ...((sourceFeatures[key] ?? {}) as Record<string, string | null>) },
      targetEdits =
        fromSheet === toSheet
          ? sourceEdits
          : { ...((targetFeatures[key] ?? {}) as Record<string, string | null>) }
    for (const address of rangeAddresses(source))
      if (sourceRecords[address] != null) sourceEdits[address] = null
    for (const entry of donor) {
      const p = parseAddress(entry.address),
        address = formatAddress(p.row + dr, p.column + dc)
      if (sourceRecords[entry.address] != null || targetRecords[address] != null)
        targetEdits[address] = sourceRecords[entry.address] ?? null
    }
    if (Object.keys(sourceEdits).length) sourceFeatures[key] = sourceEdits
    if (Object.keys(targetEdits).length) targetFeatures[key] = targetEdits
  }
  return { book, target, writes, spec, affected: [...affected] }
}
