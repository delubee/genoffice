import { CellValueType, LocaleType, type ICellData, type IWorkbookData } from '@univerjs/core'
import { WorkbookError } from '../domain'
import type { CalculationSourceData } from './source'
import type { CalculationTable } from './tables'
import type { CalculationNameDefinition } from './names'

export interface CalculationSnapshot {
  workbook: Partial<IWorkbookData>
  definedNames: CalculationNameDefinition[]
  tables: CalculationTable[]
}

function withoutFormulaStrings(formula: string) {
  const chars = formula.split('')
  for (let index = 0; index < chars.length; index++) {
    if (chars[index] !== '"') continue
    chars[index] = ' '
    for (index++; index < chars.length; index++) {
      if (chars[index] !== '"') chars[index] = ' '
      else if (chars[index + 1] === '"') {
        chars[index] = ' '
        chars[++index] = ' '
      } else {
        chars[index] = ' '
        break
      }
    }
  }
  return chars.join('')
}

/** Convert one frozen host source into the normalized formula-engine input. */
export function createCalculationSnapshot(source: CalculationSourceData): CalculationSnapshot {
  if (source.dateSystem === '1904')
    throw new WorkbookError(
      'calculation_unavailable',
      'The calculation engine does not preserve the workbook 1904 date system.',
    )
  const sheets: NonNullable<IWorkbookData['sheets']> = {}
  for (const sheet of source.sheets) {
    if (sheet.sourceArrays.length)
      throw new WorkbookError(
        'calculation_unavailable',
        'Legacy array-formula ownership is not yet bound to the calculation engine.',
      )
    const cellData: NonNullable<IWorkbookData['sheets'][string]>['cellData'] = {}
    for (const [address, cell] of Object.entries(sheet.cells)) {
      const row = /^([A-Z]+)([1-9][0-9]*)$/.exec(address)
      if (!row)
        throw new WorkbookError('invalid_document', 'A calculation cell address is invalid.')
      let column = 0
      for (const char of row[1]!) column = column * 26 + char.charCodeAt(0) - 64
      const rowIndex = Number(row[2]) - 1
      const columnIndex = column - 1
      const value: ICellData = {}
      if (cell.formula) value.f = cell.formula
      else if (cell.kind === 'error')
        throw new WorkbookError(
          'calculation_unavailable',
          'Literal source error values are not representable in the calculation snapshot.',
        )
      else if (typeof cell.rawValue === 'string') {
        value.v = cell.rawValue
        value.t = CellValueType.FORCE_STRING
      } else value.v = cell.rawValue
      if (cell.numberFormat) value.s = { n: { pattern: cell.numberFormat } }
      cellData[rowIndex] ??= {}
      cellData[rowIndex]![columnIndex] = value
    }
    sheets[sheet.id] = {
      id: sheet.id,
      name: sheet.name,
      rowCount: sheet.rows,
      columnCount: sheet.columns,
      cellData,
    }
  }
  const formulas = [
    ...source.sheets.flatMap((sheet) =>
      Object.values(sheet.cells).flatMap((cell) => (cell.formula ? [cell.formula] : [])),
    ),
    ...source.definedNames.map((definition) => definition.formula),
  ]
  if (formulas.some((formula) => /\[[^\]]+\][^!]*!/u.test(withoutFormulaStrings(formula))))
    throw new WorkbookError(
      'calculation_unavailable',
      'Formulas referencing external workbooks need their linked source resources.',
    )
  return {
    workbook: {
      id: source.stateId,
      name: 'Workbook',
      locale: LocaleType.EN_US,
      sheetOrder: source.sheets.map((sheet) => sheet.id),
      sheets,
    },
    definedNames: structuredClone(source.definedNames),
    tables: structuredClone(source.tables),
  }
}
