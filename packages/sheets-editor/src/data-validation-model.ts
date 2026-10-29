import { DataValidationRenderMode } from '@univerjs/core'
import type { WorkbookRangeResult } from '@genoffice/xlsx-gateway/shared/workbook-protocol'

export interface DvListSourceContext {
  targetSheetId: string
  definedNames: { name: string; formula: string; scopeSheetId?: string }[]
  sheets: {
    id: string
    name: string
    tables: {
      name?: string
      range: { startRow: number; endRow: number; startColumn: number; endColumn: number }
      headerRowCount: number
      totalsRowCount?: number
      columns?: string[]
    }[]
  }[]
}

type WorkbookDvRule = WorkbookRangeResult['dataValidations'][number]

const errorStyles: Record<string, number> = { stop: 1, warning: 2, information: 0 }

/** Convert the workbook protocol's data validation into Univer's live grid rule. */
export function workbookDvRuleToUniver(
  rule: WorkbookDvRule,
  uid: string,
  context?: DvListSourceContext,
): Record<string, unknown> | null {
  const type = rule.ruleType === 'none' ? 'any' : rule.ruleType
  if (!['any', 'whole', 'decimal', 'list', 'date', 'time', 'textLength', 'custom'].includes(type))
    return null

  let formula1 = rule.formulas[0]
  const formula2 = rule.formulas[1]
  if (type === 'list' && formula1 !== undefined) {
    const literal = formula1.trim()
    if (literal === '"1,0"') {
      return {
        uid,
        type: 'checkbox',
        ranges: rule.ranges,
        allowBlank: rule.allowBlank,
      }
    }
    formula1 =
      literal.startsWith('"') && literal.endsWith('"')
        ? literal
            .slice(1, -1)
            .split(',')
            .map((item) => item.trim().replaceAll('""', '"'))
            .join(',')
        : `=${resolveDvListSource(literal.replace(/^=/, ''), context)}`
  } else if (type === 'custom' && formula1 !== undefined) {
    formula1 = `=${formula1.replace(/^=/, '')}`
  }

  return {
    uid,
    type,
    ranges: rule.ranges,
    allowBlank: rule.allowBlank,
    ...(rule.operator === undefined ? {} : { operator: rule.operator }),
    ...(formula1 === undefined ? {} : { formula1 }),
    ...(formula2 === undefined || type === 'list' ? {} : { formula2 }),
    ...(type === 'list'
      ? { showDropDown: !rule.suppressDropdown, renderMode: DataValidationRenderMode.TEXT }
      : {}),
    ...(rule.showInputMessage ? { showInputMessage: true } : {}),
    ...(rule.showErrorMessage ? { showErrorMessage: true } : {}),
    ...(rule.errorStyle === undefined || errorStyles[rule.errorStyle] === undefined
      ? {}
      : { errorStyle: errorStyles[rule.errorStyle] }),
    ...(rule.errorTitle === undefined ? {} : { errorTitle: rule.errorTitle }),
    ...(rule.error === undefined ? {} : { error: rule.error }),
    ...(rule.promptTitle === undefined ? {} : { promptTitle: rule.promptTitle }),
    ...(rule.prompt === undefined ? {} : { prompt: rule.prompt }),
  }
}

/** Resolve Excel names and common structured table-column refs for Univer's list validator. */
export function resolveDvListSource(source: string, context?: DvListSourceContext): string {
  if (!context) return source
  const visit = (value: string, seen: Set<string>, depth: number): string => {
    if (depth > 8) return source
    const target = value.trim()
    const tableOpen = target.indexOf('[')
    if (tableOpen > 0 && target.endsWith(']')) {
      const tableName = target.slice(0, tableOpen).trim()
      const selectors = [...target.slice(tableOpen).matchAll(/\[([^\]]+)\]/g)].map((match) =>
        match[1]!.trim().replace(/^\[/, ''),
      )
      const columnName = selectors.filter((selector) => !selector.startsWith('#')).at(-1)
      if (columnName) {
        for (const sheet of context.sheets) {
          const table = sheet.tables.find(
            (item) => item.name?.toLowerCase() === tableName.toLowerCase(),
          )
          const column =
            table?.columns?.findIndex((name) => name.toLowerCase() === columnName.toLowerCase()) ??
            -1
          if (!table || column < 0) continue
          const totalsRowCount = table.totalsRowCount ?? 0
          const special = selectors.find((selector) => selector.startsWith('#'))?.toLowerCase()
          const startRow =
            special === '#all' || special === '#headers'
              ? table.range.startRow
              : special === '#totals'
                ? table.range.endRow - totalsRowCount + 1
                : table.range.startRow + table.headerRowCount
          const endRow =
            special === '#all'
              ? table.range.endRow
              : special === '#headers'
                ? table.range.startRow + table.headerRowCount - 1
                : special === '#totals'
                  ? table.range.endRow
                  : table.range.endRow - totalsRowCount
          if (startRow > endRow) return source
          const columnLabel = columnToLabel(table.range.startColumn + column)
          return `${quoteSheetName(sheet.name)}!$${columnLabel}$${startRow + 1}:$${columnLabel}$${endRow + 1}`
        }
      }
    }

    const qualifier = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(target)
    const requestedSheet = qualifier
      ? (qualifier[1] ?? qualifier[2]!).replaceAll("''", "'")
      : undefined
    const name = qualifier ? qualifier[3]!.trim() : target
    if (!/^[A-Za-z_\\][\w.\\]*$/.test(name)) return source
    const matches = context.definedNames.filter(
      (entry) => entry.name.toLowerCase() === name.toLowerCase(),
    )
    const requestedSheetId = requestedSheet
      ? context.sheets.find((sheet) => sheet.name.toLowerCase() === requestedSheet.toLowerCase())
          ?.id
      : context.targetSheetId
    const definition =
      matches.find((entry) => entry.scopeSheetId === requestedSheetId) ??
      matches.find((entry) => entry.scopeSheetId === undefined)
    if (!definition) return source
    const key = `${definition.scopeSheetId ?? ''}:${definition.name.toLowerCase()}`
    if (seen.has(key)) return source
    seen.add(key)
    let formula = definition.formula.trim().replace(/^=/, '')
    if (!formula.includes('!') && definition.scopeSheetId) {
      const owner = context.sheets.find((sheet) => sheet.id === definition.scopeSheetId)
      if (owner) formula = `${quoteSheetName(owner.name)}!${formula}`
    }
    return visit(formula, seen, depth + 1)
  }
  return visit(source, new Set(), 0)
}

function quoteSheetName(name: string): string {
  return `'${name.replaceAll("'", "''")}'`
}

function columnToLabel(column: number): string {
  let value = column + 1
  let label = ''
  while (value > 0) {
    const remainder = (value - 1) % 26
    label = String.fromCharCode(65 + remainder) + label
    value = Math.floor((value - 1) / 26)
  }
  return label
}
