import type { WorkbookRangeResult } from '../shared/workbook-protocol'
import {
  subtractDvAreas,
  nativeDvListValues,
  type DvWireRule,
  type DvCellArea,
} from '../gateway/xlsx-dv'
import { offsetFormulaRefs } from './formula-shift'

type ReadRule = WorkbookRangeResult['dataValidations'][number]
export interface DataValidationEdits {
  rules: DvWireRule[]
  remove: DvCellArea[]
}

/** Complete source rules plus the canonical pending overlay, in current coordinates. */
export function projectDataValidationRules(
  source: readonly ReadRule[],
  edits?: DataValidationEdits,
): ReadRule[] {
  if (!edits) return structuredClone(source) as ReadRule[]
  const cuts = [...edits.remove, ...edits.rules.flatMap((rule) => rule.ranges)]
  const kept = source.flatMap((rule) =>
    rule.ranges.flatMap((original) =>
      subtractDvAreas(original, cuts).map((range) => ({
        ...structuredClone(rule),
        ranges: [range],
        formulas: rule.formulas.map((formula) =>
          rule.ruleType === 'list' && formula.startsWith('"')
            ? formula
            : offsetFormulaRefs(
                formula,
                range.startRow - original.startRow,
                range.startColumn - original.startColumn,
              ),
        ),
      })),
    ),
  )
  const added = edits.rules.map(({ ranges, rule }): ReadRule => {
    let type = String(rule.type)
    const formulas = [rule.formula1, rule.formula2].filter(
      (value): value is string => typeof value === 'string' && value !== '',
    )
    if (type === 'checkbox') {
      type = 'list'
      formulas.splice(0, formulas.length, '"1,0"')
    } else if (type === 'list' && formulas[0] !== undefined) {
      formulas[0] = formulas[0].startsWith('=')
        ? formulas[0].slice(1)
        : `"${nativeDvListValues(formulas[0]).join(',').replaceAll('"', '""')}"`
    }
    return {
      ranges: structuredClone([...ranges]),
      ruleType: type === 'any' ? 'none' : type,
      formulas,
      allowBlank: rule.allowBlank === true,
      suppressDropdown: rule.showDropDown === false,
      showInputMessage: rule.showInputMessage === true,
      showErrorMessage: rule.showErrorMessage === true,
      ...(typeof rule.operator === 'string' ? { operator: rule.operator } : {}),
      ...Object.fromEntries(
        ['errorTitle', 'error', 'promptTitle', 'prompt'].flatMap((key) =>
          typeof rule[key] === 'string' ? [[key, rule[key]]] : [],
        ),
      ),
      ...([0, 1, 2].includes(Number(rule.errorStyle))
        ? { errorStyle: ['information', 'stop', 'warning'][Number(rule.errorStyle)]! }
        : {}),
    }
  })
  return [...kept, ...added]
}
