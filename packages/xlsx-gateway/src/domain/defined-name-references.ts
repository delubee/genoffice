import { builtinFormulaFunctions } from './formula-builtins'
import { formulaVariableOffsets } from './formula-bindings'
export interface NameReferenceChange {
  name: string
  replacement: string | null
  scopeSheetId?: string | undefined
  shadowedSheetIds: string[]
}
export interface NameReferencePlan {
  changes: NameReferenceChange[]
  sheets: { id: string; name: string }[]
}
const identifier = /[\p{L}_\\][\p{L}\p{N}_.\\]*/uy
const error = /^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|SPILL!|CALC!|GETTING_DATA)/i
/** Apply action-time bindings, preserving literals, table tokens and external references. */
export function rewriteNameReferences(
  formula: string,
  ownerSheetId: string,
  plan: NameReferencePlan,
): string {
  for (const change of plan.changes) {
    const variables = formulaVariableOffsets(formula)
    let output = '',
      at = 0
    const word = (start: number) => {
      identifier.lastIndex = start
      return identifier.exec(formula)?.[0]
    }
    while (at < formula.length) {
      const start = at
      if (formula[at] === '"') {
        at++
        while (at < formula.length) {
          if (formula[at++] !== '"') continue
          if (formula[at] === '"') {
            at++
            continue
          }
          break
        }
        output += formula.slice(start, at)
        continue
      }
      if (formula[at] === '[') {
        let depth = 0
        do {
          const c = formula[at++]
          if (c === '[') depth++
          else if (c === ']') depth--
        } while (at < formula.length && depth)
        output += formula.slice(start, at)
        continue
      }
      const literalError = formula[at] === '#' ? error.exec(formula.slice(at))?.[0] : undefined
      if (literalError) {
        output += literalError
        at += literalError.length
        continue
      }
      let qualifier: string | undefined, token: string | undefined
      if (formula[at] === "'") {
        at++
        while (at < formula.length) {
          if (formula[at++] !== "'") continue
          if (formula[at] === "'") {
            at++
            continue
          }
          break
        }
        if (formula[at] === '!') {
          qualifier = formula.slice(start, at)
          at++
          token = word(at)
          if (token) at += token.length
        }
      } else {
        token = word(at)
        if (token) {
          at += token.length
          if (formula[at] === '!') {
            qualifier = token
            at++
            token = word(at)
            if (token) at += token.length
          }
        } else at++
      }
      const before = formula[start - 1] ?? ''
      const qualifiedSheet = qualifier?.startsWith("'")
        ? qualifier.slice(1, -1).replace(/''/g, "'")
        : qualifier
      const scope =
        qualifiedSheet === undefined
          ? ownerSheetId
          : plan.sheets.find((sheet) => sheet.name.toLowerCase() === qualifiedSheet.toLowerCase())
              ?.id
      const bound =
        change.scopeSheetId === undefined
          ? scope !== undefined && !change.shadowedSheetIds.includes(scope)
          : scope === change.scopeSheetId
      const eligible =
        token &&
        token.toLowerCase() === change.name.toLowerCase() &&
        !variables.has(start) &&
        !(
          qualifier === undefined &&
          /^\s*\(/.test(formula.slice(at)) &&
          builtinFormulaFunctions.has(token.toUpperCase())
        ) &&
        bound &&
        !/[\p{L}\p{N}_.\\\]#!]/u.test(before) &&
        formula[at] !== '[' &&
        !qualifiedSheet?.includes('[')
      output += eligible
        ? change.replacement === null
          ? '#REF!'
          : (qualifier === undefined ? '' : qualifier + '!') + change.replacement
        : formula.slice(start, at)
    }
    formula = output
  }
  return formula
}
