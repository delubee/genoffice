import { randomUUID } from 'node:crypto'
import type { IWorkbookData, ICellData } from '@univerjs/core'
import {
  LexerTreeBuilder,
  ALL_IMPLEMENTED_FUNCTIONS_SET,
  sequenceNodeType,
  type IDefinedNamesServiceParam,
} from '@univerjs/engine-formula'
import { isValidDefinedName } from '@genoffice/xlsx-gateway/gateway/xlsx-defined-names'

export class CalculationBindingError extends Error {
  constructor(readonly code: string) {
    super(code)
    this.name = 'CalculationBindingError'
  }
}

export interface CalculationNameDefinition {
  name: string
  formula: string
  /** Logical sheet ID; source scope IDs must be mapped before this boundary. */
  sheetId?: string
  hidden?: boolean
}

/** Give each resolved name a private engine identifier. The engine caches names
 * and lexer trees by spelling, so two scopes must never share an identifier.
 * Only the private computation copy is rewritten; authored text stays separate.
 */
export function bindCalculationNames(
  snapshot: Partial<IWorkbookData>,
  names: readonly CalculationNameDefinition[],
) {
  const sheets = new Map(
    Object.entries(snapshot.sheets ?? {}).map(([id, sheet]) => [id, sheet.name ?? id]),
  )
  const byScope = new Map<string, CalculationNameDefinition>()
  const scopeKey = (scope: string | undefined, name: string) =>
    `${scope ?? ''}\0${name.toLowerCase()}`
  for (const raw of names) {
    const name = { ...raw }
    if (!isValidDefinedName(name.name) || (name.sheetId !== undefined && !sheets.has(name.sheetId)))
      throw new CalculationBindingError('calculation_invalid_name')
    const key = scopeKey(name.sheetId, name.name)
    if (byScope.has(key)) throw new CalculationBindingError('calculation_ambiguous_name')
    byScope.set(key, name)
  }
  const tokenizer = new LexerTreeBuilder()
  const aliases = new Map<string, IDefinedNamesServiceParam>()
  const building = new Set<string>()
  const prefix = `_genoffice_${randomUUID().replaceAll('-', '')}_`
  type Token = NonNullable<ReturnType<LexerTreeBuilder['sequenceNodesBuilder']>>[number]
  const text = (token: Token) => (typeof token === 'string' ? token : token.token)
  const unary = (raw: string) => {
    const match = /^([-+@]*)([\s\S]*)$/.exec(raw)!
    return { prefix: match[1]!, body: match[2]! }
  }
  const resolve = (token: string, owner: string, workbookOnly = false) => {
    let scope = owner,
      name = token
    const match = /^(?:'((?:[^']|'')+)'|([^!]+))!([^!]+)$/.exec(token)
    if (match) {
      const qualifier = (match[1]?.replaceAll("''", "'") ?? match[2])!
      if (qualifier.includes('[')) return undefined
      const id = [...sheets].find(([, name]) => name.toLowerCase() === qualifier.toLowerCase())?.[0]
      if (!id)
        return qualifier.toLowerCase() === snapshot.name?.toLowerCase()
          ? byScope.get(scopeKey(undefined, match[3]!))
          : undefined
      scope = id
      name = match[3]!
    }
    return (
      (!match && workbookOnly ? undefined : byScope.get(scopeKey(scope, name))) ??
      byScope.get(scopeKey(undefined, name))
    )
  }
  const alias = (definition: CalculationNameDefinition, owner: string): string => {
    const evaluationSheet = definition.sheetId ?? owner
    const key = scopeKey(definition.sheetId, definition.name) + '\0' + evaluationSheet
    if (building.has(key)) throw new CalculationBindingError('calculation_circular_name')
    const existing = aliases.get(key)
    if (existing) return existing.name
    if (aliases.size >= 100000) throw new CalculationBindingError('calculation_name_limit')
    const generated = prefix + aliases.size
    const entry: IDefinedNamesServiceParam = {
      id: generated,
      name: generated,
      formulaOrRefString: '',
      hidden: true,
      localSheetId: 'AllDefaultWorkbook',
    }
    aliases.set(key, entry)
    building.add(key)
    // Parentheses retain the precedence of a named expression when substituted.
    const body = rewrite(
      definition.formula || '#NAME?',
      evaluationSheet,
      definition.sheetId !== undefined,
      definition.sheetId === undefined,
    ).replace(/^=/, '')
    // Univer's inline LAMBDA call syntax requires LAMBDA(...)(args), rather
    // than (LAMBDA(...))(args). Other named expressions need grouping.
    entry.formulaOrRefString = /^(?:_xlfn\.)?LAMBDA\s*\(/i.test(body.trim()) ? body : `(${body})`
    building.delete(key)
    return generated
  }
  const rewrite = (formula: string, owner: string, qualifyRefs = false, workbookOnly = false) => {
    const tokens = tokenizer.sequenceNodesBuilder(formula)
    if (!tokens) throw new CalculationBindingError('calculation_name_syntax')
    const close = (open: number, end: number) => {
      let depth = 1
      for (let at = open + 1; at < end; at++) {
        if (text(tokens[at]!) === '(') depth++
        if (text(tokens[at]!) === ')' && --depth === 0) return at
      }
      throw new CalculationBindingError('calculation_name_syntax')
    }
    const argumentsOf = (start: number, end: number) => {
      const args: [number, number][] = []
      let depth = 0,
        from = start
      for (let at = start; at < end; at++) {
        const token = text(tokens[at]!)
        if (token === '(') depth++
        else if (token === ')') depth--
        else if (token === ',' && depth === 0) {
          args.push([from, at])
          from = at + 1
        }
      }
      args.push([from, end])
      return args
    }
    const walk = (start: number, end: number, bound: Set<string>): string => {
      let output = ''
      for (let at = start; at < end; at++) {
        const token = tokens[at]!,
          raw = text(token)
        if (typeof token !== 'string' && text(tokens[at + 1] ?? '') === '(') {
          const last = close(at + 1, end),
            args = argumentsOf(at + 2, last)
          const { prefix, body } = unary(raw.trim())
          const functionName = body.replace(/^_xlfn\./i, '').toUpperCase()
          const env = new Set(bound)
          const parts: string[] = []
          for (let index = 0; index < args.length; index++) {
            const [from, to] = args[index]!
            const declaration =
              (functionName === 'LET' && index < args.length - 1 && index % 2 === 0) ||
              (functionName === 'LAMBDA' && index < args.length - 1)
            if (declaration) {
              const spelling = tokens.slice(from, to).map(text).join('')
              parts.push(spelling)
              // LET's value is evaluated before its newly declared variable.
              if (functionName === 'LAMBDA') env.add(spelling.trim().toLowerCase())
            } else {
              parts.push(walk(from, to, env))
              if (functionName === 'LET' && index % 2 === 1) {
                const [a, b] = args[index - 1]!
                env.add(tokens.slice(a, b).map(text).join('').trim().toLowerCase())
              }
            }
          }
          const callable =
            !bound.has(body.toLowerCase()) && !ALL_IMPLEMENTED_FUNCTIONS_SET.has(functionName)
              ? resolve(body, owner, workbookOnly)
              : undefined
          const call = (callable ? alias(callable, owner) : raw) + '(' + parts.join(',') + ')'
          output += callable && prefix ? `${prefix}(${call})` : call
          at = last
          continue
        }
        if (raw === '(') {
          const last = close(at, end)
          output += '(' + walk(at + 1, last, bound) + ')'
          at = last
          continue
        }
        if (
          typeof token !== 'string' &&
          (token.nodeType === sequenceNodeType.FUNCTION ||
            token.nodeType === sequenceNodeType.REFERENCE) &&
          !raw.includes('[')
        ) {
          const { prefix, body } = unary(raw)
          const definition = bound.has(body.toLowerCase())
            ? undefined
            : resolve(body, owner, workbookOnly)
          if (definition) {
            const name = alias(definition, owner)
            output += prefix ? `${prefix}(${name})` : name
            continue
          }
          if (qualifyRefs && token.nodeType === sequenceNodeType.REFERENCE && !body.includes('!')) {
            output += `${prefix}'${sheets.get(owner)!.replaceAll("'", "''")}'!${body}`
            continue
          }
        }
        output += raw
      }
      return output
    }
    return '=' + walk(0, tokens.length, new Set())
  }
  try {
    for (const [id, sheet] of Object.entries(snapshot.sheets ?? {}))
      for (const row of Object.values(sheet.cellData ?? {}))
        for (const raw of Object.values(row ?? {})) {
          const cell = raw as ICellData | null | undefined
          if (cell?.f) cell.f = rewrite(cell.f, id)
        }
    return { definitions: [...aliases.values()], resolve }
  } finally {
    tokenizer.dispose()
  }
}
