import { rewriteNameReferences } from './defined-name-references'
import { isValidDefinedName } from '../gateway/xlsx-defined-names'
import type { CandidateState } from '../shared/session-changeset'

export interface SessionDefinedName {
  name: string
  formula: string
  scopeSheetId?: string | undefined
}
export interface SessionDefinedNames {
  names: SessionDefinedName[]
  preserveNames: string[]
}
export function sessionDefinedNames(
  state: Pick<CandidateState, 'features' | 'sheets'>,
): SessionDefinedNames {
  const journal = state.features?.definedNamesEdits as unknown as SessionDefinedNames | undefined
  if (journal) return structuredClone(journal)
  const original = state.features?.originalMetadata as
    | {
        definedNames?: { name: string; formula: string; sheetIndex?: number }[]
        sourceSheetIds?: string[]
      }
    | undefined
  const names: SessionDefinedName[] = [],
    preserveNames: string[] = []
  for (const entry of original?.definedNames ?? []) {
    if (entry.name.toLowerCase().startsWith('_xlnm')) continue
    const sourceId =
      entry.sheetIndex === undefined ? undefined : original?.sourceSheetIds?.[entry.sheetIndex]
    const sheet =
      entry.sheetIndex === undefined
        ? undefined
        : original?.sourceSheetIds
          ? sourceId === undefined
            ? undefined
            : state.sheets.find((sheet) => sheet.originalSheetId === sourceId)
          : state.sheets[entry.sheetIndex]
    if (entry.sheetIndex !== undefined && sourceId !== undefined && !sheet) continue
    if (
      !isValidDefinedName(entry.name) ||
      entry.formula.length > 8192 ||
      (entry.sheetIndex !== undefined && !sheet)
    ) {
      preserveNames.push(entry.name)
      continue
    }
    names.push({
      name: entry.name,
      formula: entry.formula.replace(/^=/, ''),
      ...(sheet ? { scopeSheetId: sheet.id } : {}),
    })
  }
  const count = new Map<string, number>()
  const key = (entry: SessionDefinedName) =>
    `${entry.scopeSheetId ?? ''}\u0000${entry.name.toLowerCase()}`
  for (const entry of names) count.set(key(entry), (count.get(key(entry)) ?? 0) + 1)
  const modeled = names.filter((entry) => {
    if (count.get(key(entry)) === 1) return true
    preserveNames.push(entry.name)
    return false
  })
  return { names: modeled, preserveNames: [...new Set(preserveNames)] }
}
export function sourceTableNames(state: Pick<CandidateState, 'sheets'>): string[] {
  return state.sheets.flatMap((sheet) =>
    (
      (sheet.features?.originalMetadata as { tables?: { name?: string }[] } | undefined)?.tables ??
      []
    ).flatMap((table) => (table.name ? [table.name] : [])),
  )
}
export function assertNameBindingsComplete(state: Pick<CandidateState, 'features'>) {
  if (
    (state.features?.originalMetadata as { nameBindingsComplete?: boolean } | undefined)
      ?.nameBindingsComplete !== true
  )
    throw new Error('defined_name_bindings_incomplete')
}
function reservedNameConflict(
  state: Pick<CandidateState, 'features' | 'sheets'>,
  name: string,
  scope?: string,
) {
  const entries =
    (
      state.features?.originalMetadata as
        | { reservedDefinedNames?: { name: string; sheetIndex?: number; sheetId?: string }[] }
        | undefined
    )?.reservedDefinedNames ?? []
  return entries.some((entry) => {
    if (entry.name.toLowerCase() !== name.toLowerCase()) return false
    if (entry.sheetIndex === undefined) return scope === undefined
    if (!entry.sheetId) return true
    const owner = state.sheets.find((sheet) => sheet.originalSheetId === entry.sheetId)
    return owner !== undefined && owner.id === scope
  })
}
function localBindingScopes(
  state: Pick<CandidateState, 'features' | 'sheets'>,
  names: SessionDefinedNames,
  name: string,
): string[] {
  const modeled = names.names.flatMap((entry) =>
    entry.name.toLowerCase() === name.toLowerCase() && entry.scopeSheetId
      ? [entry.scopeSheetId]
      : [],
  )
  const reserved =
    (
      state.features?.originalMetadata as
        | { reservedDefinedNames?: { name: string; sheetId?: string; sheetIndex?: number }[] }
        | undefined
    )?.reservedDefinedNames ?? []
  const hidden = reserved.flatMap((entry) => {
    if (entry.name.toLowerCase() !== name.toLowerCase() || entry.sheetIndex === undefined) return []
    const owner = state.sheets.find((sheet) => sheet.originalSheetId === entry.sheetId)
    return owner ? [owner.id] : []
  })
  return [...new Set([...modeled, ...hidden])]
}
export function editSessionDefinedName(
  state: Pick<CandidateState, 'features' | 'sheets'>,
  op:
    | { op: 'add_defined_name'; name: string; ref: string; scopeSheetId?: string | undefined }
    | { op: 'delete_defined_name'; name: string; scopeSheetId?: string | undefined },
): SessionDefinedNames {
  const protection = (
    state.features?.originalMetadata as
      { workbookProtection?: { protected?: boolean; lockStructure?: boolean } } | undefined
  )?.workbookProtection
  if (protection?.protected || protection?.lockStructure) throw new Error('protected_workbook')
  if (op.scopeSheetId && !state.sheets.some((sheet) => sheet.id === op.scopeSheetId))
    throw new Error('unknown_name_scope')
  if (!isValidDefinedName(op.name) || op.name.toLowerCase().startsWith('_xlnm'))
    throw new Error('invalid_defined_name')
  const current = sessionDefinedNames(state)
  const matches = current.names.filter(
    (name) =>
      name.name.toLowerCase() === op.name.toLowerCase() && name.scopeSheetId === op.scopeSheetId,
  )
  if (op.op === 'delete_defined_name') {
    if (matches.length !== 1) throw new Error('defined_name_missing_or_ambiguous')
    current.names = current.names.filter((name) => name !== matches[0])
  } else {
    if (
      matches.length ||
      (op.scopeSheetId === undefined &&
        sourceTableNames(state).some((name) => name.toLowerCase() === op.name.toLowerCase())) ||
      reservedNameConflict(state, op.name, op.scopeSheetId) ||
      current.preserveNames.some((name) => name.toLowerCase() === op.name.toLowerCase())
    )
      throw new Error('duplicate_defined_name')
    if (current.names.length >= 2000) throw new Error('defined_name_limit')
    const formula = op.ref.replace(/^=/, '')
    if (!formula.trim()) throw new Error('empty_defined_name_reference')
    current.names.push({
      name: op.name,
      formula,
      ...(op.scopeSheetId ? { scopeSheetId: op.scopeSheetId } : {}),
    })
  }
  if (current.names.length > 2000 || current.preserveNames.length > 2000)
    throw new Error('defined_name_limit')
  return current
}

/** Name Manager semantics: definition changes plus a replayable reference action. */
export function manageSessionDefinedName(
  state: Pick<CandidateState, 'features' | 'sheets'>,
  op: {
    name: string
    scopeSheetId?: string | undefined
    replacement: { name: string; ref: string } | null
  },
) {
  assertNameBindingsComplete(state)
  if (reservedNameConflict(state, op.name, op.scopeSheetId))
    throw new Error('unmodeled_defined_name_binding')
  const before = sessionDefinedNames(state)
  if (before.preserveNames.some((name) => name.toLowerCase() === op.name.toLowerCase()))
    throw new Error('unmodeled_defined_name_binding')
  const matches = before.names.filter(
    (name) =>
      name.name.toLowerCase() === op.name.toLowerCase() && name.scopeSheetId === op.scopeSheetId,
  )
  if (matches.length !== 1) throw new Error('defined_name_missing_or_ambiguous')
  const old = matches[0]!
  if (
    old.scopeSheetId === undefined &&
    op.replacement &&
    op.replacement.name.toLowerCase() !== old.name.toLowerCase()
  ) {
    const oldShadows = new Set(localBindingScopes(state, before, old.name))
    if (
      localBindingScopes(state, before, op.replacement.name).some((scope) => !oldShadows.has(scope))
    )
      throw new Error('defined_name_destination_shadowed')
  }
  let names = editSessionDefinedName(state, {
    op: 'delete_defined_name',
    name: old.name,
    ...(old.scopeSheetId ? { scopeSheetId: old.scopeSheetId } : {}),
  })
  if (op.replacement)
    names = editSessionDefinedName(
      {
        ...state,
        features: { ...state.features, definedNamesEdits: JSON.parse(JSON.stringify(names)) },
      },
      {
        op: 'add_defined_name',
        name: op.replacement.name,
        ref: op.replacement.ref,
        ...(old.scopeSheetId ? { scopeSheetId: old.scopeSheetId } : {}),
      },
    )
  const change =
    op.replacement?.name === old.name
      ? undefined
      : {
          name: old.name,
          replacement: op.replacement?.name ?? null,
          ...(old.scopeSheetId ? { scopeSheetId: old.scopeSheetId } : {}),
          shadowedSheetIds: localBindingScopes(state, before, old.name),
        }
  if (change) {
    const plan = { changes: [change], sheets: state.sheets.map(({ id, name }) => ({ id, name })) }
    names.names = names.names.map((entry) => {
      // The newly supplied replacement reference is already action-time text.
      if (entry.name === op.replacement?.name && entry.scopeSheetId === old.scopeSheetId)
        return entry
      return {
        ...entry,
        formula: rewriteNameReferences(entry.formula, entry.scopeSheetId ?? '', plan),
      }
    })
  }
  return { names, change }
}
