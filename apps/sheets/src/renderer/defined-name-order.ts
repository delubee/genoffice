export interface DefinedNameDefinition {
  name: string
  formula: string
  sheetIndex?: number | undefined
}

/** Match the desktop name table: live workbook names precede local/REF residues. */
export function orderDefinedNames<T extends DefinedNameDefinition>(names: readonly T[]): T[] {
  const groups = new Map<string, T[]>()
  for (const name of names) {
    const group = groups.get(name.name) ?? []
    group.push(name)
    groups.set(name.name, group)
  }
  const rank = (name: T) =>
    name.formula.includes('#REF!') ? 2 : name.sheetIndex === undefined ? 0 : 1
  return [...groups.values()].flatMap((group) => [...group].sort((a, b) => rank(a) - rank(b)))
}
