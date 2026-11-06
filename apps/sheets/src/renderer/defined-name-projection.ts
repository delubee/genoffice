/** Install source/session names through the desktop facade without making it the authority. */
export function installNameDefinitions(
  workbook: unknown,
  names: readonly { name: string; formula: string; sheetId?: string }[],
): string[] {
  const facade = workbook as {
    newDefinedNameBuilder(): { load(param: Record<string, unknown>): { build(): unknown } }
    insertDefinedNameBuilder(param: unknown): void
  }
  const uninstalled: string[] = []
  for (const name of names) {
    try {
      facade.insertDefinedNameBuilder(
        facade
          .newDefinedNameBuilder()
          .load({
            name: name.name,
            formulaOrRefString: name.formula,
            localSheetId: name.sheetId ?? 'AllDefaultWorkbook',
          })
          .build(),
      )
    } catch {
      uninstalled.push(name.name)
    }
  }
  return uninstalled
}
