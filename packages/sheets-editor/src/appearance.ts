export type WorkbookAppearance = 'excel' | 'vscode'

export function workbookAppearance(value: unknown): WorkbookAppearance {
  return value === 'vscode' ? 'vscode' : 'excel'
}
