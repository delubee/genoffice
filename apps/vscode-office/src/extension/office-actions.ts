import * as vscode from 'vscode'
import { officeActionHandler, type OfficeChatBinding } from './office-action-handler'

/** Native operations stay in the extension host; the webview cannot choose a command or URI. */
export function nativeOfficeActions(input: {
  uri: vscode.Uri
  available(): boolean
  active(): boolean
  flush(): Promise<void>
  bind(scope: OfficeChatBinding['scope']): Promise<OfficeChatBinding>
}) {
  return officeActionHandler({
    ...input,
    available: () => vscode.workspace.isTrusted && input.available(),
    async open() {
      const targets = await vscode.window.showOpenDialog({
        canSelectMany: false,
        filters: { 'GenOffice documents': ['xlsx', 'docx'] },
      })
      if (!targets?.[0]) return
      if (!input.available()) throw new Error('This document session is no longer available.')
      const target = targets[0]
      const editor = /\.docx$/i.test(target.path)
        ? 'genoffice.docx'
        : /\.xlsx$/i.test(target.path)
          ? 'genoffice.workbook'
          : null
      if (!editor || target.scheme !== 'file') throw new Error('Choose a local DOCX or XLSX file.')
      await vscode.commands.executeCommand('vscode.openWith', target, editor)
    },
    async save(asCopy) {
      await vscode.commands.executeCommand(
        asCopy ? 'workbench.action.files.saveAs' : 'workbench.action.files.save',
        input.uri,
      )
    },
    async chat(query, partial) {
      await vscode.commands.executeCommand('workbench.action.chat.open', {
        mode: 'agent',
        query,
        isPartialQuery: partial,
      })
    },
  })
}
