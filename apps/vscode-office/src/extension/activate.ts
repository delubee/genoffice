import * as vscode from 'vscode'
import { WorkbookEditorProvider, EDITOR_ID } from './provider'
import { registerWorkbookTools } from './tools/register'
import { registerBindSelection } from './commands/bind-selection'
import { registerWorkbookSkillCommands } from './commands/install-skill'
export function activate(context: vscode.ExtensionContext) {
  const output = vscode.window.createOutputChannel('GenOffice Workbook')
  const provider = new WorkbookEditorProvider(context, output)
  context.subscriptions.push(
    output,
    provider,
    registerWorkbookTools(provider),
    registerBindSelection(provider),
    registerWorkbookSkillCommands(context),
    vscode.commands.registerCommand('genoffice.recoverWorkbook', (uri?: vscode.Uri) =>
      provider.recoverWorkbook(uri),
    ),
    vscode.window.registerCustomEditorProvider(EDITOR_ID, provider, {
      supportsMultipleEditorsPerDocument: false,
      webviewOptions: { retainContextWhenHidden: true },
    }),
    vscode.commands.registerCommand('genoffice.openWorkbook', async (uri?: vscode.Uri) => {
      const target =
        uri ??
        (
          await vscode.window.showOpenDialog({
            canSelectMany: false,
            filters: { Workbook: ['xlsx'] },
          })
        )?.[0]
      if (target) await vscode.commands.executeCommand('vscode.openWith', target, EDITOR_ID)
    }),
    vscode.commands.registerCommand('genoffice.saveWorkbookAs', async (uri?: vscode.Uri) => {
      await provider.saveCopy(uri)
    }),
  )
  output.appendLine(
    `GenOffice host: ${process.platform}/${process.arch}; VS Code ${vscode.version}`,
  )
  return {
    getDocument: provider.getDocument.bind(provider),
    bindings: provider.bindings,
    operationStatus: provider.operationStatus.bind(provider),
    ...(context.extensionMode === vscode.ExtensionMode.Test
      ? {
          revertForTests: async (uri: vscode.Uri) => {
            const document = provider.getDocument(uri)
            if (!document) throw new Error('The test workbook is not open.')
            await provider.revertCustomDocument(document)
          },
        }
      : {}),
  }
}
