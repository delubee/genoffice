import * as vscode from 'vscode'
import { join } from 'node:path'
import { docxToolDefinitions, type DocxToolKind } from './docx-schemas'

// Keep the lazy boundary structural: importing the host's source types here
// would also pull the upstream Docs implementation into the XLSX type program.
type DocxDocument = vscode.CustomDocument
interface DocxEditorProvider extends vscode.Disposable {
  backupForTests(uri: vscode.Uri): Promise<vscode.CustomDocumentBackup>
  restoreForTests(uri: vscode.Uri, backupId: string): Promise<void>
  prepareTool(kind: DocxToolKind, raw: unknown): { message: string; confirmationRequired: boolean }
  invokeTool(kind: DocxToolKind, raw: unknown, token: vscode.CancellationToken): Promise<unknown>
  onDidChangeCustomDocument: vscode.Event<
    | vscode.CustomDocumentEditEvent<DocxDocument>
    | vscode.CustomDocumentContentChangeEvent<DocxDocument>
  >
  openCustomDocument(
    uri: vscode.Uri,
    options: vscode.CustomDocumentOpenContext,
  ): Promise<DocxDocument>
  resolveCustomEditor(document: DocxDocument, panel: vscode.WebviewPanel): Promise<void>
  saveCustomDocument(document: DocxDocument, token: vscode.CancellationToken): Promise<void>
  saveCustomDocumentAs(
    document: DocxDocument,
    destination: vscode.Uri,
    token: vscode.CancellationToken,
  ): Promise<void>
  revertCustomDocument(document: DocxDocument): Promise<void>
  backupCustomDocument(
    document: DocxDocument,
    context: vscode.CustomDocumentBackupContext,
    token: vscode.CancellationToken,
  ): Promise<vscode.CustomDocumentBackup>
}

/** DOCX engine/schema are parsed only when this editor is first used. */
export class LazyDocxProvider
  implements vscode.CustomEditorProvider<DocxDocument>, vscode.Disposable
{
  private provider?: DocxEditorProvider
  private changes = new vscode.EventEmitter<
    | vscode.CustomDocumentEditEvent<DocxDocument>
    | vscode.CustomDocumentContentChangeEvent<DocxDocument>
  >()
  readonly onDidChangeCustomDocument = this.changes.event
  constructor(private readonly context: vscode.ExtensionContext) {}
  private load(): DocxEditorProvider {
    if (!this.provider) {
      // eslint-disable-next-line @typescript-eslint/no-require-imports -- Load the separate CJS host only on first DOCX use.
      const host = require(join(this.context.extensionPath, 'dist', 'docx-host.cjs')) as {
        DocxEditorProvider: new (context: vscode.ExtensionContext) => DocxEditorProvider
      }
      this.provider = new host.DocxEditorProvider(this.context)
      this.provider.onDidChangeCustomDocument((event) => this.changes.fire(event))
    }
    return this.provider
  }
  openCustomDocument(uri: vscode.Uri, options: vscode.CustomDocumentOpenContext) {
    return this.load().openCustomDocument(uri, options)
  }
  resolveCustomEditor(document: DocxDocument, panel: vscode.WebviewPanel) {
    return this.load().resolveCustomEditor(document, panel)
  }
  saveCustomDocument(document: DocxDocument, token: vscode.CancellationToken) {
    return this.load().saveCustomDocument(document, token)
  }
  saveCustomDocumentAs(
    document: DocxDocument,
    destination: vscode.Uri,
    token: vscode.CancellationToken,
  ) {
    return this.load().saveCustomDocumentAs(document, destination, token)
  }
  revertCustomDocument(document: DocxDocument) {
    return this.load().revertCustomDocument(document)
  }
  backupCustomDocument(
    document: DocxDocument,
    context: vscode.CustomDocumentBackupContext,
    token: vscode.CancellationToken,
  ) {
    return this.load().backupCustomDocument(document, context, token)
  }
  registerTools() {
    return vscode.Disposable.from(
      ...docxToolDefinitions.map((definition) =>
        vscode.lm.registerTool<unknown>(definition.name, {
          prepareInvocation: (options, token) => {
            if (token.isCancellationRequested) throw new vscode.CancellationError()
            const prepared = this.load().prepareTool(definition.kind, options.input)
            const message = new vscode.MarkdownString().appendText(prepared.message)
            return {
              invocationMessage: message,
              ...(prepared.confirmationRequired
                ? { confirmationMessages: { title: 'Edit DOCX', message } }
                : {}),
            }
          },
          invoke: async (options, token) =>
            new vscode.LanguageModelToolResult([
              new vscode.LanguageModelTextPart(
                JSON.stringify(await this.load().invokeTool(definition.kind, options.input, token)),
              ),
            ]),
        }),
      ),
    )
  }
  backupForTests(uri: vscode.Uri) {
    return this.load().backupForTests(uri)
  }
  restoreForTests(uri: vscode.Uri, backupId: string) {
    return this.load().restoreForTests(uri, backupId)
  }
  dispose() {
    this.provider?.dispose()
    this.changes.dispose()
  }
}
