import * as vscode from 'vscode'
import type { WorkbookEditorProvider } from '../provider'
import { WorkbookTools } from './service'
import { workbookToolDefinitions } from './schemas'
import { WorkbookExchangeTools } from './exchange-tools'
import { flushCurrentEditor } from './current-state'

export function registerWorkbookTools(provider: WorkbookEditorProvider) {
  const service = new WorkbookTools({
    exchange: new WorkbookExchangeTools(provider),
    bindings: provider.bindings,
    documents: () =>
      provider.openDocuments().map((document) => ({
        descriptor: () => document.engine.descriptor,
        state: () => document.engine.state,
        projectionError: () =>
          document.engine.projectionError ? String(document.engine.projectionError) : null,
        flush: (signal) => flushCurrentEditor(document, signal),
        snapshot: () => document.snapshot(),
        submit: (intent, signal) => document.engine.submit(intent, signal),
        operationStatus: (transactionId) => document.engine.operationStatus(transactionId),
      })),
    operationStatus: (sessionId, transactionId) =>
      provider.operationStatus(sessionId, transactionId),
  })
  return vscode.Disposable.from(
    ...workbookToolDefinitions.map((definition) =>
      vscode.lm.registerTool<unknown>(definition.name, {
        prepareInvocation(options, token) {
          if (token.isCancellationRequested) throw new vscode.CancellationError()
          const prepared = service.prepare(definition.kind, options.input)
          const message = new vscode.MarkdownString().appendText(prepared.message)
          return {
            invocationMessage: message,
            ...(prepared.confirmationRequired
              ? { confirmationMessages: { title: prepared.title, message } }
              : {}),
          }
        },
        async invoke(options, token) {
          const controller = new AbortController(),
            subscription = token.onCancellationRequested(() => controller.abort())
          if (token.isCancellationRequested) controller.abort()
          try {
            const result = await service.invoke(definition.kind, options.input, controller.signal)
            return new vscode.LanguageModelToolResult([
              new vscode.LanguageModelTextPart(JSON.stringify(result)),
            ])
          } finally {
            subscription.dispose()
          }
        },
      }),
    ),
  )
}
