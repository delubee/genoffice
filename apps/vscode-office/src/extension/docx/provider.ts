import * as vscode from 'vscode'
import { nativeOfficeActions } from '../office-actions'
import { randomBytes } from 'node:crypto'
import { mkdir, readFile, unlink } from 'node:fs/promises'
import { join } from 'node:path'
import { z } from 'zod'
import { openDocument } from '@genoffice/docx-session/open'
import { saveDocument } from '@genoffice/docx-session/save'
import { JsonBridge } from '@genoffice/xlsx-gateway/shared/session-bridge'
import { workbookHtml } from '../webview-html'
import { DocxDocument } from './document'
import { writeBackup, restoreBackup } from '@genoffice/docx-session/backup'
import { DocumentBindings } from '@genoffice/docx-session/bindings'
import { docxToolSchemas, type DocxToolKind } from '../docx-schemas'
import { operationsSchema, type DocxOperation } from '@genoffice/docs-editor/shared/operations'
import { DocumentExchange } from '@genoffice/docx-session/exchange'

const requestSchema = z
  .object({
    sessionId: z.string().uuid(),
    epoch: z.number().int().nonnegative(),
    transactionId: z.string().min(1).max(200),
    expectedRevision: z.number().int().nonnegative(),
    baseCollabVersion: z.number().int().nonnegative(),
    clientId: z.string().min(1).max(200),
    steps: z.array(z.json()).max(1000),
  })
  .strict()

export class DocxEditorProvider
  implements vscode.CustomEditorProvider<DocxDocument>, vscode.Disposable
{
  private changes = new vscode.EventEmitter<
    | vscode.CustomDocumentEditEvent<DocxDocument>
    | vscode.CustomDocumentContentChangeEvent<DocxDocument>
  >()
  readonly onDidChangeCustomDocument = this.changes.event
  private documents = new Map<string, DocxDocument>()
  private restored = new WeakSet<DocxDocument>()
  private bindings = new DocumentBindings()
  private exchange = new DocumentExchange()
  constructor(private readonly context: vscode.ExtensionContext) {}
  getDocument(uri: vscode.Uri) {
    return this.documents.get(uri.toString())
  }
  private get storage() {
    return (this.context.storageUri ?? this.context.globalStorageUri).fsPath
  }

  async openCustomDocument(
    uri: vscode.Uri,
    options: vscode.CustomDocumentOpenContext,
  ): Promise<DocxDocument> {
    if (!vscode.workspace.isTrusted || uri.scheme !== 'file')
      throw new Error('DOCX requires a trusted local workspace')
    let opened
    if (options.backupId) {
      const backup = vscode.Uri.parse(options.backupId)
      if (backup.scheme !== 'file') throw new Error('Invalid DOCX recovery path')
      opened = await restoreBackup({
        path: backup.fsPath,
        storage: this.storage,
        sourceUri: uri.toString(),
      })
    } else opened = await openDocument(uri.fsPath)
    const document = new DocxDocument(uri, opened.session, opened.baseline)
    if (options.backupId) this.restored.add(document)
    this.documents.set(uri.toString(), document)
    document.onDidDispose(() => {
      this.documents.delete(uri.toString())
      this.bindings.release(uri.toString())
    })
    return document
  }

  async resolveCustomEditor(document: DocxDocument, panel: vscode.WebviewPanel) {
    const root = vscode.Uri.joinPath(this.context.extensionUri, 'dist', 'webview')
    panel.webview.options = { enableScripts: true, localResourceRoots: [root] }
    document.bridge?.dispose()
    document.prepareView()
    const bridge = (document.bridge = new JsonBridge(document.session.sessionId, {
      post: (message) => panel.webview.postMessage(message),
      subscribe(listener) {
        const subscription = panel.webview.onDidReceiveMessage(listener)
        return () => subscription.dispose()
      },
    }))
    bridge.handle('view.ready', () => document.projection())
    bridge.handle('view.resync', () => document.projection())
    const actionSession = document.session.sessionId
    bridge.handle(
      'office.action',
      nativeOfficeActions({
        uri: document.uri,
        available: () =>
          !document.isClosed &&
          this.getDocument(document.uri) === document &&
          document.session.sessionId === actionSession,
        active: () => panel.active,
        flush: () => document.flush(),
        bind: async (scope) => {
          const cancellation = new vscode.CancellationTokenSource()
          try {
            const context = (await this.invokeTool(
              'context',
              {
                documentUri: document.uri.toString(),
                selection: scope === 'selection',
              },
              cancellation.token,
            )) as { bindingId: string; descriptor: { revision: number } }
            return {
              kind: 'docx',
              scope,
              documentUri: document.uri.toString(),
              bindingId: context.bindingId,
              revision: context.descriptor.revision,
            }
          } finally {
            cancellation.dispose()
          }
        },
      }),
    )
    bridge.handle('view.presentation', (raw) => {
      const status = z
        .object({ ready: z.boolean().optional(), error: z.string().max(4096).optional() })
        .strict()
        .parse(raw)
      document.viewReady(status.error)
      return { ready: true }
    })
    bridge.handle('view.intent', (raw) =>
      document.exclusive(() => {
        const revision = document.session.descriptor().revision
        const commit = document.session.submit(requestSchema.parse(raw))
        if (commit.changed && commit.descriptor.revision > revision) this.record(document)
        return commit
      }),
    )
    bridge.handle('docx.operations', (raw) =>
      document.exclusive(() => {
        const input = z
          .object({
            sessionId: z.string().uuid(),
            transactionId: z.string().min(1).max(200),
            expectedRevision: z.number().int().nonnegative(),
            operations: operationsSchema,
          })
          .strict()
          .parse(raw)
        if (input.sessionId !== document.session.sessionId) throw new Error('Wrong DOCX session')
        const revision = document.session.descriptor().revision
        const commit = document.session.applyOperations(input)
        if (commit.changed && commit.descriptor.revision > revision) this.record(document)
        return commit
      }),
    )
    bridge.handle('history.request', async (raw) => {
      if (!panel.active) throw new Error('Activate this DOCX editor before requesting history')
      const { direction } = z
        .object({ direction: z.enum(['undo', 'redo']) })
        .strict()
        .parse(raw)
      await vscode.commands.executeCommand(direction)
      return { ready: true }
    })
    const appearance = () =>
      vscode.workspace.getConfiguration('genoffice.docx', document.uri).get('appearance') ===
      'vscode'
        ? 'vscode'
        : 'word'
    const configuration = vscode.workspace.onDidChangeConfiguration((event) => {
      if (event.affectsConfiguration('genoffice.docx.appearance', document.uri))
        void bridge.request('appearance.changed', { appearance: appearance() }).catch(() => {})
    })
    panel.onDidDispose(() => {
      configuration.dispose()
      bridge.dispose()
      if (document.bridge === bridge) document.bridge = undefined
    })
    const html = await readFile(join(root.fsPath, 'docx.html'), 'utf8')
    panel.webview.html = workbookHtml({
      html,
      root: root.fsPath,
      nonce: randomBytes(24).toString('base64'),
      cspSource: panel.webview.cspSource,
      language: vscode.env.language,
      sessionId: document.session.sessionId,
      resource: (path) => panel.webview.asWebviewUri(vscode.Uri.file(path)).toString(),
    }).replace('<html ', `<html data-docx-appearance="${appearance()}" `)
    if (this.restored.delete(document)) this.changes.fire({ document })
  }

  private record(document: DocxDocument) {
    this.changes.fire({
      document,
      label: 'DOCX edit',
      undo: () => this.history(document, 'undo'),
      redo: () => this.history(document, 'redo'),
    })
  }
  private boundDocument(bindingId: string) {
    const binding = this.bindings.resolve(bindingId)
    const document = this.documents.get(binding.uri)
    if (!document || document.session.sessionId !== binding.sessionId)
      throw new Error('DOCX session changed; request context again')
    return document
  }
  private scope(document: DocxDocument, bindingId: string, operations?: DocxOperation[]) {
    const selection = this.bindings.resolve(bindingId).selection
    if (!selection) return undefined
    const descriptor = document.session.descriptor()
    if (selection.stateId !== descriptor.stateId || selection.revision !== descriptor.revision)
      throw new Error('Selection binding is stale; bind the current selection again')
    for (const operation of operations ?? []) {
      const block = selection.blocks.find((item) => item.blockId === operation.blockId)
      if (!block || block.protected)
        throw new Error('Operation is outside the frozen editable selection')
      if (
        (operation.type === 'replaceText' || operation.type === 'formatText') &&
        (operation.start < (block.start ?? 0) || operation.end > (block.end ?? 0))
      )
        throw new Error('Text range is outside the frozen selection')
      if (operation.type === 'deleteParagraph' && (block.start !== 0 || block.end !== block.length))
        throw new Error('Select the whole paragraph before deleting it')
    }
    return selection
  }
  prepareTool(kind: DocxToolKind, raw: unknown) {
    const input = docxToolSchemas[kind].parse(raw)
    const document = 'bindingId' in input ? this.boundDocument(input.bindingId) : undefined
    return {
      message:
        kind === 'apply'
          ? `Edit ${document!.uri.fsPath}: ${(input as { summary: string }).summary}`
          : kind === 'applyImport'
            ? `Apply the prepared Python text changes to ${document!.uri.fsPath}`
            : `GenOffice DOCX ${kind}${document ? `: ${document.uri.fsPath}` : ''}`,
      confirmationRequired: ['apply', 'applyImport', 'export'].includes(kind),
    }
  }
  async invokeTool(
    kind: DocxToolKind,
    raw: unknown,
    token: vscode.CancellationToken,
  ): Promise<unknown> {
    if (!vscode.workspace.isTrusted) throw new Error('DOCX tools require a trusted workspace')
    if (token.isCancellationRequested) throw new vscode.CancellationError()
    if (kind === 'context') {
      const input = docxToolSchemas.context.parse(raw)
      if (!input.documentUri)
        return {
          documents: [...this.documents.values()].map((document) => ({
            documentUri: document.uri.toString(),
            descriptor: document.session.descriptor(),
          })),
        }
      const document = this.documents.get(input.documentUri)
      if (!document) throw new Error('Open this document with GenOffice Word first')
      await document.flush()
      const selection = input.selection
        ? document.session.selection(
            z
              .object({
                stateId: z.string().uuid(),
                revision: z.number().int().nonnegative(),
                from: z.number().int().nonnegative(),
                to: z.number().int().nonnegative(),
              })
              .strict()
              .parse(await document.bridge?.request('view.selection', {})),
          )
        : undefined
      return {
        documentUri: document.uri.toString(),
        bindingId: this.bindings.bind(
          document.uri.toString(),
          document.session.sessionId,
          selection,
        ),
        descriptor: document.session.descriptor(),
        ...(selection ? { selection } : {}),
        supportedOperations: [
          'replaceText',
          'formatParagraph',
          'setHeading',
          'setList',
          'insertParagraph',
          'deleteParagraph',
          'formatText',
        ],
        limits: { readBlocks: 100, operations: 100 },
      }
    }
    if (kind === 'read') {
      const input = docxToolSchemas.read.parse(raw)
      const document = this.boundDocument(input.bindingId)
      await document.flush()
      if (input.expectedRevision !== document.session.descriptor().revision)
        throw new Error('Stale document version')
      const selection = this.scope(document, input.bindingId)
      if (selection)
        return {
          descriptor: document.session.descriptor(),
          blocks: selection.blocks.slice(input.offset, input.offset + input.limit),
        }
      return document.session.read(input.offset, input.limit)
    }
    if (kind === 'operationStatus') {
      const input = docxToolSchemas.operationStatus.parse(raw)
      return this.boundDocument(input.bindingId).session.operationStatus(input.transactionId)
    }
    const roots =
      vscode.workspace.workspaceFolders
        ?.filter((folder) => folder.uri.scheme === 'file')
        .map((folder) => folder.uri.fsPath) ?? []
    if (kind === 'export') {
      const input = docxToolSchemas.export.parse(raw)
      const document = this.boundDocument(input.bindingId)
      await document.flush()
      if (this.bindings.resolve(input.bindingId).selection)
        throw new Error('Use a document binding without selection for full-document export')
      return this.exchange.export(document.session, { ...input, roots })
    }
    if (kind === 'prepareImport') {
      const input = docxToolSchemas.prepareImport.parse(raw)
      const document = this.boundDocument(input.bindingId)
      await document.flush()
      if (this.bindings.resolve(input.bindingId).selection)
        throw new Error('Use a document binding without selection for full-document import')
      return this.exchange.prepareImport(document.session, { ...input, roots })
    }
    if (kind === 'applyImport') {
      const input = docxToolSchemas.applyImport.parse(raw)
      const document = this.boundDocument(input.bindingId)
      await document.flush(true)
      return document.exclusive(async () => {
        try {
          if (token.isCancellationRequested) throw new vscode.CancellationError()
          if (this.bindings.resolve(input.bindingId).selection)
            throw new Error('Use a document binding without selection for full-document import')
          const revision = document.session.descriptor().revision
          const commit = await this.exchange.applyImport(document.session, {
            ...input,
            roots,
            assertActive: () => {
              if (document.isClosed) throw new Error('DOCX session is closed')
              if (token.isCancellationRequested) throw new vscode.CancellationError()
            },
          })
          if (commit.changed && commit.descriptor.revision > revision) this.record(document)
          await document.bridge?.request('state.committed', commit)
          return {
            status: 'applied',
            transactionId: input.transactionId,
            descriptor: commit.descriptor,
          }
        } finally {
          await document.project()
        }
      })
    }
    const input = docxToolSchemas.apply.parse(raw)
    const document = this.boundDocument(input.bindingId)
    await document.flush(true)
    return document.exclusive(async () => {
      try {
        if (token.isCancellationRequested) throw new vscode.CancellationError()
        const revision = document.session.descriptor().revision
        if (document.session.operationStatus(input.transactionId).status === 'unknown')
          this.scope(document, input.bindingId, input.operations)
        const commit = document.session.applyOperations(input)
        if (commit.changed && commit.descriptor.revision > revision) this.record(document)
        await document.bridge?.request('state.committed', commit)
        return {
          status: 'applied',
          transactionId: input.transactionId,
          descriptor: commit.descriptor,
        }
      } finally {
        await document.project()
      }
    })
  }
  private async history(document: DocxDocument, direction: 'undo' | 'redo') {
    await document.flush(true)
    return document.exclusive(async () => {
      try {
        document.session.history(direction)
      } finally {
        await document.project()
      }
    })
  }
  async saveCustomDocument(document: DocxDocument, token: vscode.CancellationToken) {
    await document.flush(true)
    return document.exclusive(async () => {
      try {
        if (token.isCancellationRequested) throw new vscode.CancellationError()
        await mkdir(this.storage, { recursive: true, mode: 0o700 })
        document.baseline = await saveDocument({
          session: document.session,
          target: document.uri.fsPath,
          baseline: document.baseline,
          storage: this.storage,
          assertActive: () => {
            if (document.isClosed || token.isCancellationRequested)
              throw new vscode.CancellationError()
          },
        })
        document.session.markSaved()
      } finally {
        await document.project()
      }
    })
  }
  async saveCustomDocumentAs(
    document: DocxDocument,
    destination: vscode.Uri,
    token: vscode.CancellationToken,
  ) {
    if (destination.scheme !== 'file') throw new Error('DOCX save requires a local file')
    await document.flush(true)
    return document.exclusive(async () => {
      try {
        if (token.isCancellationRequested) throw new vscode.CancellationError()
        await mkdir(this.storage, { recursive: true, mode: 0o700 })
        await saveDocument({
          session: document.session,
          target: destination.fsPath,
          storage: this.storage,
          assertActive: () => {
            if (document.isClosed || token.isCancellationRequested)
              throw new vscode.CancellationError()
          },
        })
      } finally {
        await document.project()
      }
    })
  }
  async revertCustomDocument(document: DocxDocument) {
    await document.flush(true)
    return document.exclusive(async () => {
      try {
        const opened = await openDocument(document.uri.fsPath)
        document.session = opened.session
        document.baseline = opened.baseline
      } finally {
        await document.project()
      }
    })
  }
  async backupCustomDocument(
    document: DocxDocument,
    _context: vscode.CustomDocumentBackupContext,
    token: vscode.CancellationToken,
  ): Promise<vscode.CustomDocumentBackup> {
    await document.flush(true)
    return document.exclusive(async () => {
      try {
        if (token.isCancellationRequested) throw new vscode.CancellationError()
        const path = await writeBackup({
          session: document.session,
          baseline: document.baseline,
          sourceUri: document.uri.toString(),
          storage: this.storage,
        })
        return {
          id: vscode.Uri.file(path).toString(),
          delete: async () => {
            await unlink(path).catch(() => {})
          },
        }
      } finally {
        await document.project()
      }
    })
  }
  async backupForTests(uri: vscode.Uri) {
    if (this.context.extensionMode !== vscode.ExtensionMode.Test)
      throw new Error('Test mode required')
    const document = this.getDocument(uri)
    if (!document) throw new Error('Test DOCX is not open')
    const cancellation = new vscode.CancellationTokenSource()
    try {
      return await this.backupCustomDocument(
        document,
        { destination: this.context.globalStorageUri },
        cancellation.token,
      )
    } finally {
      cancellation.dispose()
    }
  }
  async restoreForTests(uri: vscode.Uri, backupId: string) {
    if (this.context.extensionMode !== vscode.ExtensionMode.Test)
      throw new Error('Test mode required')
    const document = this.getDocument(uri)
    if (!document) throw new Error('Test DOCX is not open')
    await document.flush(true)
    return document.exclusive(async () => {
      try {
        const restored = await restoreBackup({
          path: vscode.Uri.parse(backupId).fsPath,
          storage: this.storage,
          sourceUri: uri.toString(),
        })
        document.session = restored.session
        document.baseline = restored.baseline
        this.bindings.release(uri.toString())
        this.changes.fire({ document })
      } finally {
        await document.project()
      }
    })
  }
  dispose() {
    for (const document of this.documents.values()) document.dispose()
    this.changes.dispose()
  }
}
