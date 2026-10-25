import { Editor, Extension } from '@tiptap/core'
import { EditorState, Plugin, Selection, TextSelection } from '@tiptap/pm/state'
import { Step } from '@tiptap/pm/transform'
import { collab, getVersion, receiveTransaction, sendableSteps } from 'prosemirror-collab'
import { documentExtensions } from './shared/schema'
import { validateAndNormalize } from './shared/policy'
import type { EditorJSON } from './conversion'
import type {
  Commit,
  DocumentProjection,
  StepRequest,
  DocumentPresentation,
} from './shared/protocol'
import type { DocxOperation } from './shared/operations'
import type { OfficeAction } from '@genoffice/xlsx-gateway/shared/office-actions'

export interface DocumentBridge {
  action?(request: OfficeAction): Promise<void>
  submit(request: StepRequest): Promise<Commit>
  history(direction: 'undo' | 'redo'): Promise<void>
  resync(): Promise<DocumentProjection>
  error(message: string): void
  operations(request: {
    sessionId: string
    transactionId: string
    expectedRevision: number
    operations: DocxOperation[]
  }): Promise<Commit>
  presentation?(presentation: DocumentPresentation): void
}

/** One native undo stack; collab retains only unacknowledged local steps. */
export class DocumentRuntime {
  get hasOfficeActions() {
    return !!this.bridge.action
  }
  async officeAction(request: OfficeAction) {
    if (!this.bridge.action) throw new Error('Office menu actions are unavailable in this host.')
    if (!(await this.flush())) throw new Error('Finish editing before using this menu command.')
    await this.bridge.action(request)
  }
  reportError(error: unknown) {
    this.bridge.error(String(error))
  }
  readonly editor: Editor
  private descriptor
  private readonly clientId = crypto.randomUUID()
  private pending?: Promise<void>
  private disposed = false
  private hostFrozen = false
  private historyPending = false
  private unavailable = false
  private operationPending = false
  private get frozen() {
    return this.hostFrozen || this.historyPending || this.operationPending || this.unavailable
  }

  constructor(
    initial: DocumentProjection,
    private readonly bridge: DocumentBridge,
  ) {
    this.descriptor = initial.descriptor
    // eslint-disable-next-line @typescript-eslint/no-this-alias -- Plugin methods have their own this; retain the document authority.
    const runtime = this
    const collaboration = Extension.create({
      name: 'officeCollaboration',
      addProseMirrorPlugins: () => [
        collab({ version: initial.descriptor.collabVersion, clientID: runtime.clientId }),
        new Plugin({
          filterTransaction(tr) {
            if (!tr.docChanged || tr.getMeta('officeRemote')) return true
            if (runtime.frozen || runtime.descriptor.readonly) return false
            try {
              validateAndNormalize(
                tr.before.toJSON() as EditorJSON,
                tr.doc.toJSON() as EditorJSON,
                () => 'pending',
              )
              return true
            } catch (error) {
              bridge.error(String(error))
              return false
            }
          },
        }),
      ],
      addKeyboardShortcuts: () => ({
        'Mod-z': () => {
          void runtime.history('undo')
          return true
        },
        'Mod-Shift-z': () => {
          void runtime.history('redo')
          return true
        },
        'Mod-y': () => {
          void runtime.history('redo')
          return true
        },
      }),
    })
    this.editor = new Editor({
      extensions: [
        ...documentExtensions.map((extension) =>
          extension.name === 'listNumbering' && initial.presentation
            ? extension.extend({
                addStorage: () => ({
                  defs: new Map(initial.presentation!.numbering),
                  styles: new Map(initial.presentation!.styles),
                  docDefaults: initial.presentation!.docDefaults,
                }),
              })
            : extension,
        ),
        collaboration,
      ],
      content: initial.content,
      editable: !initial.descriptor.readonly,
      onTransaction: () => {
        queueMicrotask(() => this.pump())
      },
    })
  }

  private pump() {
    if (this.disposed || this.pending || this.frozen) return
    const batch = sendableSteps(this.editor.state)
    if (!batch) return
    const request: StepRequest = {
      sessionId: this.descriptor.sessionId,
      epoch: this.descriptor.epoch,
      expectedRevision: this.descriptor.revision,
      baseCollabVersion: batch.version,
      clientId: this.clientId,
      transactionId: crypto.randomUUID(),
      steps: batch.steps.map((step) => step.toJSON()),
    }
    this.pending = this.bridge
      .submit(request)
      .then((commit) => this.accept(commit))
      .catch(async (error) => {
        if (this.disposed) return
        this.bridge.error(String(error))
        this.unavailable = true
        // Rejected optimistic content must never remain displayed as committed.
        // Recreate the collab plugin and document from an authoritative projection.
        try {
          this.reset(await this.bridge.resync())
        } catch (resyncError) {
          this.bridge.error(String(resyncError))
          this.updateEditability()
        }
      })
      .finally(() => {
        this.pending = undefined
        this.pump()
      })
  }

  accept(commit: Commit) {
    if (this.disposed) return
    if (
      commit.descriptor.sessionId !== this.descriptor.sessionId ||
      commit.descriptor.epoch !== this.descriptor.epoch
    )
      throw new Error('Collaboration session changed')
    const current = getVersion(this.editor.state)
    if (commit.descriptor.collabVersion <= current) return
    if (commit.fromVersion !== current) throw new Error('Missing collaboration commit')
    const tr = receiveTransaction(
      this.editor.state,
      commit.steps.map((json) => Step.fromJSON(this.editor.schema, json)),
      commit.clientIds,
    )
    if (commit.presentation) this.present(commit.presentation)
    tr.setMeta('officeRemote', true)
    this.descriptor = commit.descriptor
    this.editor.view.dispatch(tr)
  }

  async flush(): Promise<boolean> {
    if (this.unavailable) return false
    const deadline = Date.now() + 10000
    while (!this.disposed && Date.now() < deadline) {
      if (!this.editor.view.composing) {
        this.pump()
        if (!this.operationPending && !this.pending && !sendableSteps(this.editor.state))
          return true
      }
      await new Promise((resolve) => setTimeout(resolve, 20))
    }
    return false
  }
  selection() {
    return {
      stateId: this.descriptor.stateId,
      revision: this.descriptor.revision,
      from: this.editor.state.selection.from,
      to: this.editor.state.selection.to,
    }
  }

  async history(direction: 'undo' | 'redo') {
    try {
      if (this.hostFrozen || this.unavailable) throw new Error('DOCX editor is busy')
      if (!(await this.flush())) throw new Error('Document is busy composing or synchronizing')
      this.historyPending = true
      this.updateEditability()
      await this.bridge.history(direction)
    } catch (error) {
      this.bridge.error(String(error))
    } finally {
      this.historyPending = false
      this.updateEditability()
    }
  }

  async applyToSelection(create: (blockId: string) => DocxOperation) {
    try {
      if (this.hostFrozen || this.unavailable) throw new Error('DOCX editor is busy')
      if (!(await this.flush())) throw new Error('Document is busy composing or synchronizing')
      this.operationPending = true
      this.updateEditability()
      const ids = new Set<string>()
      const { selection, doc } = this.editor.state
      doc.nodesBetween(selection.from, selection.to, (node) => {
        if (
          ['docParagraph', 'docHeading', 'docListItem'].includes(node.type.name) &&
          typeof node.attrs.officeId === 'string'
        )
          ids.add(node.attrs.officeId)
      })
      if (!ids.size && selection.empty) {
        for (let depth = selection.$from.depth; depth > 0; depth--) {
          const node = selection.$from.node(depth)
          if (
            ['docParagraph', 'docHeading', 'docListItem'].includes(node.type.name) &&
            typeof node.attrs.officeId === 'string'
          ) {
            ids.add(node.attrs.officeId)
            break
          }
        }
      }
      if (!ids.size) throw new Error('Select an editable paragraph')
      const commit = await this.bridge.operations({
        sessionId: this.descriptor.sessionId,
        transactionId: crypto.randomUUID(),
        expectedRevision: this.descriptor.revision,
        operations: [...ids].map(create),
      })
      this.accept(commit)
    } catch (error) {
      this.bridge.error(String(error))
    } finally {
      this.operationPending = false
      this.updateEditability()
    }
  }

  freeze(value: boolean) {
    this.hostFrozen = value
    this.updateEditability()
  }
  private updateEditability() {
    if (!this.disposed) this.editor.setEditable(!this.frozen && !this.descriptor.readonly)
  }

  reset(projection: DocumentProjection) {
    if (this.disposed) return
    if (projection.presentation) this.present(projection.presentation)
    // Reconfigure creates a fresh collaboration state with no local undo stack.
    const kept = this.editor.state.plugins.filter(
      (plugin) => !(plugin as unknown as { key: string }).key.startsWith('collab$'),
    )
    const selection = this.editor.state.selection
    const doc = this.editor.schema.nodeFromJSON(projection.content)
    let restoredSelection: Selection
    try {
      restoredSelection = Selection.fromJSON(doc, selection.toJSON())
    } catch {
      restoredSelection = TextSelection.near(
        doc.resolve(Math.min(selection.from, doc.content.size)),
      )
    }
    this.editor.view.updateState(
      EditorState.create({
        schema: this.editor.schema,
        doc,
        selection: restoredSelection,
        plugins: [
          ...kept,
          collab({ version: projection.descriptor.collabVersion, clientID: this.clientId }),
        ],
      }),
    )
    this.descriptor = projection.descriptor
    this.unavailable = false
    this.freeze(false)
    // updateState bypasses Tiptap's transaction events. Refresh selection-driven
    // controls now, before focus enters an input whose key depends on that state.
    this.editor.view.dispatch(this.editor.state.tr.setMeta('officeRemote', true))
  }

  private present(presentation: DocumentPresentation) {
    this.bridge.presentation?.(presentation)
    const storage = this.editor.storage.listNumbering
    storage.defs = new Map(presentation.numbering)
    storage.styles = new Map(presentation.styles)
    storage.docDefaults = presentation.docDefaults
  }

  dispose() {
    this.disposed = true
    this.editor.destroy()
  }
}
