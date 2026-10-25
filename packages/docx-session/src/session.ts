import { createHash, randomUUID } from 'node:crypto'
import {
  parseDocx,
  saveDocx,
  type ParsedDocFull,
  type StyleUpsert,
  type NumberingDef,
} from '@genoffice/docx-engine'
import { Transform, Step } from '@tiptap/pm/transform'
import type { Node as PMNode } from '@tiptap/pm/model'
import {
  documentToEditor,
  editorToSavePlan,
  docStyleCss,
  docThemeCss,
  docBodyFont,
  type EditorJSON,
} from '@genoffice/docs-editor/conversion'
import { docxSchema } from '@genoffice/docs-editor/shared/schema'
import {
  DOCX_LIMITS,
  jsonBytes,
  prepareDocument,
  validateAndNormalize,
} from '@genoffice/docs-editor/shared/policy'
import { operationSteps } from './operations'
import { readBlocks, type SelectionScope } from './read'
import type {
  SessionDescriptor,
  StepRequest,
  Commit,
  DocumentPresentation,
} from '@genoffice/docs-editor/shared/protocol'
export type { SessionDescriptor, StepRequest, Commit } from '@genoffice/docs-editor/shared/protocol'

interface Overlays {
  numbering: Array<{ numId: string; kind: 'bullet' | 'ordered' }>
  styles: StyleUpsert[]
}
interface Root {
  doc: PMNode
  stateId: string
  bytes: number
  overlays: Overlays
}

/** Authoritative host state. Every save derives from the immutable opened package. */
export class DocxSession {
  readonly sessionId = randomUUID()
  private root: Root
  private revision = 0
  private collabVersion = 0
  private epoch = 0
  private undoRoots: Root[] = []
  private redoRoots: Root[] = []
  private receipts = new Map<string, { hash: string; commit: Commit }>()
  private operationReceipts = new Map<string, { hash: string; commit: Commit }>()
  private nextNumId = 1
  private readonly baseCSS: string
  private savedStateId: string | null
  private conflict = false
  readonly readonly: boolean

  private constructor(private readonly original: ParsedDocFull) {
    this.readonly = Boolean(original.protection?.enforced || original.writeProtection?.hash)
    const json = prepareDocument(documentToEditor(original), randomUUID)
    if (jsonBytes(json) > DOCX_LIMITS.root) throw new Error('Document exceeds size limit')
    this.root = {
      doc: docxSchema.nodeFromJSON(json),
      stateId: randomUUID(),
      bytes: jsonBytes(json),
      overlays: { numbering: [], styles: [] },
    }
    this.savedStateId = this.root.stateId
    this.baseCSS =
      docStyleCss(original) +
      docThemeCss(original.themeFonts, original.themeColors, Boolean(docBodyFont(original)))
    for (const numId of original.numbering.keys())
      this.nextNumId = Math.max(this.nextNumId, (Number(numId) || 0) + 1)
  }

  static async open(bytes: Uint8Array): Promise<DocxSession> {
    if (bytes.byteLength > DOCX_LIMITS.file) throw new Error('DOCX exceeds 16 MiB limit')
    return new DocxSession(await parseDocx(Uint8Array.from(bytes)))
  }

  descriptor(): SessionDescriptor {
    return {
      schemaVersion: 1,
      sessionId: this.sessionId,
      stateId: this.root.stateId,
      revision: this.revision,
      collabVersion: this.collabVersion,
      epoch: this.epoch,
      readonly: this.readonly,
      dirty: this.root.stateId !== this.savedStateId,
      conflict: this.conflict,
    }
  }
  content(): EditorJSON {
    return this.root.doc.toJSON() as EditorJSON
  }
  markSaved() {
    this.savedStateId = this.root.stateId
    this.conflict = false
  }
  markUnsaved() {
    this.savedStateId = null
  }
  markConflict() {
    this.conflict = true
  }
  presentation(): DocumentPresentation {
    const numbering = new Map(this.original.numbering)
    for (const definition of this.root.overlays.numbering)
      numbering.set(definition.numId, {
        numId: definition.numId,
        abstractNumId: `pending-${definition.numId}`,
        levels: {
          0: {
            numFmt: definition.kind === 'ordered' ? 'decimal' : 'bullet',
            lvlText: definition.kind === 'ordered' ? '%1.' : '•',
            start: 1,
            indentLeft: 720,
            hanging: 360,
          },
        },
        startOverrides: {},
      } as NumberingDef)
    // Engine models contain optional undefined properties; explicitly project
    // trusted parsed data into JSON before the strict transport sees it.
    const pendingCSS = this.root.overlays.styles
      .map(
        (style) =>
          `.doc-page [data-style="${style.styleId}"] { font-weight:700; font-size:${(style.rPr?.sizeHalfPoints ?? 24) / 2}pt }`,
      )
      .join('\n')
    return JSON.parse(
      JSON.stringify({
        css: this.baseCSS + pendingCSS,
        numbering: [...numbering.entries()],
        styles: [...this.original.styles.entries()],
        docDefaults: this.original.docDefaults,
      }),
    ) as DocumentPresentation
  }
  read(offset = 0, limit = 100) {
    return { descriptor: this.descriptor(), blocks: readBlocks(this.content(), offset, limit) }
  }
  selection(input: {
    stateId: string
    revision: number
    from: number
    to: number
  }): SelectionScope {
    if (input.stateId !== this.root.stateId || input.revision !== this.revision)
      throw new Error('Stale selection version')
    const { from, to } = input
    if (
      !Number.isInteger(from) ||
      !Number.isInteger(to) ||
      from < 0 ||
      to < from ||
      to > this.root.doc.content.size
    )
      throw new Error('Invalid selection positions')
    const blocks: SelectionScope['blocks'] = []
    this.root.doc.nodesBetween(from, to, (node, pos) => {
      if (node.attrs.officeLocked) {
        blocks.push({ blockId: node.attrs.officeId, type: node.type.name, protected: true })
        return false
      }
      if (!['docParagraph', 'docHeading', 'docListItem'].includes(node.type.name)) return
      const start = Math.min(node.content.size, Math.max(0, from - pos - 1)),
        end = Math.max(start, Math.min(node.content.size, to - pos - 1))
      blocks.push({
        blockId: node.attrs.officeId,
        type: node.type.name,
        protected: false,
        start,
        end,
        length: node.content.size,
        text: node.textBetween(start, Math.max(start, end), '\n'),
      })
      return false
    })
    if (blocks.length > 100 || jsonBytes(blocks) > 256 * 1024)
      throw new Error('Selection exceeds read limit')
    return { stateId: this.root.stateId, revision: this.revision, from, to, blocks }
  }
  operationStatus(transactionId: string) {
    const receipt = this.operationReceipts.get(transactionId)
    return receipt
      ? { status: 'applied' as const, transactionId, descriptor: receipt.commit.descriptor }
      : { status: 'unknown' as const, transactionId }
  }

  applyOperations(input: {
    transactionId: string
    expectedRevision: number
    operations: unknown
  }): Commit {
    const hash = createHash('sha256').update(JSON.stringify(input)).digest('hex')
    const previous = this.operationReceipts.get(input.transactionId)
    if (previous) {
      if (hash !== previous.hash)
        throw new Error('Transaction identity reused with different content')
      return structuredClone(previous.commit)
    }
    // Check before resolving block offsets, even if a stale request would be a no-op.
    if (input.expectedRevision !== this.revision) throw new Error('Stale document version')
    const overlays = structuredClone(this.root.overlays)
    const steps = operationSteps(this.root.doc, input.operations, {
      heading: (level) => {
        const existing = this.original.headingStyleIds.get(level)
        if (existing) return existing
        const pending = overlays.styles.find((style) => style.name === `heading ${level}`)
        if (pending) return pending.styleId
        let styleId = `GenOfficeHeading${level}`
        while (
          this.original.styles.has(styleId) ||
          overlays.styles.some((style) => style.styleId === styleId)
        )
          styleId += '_'
        overlays.styles.push({
          styleId,
          type: 'paragraph',
          name: `heading ${level}`,
          pPr: { outlineLevel: level, keepNext: true },
          rPr: { bold: true, sizeHalfPoints: Math.max(24, 36 - level * 2) },
        })
        return styleId
      },
      list: (kind) => {
        const pending = overlays.numbering.find((definition) => definition.kind === kind)
        if (pending) return pending.numId
        const numId = String(this.nextNumId++)
        overlays.numbering.push({ numId, kind })
        return numId
      },
    })
    const commit = this.submitSteps(
      {
        sessionId: this.sessionId,
        epoch: this.epoch,
        transactionId: input.transactionId,
        expectedRevision: input.expectedRevision,
        baseCollabVersion: this.collabVersion,
        clientId: 'host',
        steps,
      },
      overlays,
    )
    commit.presentation = this.presentation()
    this.operationReceipts.set(input.transactionId, { hash, commit: structuredClone(commit) })
    if (this.operationReceipts.size > 512)
      this.operationReceipts.delete(this.operationReceipts.keys().next().value!)
    return commit
  }

  submit(request: StepRequest): Commit {
    return this.submitSteps(request, this.root.overlays)
  }
  private submitSteps(request: StepRequest, overlays: Overlays): Commit {
    if (request.sessionId !== this.sessionId) throw new Error('Wrong document session')
    if (
      !request.transactionId ||
      request.transactionId.length > 200 ||
      !request.clientId ||
      request.clientId.length > 200
    )
      throw new Error('Invalid transaction identity')
    if (
      !Array.isArray(request.steps) ||
      request.steps.length > DOCX_LIMITS.steps ||
      jsonBytes(request.steps) > DOCX_LIMITS.stepBytes
    )
      throw new Error('Transaction exceeds limit')
    const hash = createHash('sha256').update(JSON.stringify(request)).digest('hex')
    const previous = this.receipts.get(request.transactionId)
    if (previous) {
      if (previous.hash !== hash)
        throw new Error('Transaction identity reused with different content')
      return structuredClone(previous.commit)
    }
    if (this.readonly) throw new Error('Document is read-only')
    if (
      request.epoch !== this.epoch ||
      request.expectedRevision !== this.revision ||
      request.baseCollabVersion !== this.collabVersion
    )
      throw new Error('Stale document version')
    const transform = new Transform(this.root.doc)
    for (const json of request.steps) transform.step(Step.fromJSON(docxSchema, json))
    const normalized = validateAndNormalize(
      this.content(),
      transform.doc.toJSON() as EditorJSON,
      randomUUID,
    )
    const canonical = docxSchema.nodeFromJSON(normalized)
    canonical.descendants((node) => {
      if (
        node.attrs.styleId &&
        !this.original.styles.has(node.attrs.styleId) &&
        !overlays.styles.some((style) => style.styleId === node.attrs.styleId)
      )
        throw new Error('Unknown paragraph style')
      if (
        node.type.name === 'docListItem' &&
        node.attrs.numId &&
        !this.original.numbering.has(node.attrs.numId) &&
        !overlays.numbering.some((definition) => definition.numId === node.attrs.numId)
      )
        throw new Error('Unknown numbering definition')
    })
    // Normalize identity attributes using small mapped steps, preserving selection
    // and allowing prosemirror-collab to rebase pending browser edits correctly.
    canonical.descendants((node, pos) => {
      const current = transform.doc.nodeAt(pos)
      if (current && JSON.stringify(current.attrs) !== JSON.stringify(node.attrs))
        transform.setNodeMarkup(pos, undefined, node.attrs)
    })
    if (!transform.doc.eq(canonical)) throw new Error('Normalization changed document structure')
    const changed = !this.root.doc.eq(canonical)
    const fromVersion = this.collabVersion
    if (changed) this.commitRoot(canonical, overlays)
    this.collabVersion += transform.steps.length
    const commit: Commit = {
      descriptor: this.descriptor(),
      fromVersion,
      changed,
      steps: transform.steps.map((step) => step.toJSON()),
      clientIds: transform.steps.map((_, index) =>
        index < request.steps.length ? request.clientId : 'host',
      ),
    }
    this.receipts.set(request.transactionId, { hash, commit: structuredClone(commit) })
    if (this.receipts.size > 512) this.receipts.delete(this.receipts.keys().next().value!)
    return commit
  }

  private commitRoot(doc: PMNode, overlays: Overlays) {
    const bytes = jsonBytes(doc.toJSON()) + jsonBytes(overlays)
    const retained = this.undoRoots.reduce((sum, root) => sum + root.bytes, this.root.bytes)
    if (retained + bytes > DOCX_LIMITS.history) throw new Error('Document history budget exceeded')
    this.undoRoots.push(this.root)
    this.redoRoots = []
    this.root = { doc, stateId: randomUUID(), bytes, overlays }
    this.revision++
  }

  /** The provider flushes and freezes the webview before native history calls. */
  history(direction: 'undo' | 'redo'): SessionDescriptor {
    if (this.readonly) throw new Error('Document is read-only')
    const source = direction === 'undo' ? this.undoRoots : this.redoRoots
    const target = direction === 'undo' ? this.redoRoots : this.undoRoots
    const root = source.pop()
    if (!root) throw new Error(`Nothing to ${direction}`)
    target.push(this.root)
    this.root = root
    this.revision++
    this.epoch++
    this.collabVersion = 0
    return this.descriptor()
  }

  async save(): Promise<Uint8Array> {
    const plan = editorToSavePlan(this.content(), this.original.blocks)
    return saveDocx(this.original, plan.saveBlocks, {
      newPictures: plan.newPictures,
      ...(this.root.overlays.numbering.length
        ? { numbering: { newDefs: this.root.overlays.numbering } }
        : {}),
      ...(this.root.overlays.styles.length ? { styleUpserts: this.root.overlays.styles } : {}),
    })
  }
}
