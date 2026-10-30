import { createHash, randomUUID } from 'node:crypto'
import {
  parseDocx,
  saveDocx,
  type ParsedDocFull,
  numberingLevelFromCustom,
  readPageColor,
  previewFontSettings,
  type StyleInfo,
  readSectionSettings,
  type StyleUpsert,
  type NumberingDef,
} from '@genoffice/docx-engine'
import {
  HF_RAW_KEYS,
  commentsSchema,
  notesSchema,
  styleUpsertsSchema,
  documentSettingsSchema,
  pendingListDefinitionsSchema,
} from '@genoffice/docs-editor/shared/operations'
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
  DocumentSettings,
  PendingListDefinitions,
} from '@genoffice/docs-editor/shared/protocol'
export type { SessionDescriptor, StepRequest, Commit } from '@genoffice/docs-editor/shared/protocol'

interface Overlays {
  numbering: PendingListDefinitions['newDefs']
  restartNums: PendingListDefinitions['restartNums']
  styles: StyleUpsert[]
  settings: DocumentSettings
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
  /** Styles resolved for the pending style definitions (refreshed by stylesReady). */
  private liveStyles: { key: string; styles: Map<string, StyleInfo>; css: string } | null = null
  readonly readonly: boolean

  private constructor(private readonly original: ParsedDocFull) {
    this.readonly = Boolean(original.protection?.enforced || original.writeProtection?.hash)
    const json = prepareDocument(documentToEditor(original), randomUUID)
    if (jsonBytes(json) > DOCX_LIMITS.root) throw new Error('Document exceeds size limit')
    this.root = {
      doc: docxSchema.nodeFromJSON(json),
      stateId: randomUUID(),
      bytes: jsonBytes(json),
      overlays: { numbering: [], restartNums: [], styles: [], settings: {} },
    }
    this.savedStateId = this.root.stateId
    this.baseCSS = docStyleCss(original)
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
  /** Theme fonts / colors as edited in the Design tab, over the file's theme. */
  private themeCss() {
    const edits = this.root.overlays.settings
    return docThemeCss(
      edits.themeFonts ?? this.original.themeFonts,
      edits.themeColors ?? this.original.themeColors,
      Boolean(docBodyFont(this.original)),
    )
  }
  /** Resolve styles.xml with the pending definitions before presenting them. */
  async stylesReady(): Promise<void> {
    const upserts = this.root.overlays.styles
    const key = JSON.stringify(upserts)
    if (!upserts.length || this.liveStyles?.key === key) return
    const { styles } = await previewFontSettings(this.original, upserts)
    this.liveStyles = {
      key,
      styles,
      css: docStyleCss({ ...this.original, styles }),
    }
  }
  /** Style definitions join the overlay merged as the desktop's upsert does. */
  private withStyles(overlays: Overlays, input: unknown) {
    for (const up of styleUpsertsSchema.parse(input) as StyleUpsert[]) {
      const previous = overlays.styles.find((style) => style.styleId === up.styleId)
      overlays.styles = previous
        ? overlays.styles.map((style) =>
            style === previous
              ? {
                  ...previous,
                  ...up,
                  ...(up.pPr || previous.pPr ? { pPr: { ...previous.pPr, ...up.pPr } } : {}),
                  ...(up.rPr || previous.rPr ? { rPr: { ...previous.rPr, ...up.rPr } } : {}),
                }
              : style,
          )
        : [...overlays.styles, up]
    }
  }
  presentation(): DocumentPresentation {
    const numbering = new Map(this.original.numbering)
    for (const definition of this.root.overlays.numbering)
      numbering.set(definition.numId, {
        numId: definition.numId,
        abstractNumId: `pending-${definition.numId}`,
        levels: definition.levels
          ? Object.fromEntries(
              definition.levels.map((level, index) => [index, numberingLevelFromCustom(level)]),
            )
          : {
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
    for (const restart of this.root.overlays.restartNums) {
      const source = [...this.original.numbering.values()].find(
        (definition) => definition.abstractNumId === restart.abstractNumId,
      )
      if (source)
        numbering.set(restart.numId, {
          ...source,
          numId: restart.numId,
          startOverrides: restart.startOverrides,
        })
    }
    // Engine models contain optional undefined properties; explicitly project
    // trusted parsed data into JSON before the strict transport sees it.
    const live =
      this.liveStyles?.key === JSON.stringify(this.root.overlays.styles) ? this.liveStyles : null
    const pendingCSS = this.root.overlays.styles
      .map(
        (style) =>
          `.doc-page [data-style="${style.styleId}"] { font-weight:700; font-size:${(style.rPr?.sizeHalfPoints ?? 24) / 2}pt }`,
      )
      .join('\n')
    return JSON.parse(
      JSON.stringify({
        css: (live ? live.css : this.baseCSS) + this.themeCss() + pendingCSS,
        numbering: [...numbering.entries()],
        styles: [...(live ? live.styles : this.original.styles).entries()],
        docDefaults: this.original.docDefaults,
        settings: this.settings(),
      }),
    ) as DocumentPresentation
  }
  /** The document settings as the view shows them: session edits over the opened file. */
  private settings(): DocumentSettings {
    const edits = this.root.overlays.settings
    return {
      section: edits.section ?? readSectionSettings(this.original),
      pageColor: edits.pageColor !== undefined ? edits.pageColor : readPageColor(this.original),
      mirrorMargins: edits.mirrorMargins ?? this.original.mirrorMargins ?? false,
      watermark:
        edits.watermark !== undefined ? edits.watermark : (this.original.watermarkText ?? null),
      comments: edits.comments ?? this.original.comments,
      protection: edits.protection !== undefined ? edits.protection : this.original.protection,
      writeProtection:
        edits.writeProtection !== undefined ? edits.writeProtection : this.original.writeProtection,
      removePersonalInfo: edits.removePersonalInfo ?? this.original.removePersonalInfo,
      themeFonts: edits.themeFonts ?? this.original.themeFonts ?? null,
      themeColors: edits.themeColors ?? this.original.themeColors ?? null,
      footnotes: edits.footnotes ?? this.original.footnotes,
      endnotes: edits.endnotes ?? this.original.endnotes,
      header: edits.header ?? {
        text: this.original.headerText ?? '',
        ...(this.original.headerParas ? { paras: this.original.headerParas } : {}),
      },
      footer: edits.footer ?? {
        text: this.original.footerText ?? '',
        ...(this.original.footerParas ? { paras: this.original.footerParas } : {}),
      },
    }
  }
  /** Layout / Design settings edits: one history entry without document steps. */
  applySettings(input: {
    transactionId: string
    expectedRevision: number
    settings: unknown
  }): Commit {
    const hash = createHash('sha256').update(JSON.stringify(input)).digest('hex')
    const previous = this.operationReceipts.get(input.transactionId)
    if (previous) {
      if (hash !== previous.hash)
        throw new Error('Transaction identity reused with different content')
      return structuredClone(previous.commit)
    }
    if (this.readonly) throw new Error('Document is read-only')
    if (input.expectedRevision !== this.revision) throw new Error('Stale document version')
    const edit = documentSettingsSchema.parse(input.settings)
    // Verbatim run XML in a header/footer may only be the file's own.
    const fileRaw = new Set<string>()
    const collect = (value: unknown, key = '') => {
      if (HF_RAW_KEYS.has(key) && typeof value === 'string') fileRaw.add(value)
      else if (Array.isArray(value)) for (const item of value) collect(item)
      else if (value && typeof value === 'object')
        for (const [name, item] of Object.entries(value)) collect(item, name)
    }
    collect([this.original.headerParas, this.original.footerParas])
    const check = (value: unknown, key = '') => {
      if (HF_RAW_KEYS.has(key) && typeof value === 'string' && !fileRaw.has(value))
        throw new Error('Untrusted header or footer content')
      if (Array.isArray(value)) for (const item of value) check(item)
      else if (value && typeof value === 'object')
        for (const [name, item] of Object.entries(value)) check(item, name)
    }
    check([edit.header?.paras, edit.footer?.paras])
    const overlays = structuredClone(this.root.overlays)
    const { styles, ...settings } = edit
    if (styles) this.withStyles(overlays, styles)
    overlays.settings = { ...overlays.settings, ...settings } as DocumentSettings
    const fromVersion = this.collabVersion
    this.commitRoot(this.root.doc, overlays)
    const commit: Commit = {
      descriptor: this.descriptor(),
      fromVersion,
      changed: true,
      steps: [],
      clientIds: [],
      presentation: this.presentation(),
    }
    this.operationReceipts.set(input.transactionId, { hash, commit: structuredClone(commit) })
    if (this.operationReceipts.size > 512)
      this.operationReceipts.delete(this.operationReceipts.keys().next().value!)
    return commit
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
    if (
      !request.numbering &&
      !request.comments &&
      !request.footnotes &&
      !request.endnotes &&
      !request.styles
    )
      return this.submitSteps(request, this.root.overlays)
    const overlays = request.numbering
      ? this.withListDefinitions(request.numbering)
      : structuredClone(this.root.overlays)
    if (request.comments)
      overlays.settings = { ...overlays.settings, comments: commentsSchema.parse(request.comments) }
    if (request.styles) this.withStyles(overlays, request.styles)
    for (const kind of ['footnotes', 'endnotes'] as const)
      if (request[kind])
        overlays.settings = {
          ...overlays.settings,
          [kind]: notesSchema.parse(request[kind]),
        } as DocumentSettings
    const commit = this.submitSteps(request, overlays)
    commit.presentation = this.presentation()
    return commit
  }
  /** Ribbon list definitions join the overlay with the steps that use them (one history entry). */
  private withListDefinitions(input: unknown): Overlays {
    const pending = pendingListDefinitionsSchema.parse(input)
    const overlays = structuredClone(this.root.overlays)
    const taken = (numId: string) =>
      this.original.numbering.has(numId) ||
      overlays.numbering.some((definition) => definition.numId === numId) ||
      overlays.restartNums.some((restart) => restart.numId === numId)
    for (const definition of pending.newDefs) {
      const existing = overlays.numbering.find((item) => item.numId === definition.numId)
      if (existing) {
        if (JSON.stringify(existing) !== JSON.stringify(definition))
          throw new Error('List definition identity reused')
        continue
      }
      if (taken(definition.numId)) throw new Error('List definition identity reused')
      overlays.numbering.push(definition)
    }
    for (const restart of pending.restartNums) {
      const existing = overlays.restartNums.find((item) => item.numId === restart.numId)
      if (existing) {
        if (JSON.stringify(existing) !== JSON.stringify(restart))
          throw new Error('List definition identity reused')
        continue
      }
      if (taken(restart.numId)) throw new Error('List definition identity reused')
      if (
        ![...this.original.numbering.values()].some(
          (definition) => definition.abstractNumId === restart.abstractNumId,
        )
      )
        throw new Error('Unknown list definition')
      overlays.restartNums.push(restart)
    }
    for (const numId of [...overlays.numbering, ...overlays.restartNums].map((item) => item.numId))
      this.nextNumId = Math.max(this.nextNumId, Number(numId) + 1)
    return overlays
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
        !overlays.numbering.some((definition) => definition.numId === node.attrs.numId) &&
        !overlays.restartNums.some((restart) => restart.numId === node.attrs.numId)
      )
        throw new Error('Unknown numbering definition')
    })
    // Comment anchors must name comments of the resulting list.
    const commentIds = new Set(
      (overlays.settings.comments ?? this.original.comments).map((comment) => comment.id),
    )
    const noteIds = {
      footnote: new Set((overlays.settings.footnotes ?? this.original.footnotes).map((n) => n.id)),
      endnote: new Set((overlays.settings.endnotes ?? this.original.endnotes).map((n) => n.id)),
    }
    canonical.descendants((node) => {
      // Note references must name notes of the resulting lists.
      if (
        node.type.name === 'docNoteRef' &&
        !noteIds[node.attrs.kind === 'endnote' ? 'endnote' : 'footnote'].has(String(node.attrs.id))
      )
        throw new Error('Unknown note reference')
      for (const mark of node.marks)
        if (mark.type.name === 'comment')
          for (const id of String(mark.attrs.ids ?? '').split(' '))
            if (id && !commentIds.has(id)) throw new Error('Unknown comment anchor')
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
      ...(this.root.overlays.numbering.length || this.root.overlays.restartNums.length
        ? {
            numbering: {
              newDefs: this.root.overlays.numbering,
              restartNums: this.root.overlays.restartNums,
            },
          }
        : {}),
      ...(this.root.overlays.styles.length ? { styleUpserts: this.root.overlays.styles } : {}),
      ...(this.root.overlays.settings.section
        ? { section: this.root.overlays.settings.section }
        : {}),
      ...(this.root.overlays.settings.pageColor !== undefined
        ? { pageColor: this.root.overlays.settings.pageColor }
        : {}),
      ...(this.root.overlays.settings.mirrorMargins !== undefined
        ? { mirrorMargins: this.root.overlays.settings.mirrorMargins }
        : {}),
      ...(this.root.overlays.settings.comments
        ? { comments: this.root.overlays.settings.comments }
        : {}),
      ...(this.root.overlays.settings.protection !== undefined
        ? { protection: this.root.overlays.settings.protection }
        : {}),
      ...(this.root.overlays.settings.writeProtection !== undefined
        ? { writeProtection: this.root.overlays.settings.writeProtection }
        : {}),
      ...(this.root.overlays.settings.removePersonalInfo !== undefined
        ? { removePersonalInfo: this.root.overlays.settings.removePersonalInfo }
        : {}),
      ...(this.root.overlays.settings.themeFonts
        ? { themeFonts: this.root.overlays.settings.themeFonts }
        : {}),
      ...(this.root.overlays.settings.themeColors
        ? { themeColors: this.root.overlays.settings.themeColors }
        : {}),
      ...(this.root.overlays.settings.footnotes
        ? { footnotes: this.root.overlays.settings.footnotes }
        : {}),
      ...(this.root.overlays.settings.endnotes
        ? { endnotes: this.root.overlays.settings.endnotes }
        : {}),
      ...(this.root.overlays.settings.header ? { header: this.root.overlays.settings.header } : {}),
      ...(this.root.overlays.settings.footer ? { footer: this.root.overlays.settings.footer } : {}),
      ...(this.root.overlays.settings.watermark !== undefined
        ? { watermark: this.root.overlays.settings.watermark }
        : {}),
    })
  }
}
