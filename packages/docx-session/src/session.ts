import { createHash, randomUUID } from 'node:crypto'
import {
  parseDocx,
  saveDocx,
  type ParsedDocFull,
  numberingLevelFromCustom,
  readPageColor,
  previewFontSettings,
  type StyleInfo,
  type HeaderFooter,
  readSectionSettings,
  readSections,
  applySectionSettings,
  applySectionStartType,
  applyTitlePg,
  applyPageNumType,
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
import {
  annotationsFromParsed,
  buildInkImages,
  type InkStroke,
} from '../../../apps/docs/src/renderer/editor/ink'
import {
  applySectPrRewrites,
  type SectPrRewrite,
} from '../../../apps/docs/src/renderer/sectpr-rewrite'
import {
  hfLinked,
  hfSaveOptions,
  resolveHf,
  withHfEdit,
  withHfLink,
  type HfSectionState,
} from '../../../apps/docs/src/renderer/hf-sections'
import { hfFromPart, hfVariantsFromParsed } from '../../../apps/docs/src/renderer/doc-state'
import { readBlocks, type SelectionScope } from './read'
import type {
  SessionDescriptor,
  StepRequest,
  Commit,
  DocumentPresentation,
  DocumentSettings,
  PendingListDefinitions,
  SectionEntry,
  SectionStrip,
} from '@genoffice/docs-editor/shared/protocol'
export type { SessionDescriptor, StepRequest, Commit } from '@genoffice/docs-editor/shared/protocol'

interface Overlays {
  numbering: PendingListDefinitions['newDefs']
  restartNums: PendingListDefinitions['restartNums']
  levelEdits: NonNullable<PendingListDefinitions['levelEdits']>
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
      overlays: { numbering: [], restartNums: [], levelEdits: [], styles: [], settings: {} },
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
    for (const edit of this.root.overlays.levelEdits)
      for (const [numId, definition] of numbering)
        if (definition.abstractNumId === edit.abstractNumId)
          numbering.set(numId, {
            ...definition,
            levels: { ...definition.levels, [edit.ilvl]: numberingLevelFromCustom(edit.level) },
          })
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
  /** The header/footer state of every section: the file's parts with the session's pending edits. */
  private hfSectionState(pending = this.root.overlays.settings.sectionHfState): HfSectionState {
    return {
      sections: readSections(this.original),
      edits: pending?.edits ?? {},
      links: pending?.links ?? {},
      docStrips: {
        header: hfFromPart({
          text: this.original.headerText ?? '',
          hasPageNumber: !!this.original.headerHasPageNumber,
          paras: this.original.headerParas ?? [],
        }),
        footer: hfFromPart({
          text: this.original.footerText ?? '',
          hasPageNumber: !!this.original.footerHasPageNumber,
          paras: this.original.footerParas ?? [],
        }),
        variants: hfVariantsFromParsed(this.original),
      },
      parsed: this.original,
    }
  }
  /** The strip each section shows (header and footer, default/first/even), following Word's inheritance. */
  private sectionStrips(): SectionStrip[] {
    const state = this.hfSectionState()
    if (state.sections.length < 2) return []
    const strips: SectionStrip[] = []
    for (let section = 0; section < state.sections.length; section++)
      for (const kind of ['header', 'footer'] as const)
        for (const variant of ['default', 'first', 'even'] as const) {
          const resolved = resolveHf(state, section, kind, variant)
          strips.push({
            section,
            kind,
            variant,
            value: resolved.value,
            owner: resolved.owner,
            linked: hfLinked(state, section, kind, variant),
          })
        }
    return strips
  }
  /** Every section of the file with the session's edits applied. */
  private sectionEntries(settings: DocumentSettings): SectionEntry[] {
    const edits = this.root.overlays.settings.sectionEdits ?? {}
    const sections = readSections(this.original)
    return sections.map((sec, index) => {
      const final = index === sections.length - 1
      const edit = final ? undefined : edits[String(sec.lastBlockIndex)]
      const type = edit?.startType ?? sec.startType
      return {
        blockIndex: sec.lastBlockIndex,
        startType: type === 'nextColumn' ? 'nextPage' : type,
        section: final ? (settings.section ?? sec.settings) : (edit?.section ?? sec.settings),
        titlePg: final ? (settings.titlePg ?? sec.titlePg) : (edit?.titlePg ?? sec.titlePg),
        pageNumbering: final
          ? (settings.pageNumbering ?? {})
          : (edit?.pageNumbering ?? {
              ...(sec.pageNumberFmt ? { fmt: sec.pageNumberFmt } : {}),
              ...(sec.pageNumberStart !== undefined ? { start: sec.pageNumberStart } : {}),
            }),
        final,
      }
    })
  }
  /** w:pgNumType of the final section as opened. */
  private originalPageNumbering(): NonNullable<DocumentSettings['pageNumbering']> {
    const last = readSections(this.original).at(-1)
    return {
      ...(last?.pageNumberFmt ? { fmt: last.pageNumberFmt } : {}),
      ...(last?.pageNumberStart !== undefined ? { start: last.pageNumberStart } : {}),
    }
  }
  /** The document settings as the view shows them: session edits over the opened file. */
  private settings(): DocumentSettings {
    const edits = this.root.overlays.settings
    const base: DocumentSettings = {
      section: edits.section ?? readSectionSettings(this.original),
      pageColor: edits.pageColor !== undefined ? edits.pageColor : readPageColor(this.original),
      mirrorMargins: edits.mirrorMargins ?? this.original.mirrorMargins ?? false,
      pageNumbering: edits.pageNumbering ?? this.originalPageNumbering(),
      ...(edits.sectionStart ? { sectionStart: edits.sectionStart } : {}),
      titlePg: edits.titlePg ?? this.original.titlePg ?? false,
      evenAndOddHeaders: edits.evenAndOddHeaders ?? this.original.evenAndOddHeaders ?? false,
      watermark:
        edits.watermark !== undefined ? edits.watermark : (this.original.watermarkText ?? null),
      comments: edits.comments ?? this.original.comments,
      inks: edits.inks ?? annotationsFromParsed(this.original.inks ?? []),
      sources: edits.sources ?? this.original.sources,
      protection: edits.protection !== undefined ? edits.protection : this.original.protection,
      writeProtection:
        edits.writeProtection !== undefined ? edits.writeProtection : this.original.writeProtection,
      removePersonalInfo: edits.removePersonalInfo ?? this.original.removePersonalInfo,
      themeFonts: edits.themeFonts ?? this.original.themeFonts ?? null,
      themeColors: edits.themeColors ?? this.original.themeColors ?? null,
      footnotes: edits.footnotes ?? this.original.footnotes,
      endnotes: edits.endnotes ?? this.original.endnotes,
      ...Object.fromEntries(
        (['headerFirst', 'footerFirst', 'headerEven', 'footerEven'] as const).flatMap((key) => {
          const value = edits[key] ?? this.original[key]
          return value ? [[key, { text: value.text, paras: value.paras }]] : []
        }),
      ),
      header: edits.header ?? {
        text: this.original.headerText ?? '',
        ...(this.original.headerParas ? { paras: this.original.headerParas } : {}),
      },
      footer: edits.footer ?? {
        text: this.original.footerText ?? '',
        ...(this.original.footerParas ? { paras: this.original.footerParas } : {}),
      },
    }
    return {
      ...base,
      sections: this.sectionEntries(base),
      sectionStrips: this.sectionStrips(),
      ...(edits.sectionEdits ? { sectionEdits: edits.sectionEdits } : {}),
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
    // Section edits address the closing paragraph of a non-final section of the opened file.
    if (edit.sectionEdits) {
      const sections = readSections(this.original)
      const editable = new Set(sections.slice(0, -1).map((sec) => String(sec.lastBlockIndex)))
      for (const key of Object.keys(edit.sectionEdits))
        if (!editable.has(key)) throw new Error('Unknown section')
    }
    // Ink anchors to paragraphs of the opened file (the save re-anchors by docxIndex).
    if (edit.inks) {
      const anchors = new Set(this.original.blocks.map((block) => block.docxIndex))
      for (const ink of edit.inks)
        if (!anchors.has(ink.anchorIndex)) throw new Error('Unknown ink anchor')
    }
    // Verbatim run XML in a header/footer may only be the file's own.
    const fileRaw = new Set<string>()
    const collect = (value: unknown, key = '') => {
      if (HF_RAW_KEYS.has(key) && typeof value === 'string') fileRaw.add(value)
      else if (Array.isArray(value)) for (const item of value) collect(item)
      else if (value && typeof value === 'object')
        for (const [name, item] of Object.entries(value)) collect(item, name)
    }
    collect([
      ...Object.values(this.original.hfParts ?? {}).map((part) => part.paras),
      this.original.headerParas,
      this.original.footerParas,
      ...[
        this.original.headerFirst,
        this.original.footerFirst,
        this.original.headerEven,
        this.original.footerEven,
      ].map((part) => part?.paras),
    ])
    const check = (value: unknown, key = '') => {
      if (HF_RAW_KEYS.has(key) && typeof value === 'string' && !fileRaw.has(value))
        throw new Error('Untrusted header or footer content')
      if (Array.isArray(value)) for (const item of value) check(item)
      else if (value && typeof value === 'object')
        for (const [name, item] of Object.entries(value)) check(item, name)
    }
    check([
      edit.sectionHfEdit?.hf.paras,
      edit.header?.paras,
      edit.footer?.paras,
      edit.headerFirst?.paras,
      edit.footerFirst?.paras,
      edit.headerEven?.paras,
      edit.footerEven?.paras,
    ])
    const overlays = structuredClone(this.root.overlays)
    const { styles, numbering, sectionHfEdit, sectionHfLink, ...settings } = edit
    if (styles) this.withStyles(overlays, styles)
    if (numbering) {
      const merged = this.withListDefinitions(numbering, overlays)
      overlays.numbering = merged.numbering
      overlays.restartNums = merged.restartNums
      overlays.levelEdits = merged.levelEdits
    }
    overlays.settings = { ...overlays.settings, ...settings } as DocumentSettings
    if (sectionHfEdit || sectionHfLink) {
      const count = readSections(this.original).length
      if ((sectionHfEdit ?? sectionHfLink)!.section >= count)
        throw new Error('Unknown section strip')
      const state = this.hfSectionState(overlays.settings.sectionHfState)
      if (sectionHfEdit) {
        const next = withHfEdit(
          state,
          sectionHfEdit.section,
          sectionHfEdit.kind,
          sectionHfEdit.variant,
          sectionHfEdit.hf as HeaderFooter,
        )
        if (!next) throw new Error('Unknown section strip')
        overlays.settings.sectionHfState = { edits: next.edits, links: next.links }
      }
      if (sectionHfLink) {
        const next = withHfLink(
          this.hfSectionState(overlays.settings.sectionHfState),
          sectionHfLink.section,
          sectionHfLink.kind,
          sectionHfLink.variant,
          sectionHfLink.on,
        )
        if (!next) throw new Error('Unknown section strip')
        overlays.settings.sectionHfState = { edits: next.edits, links: next.links }
      }
    }
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
  private withListDefinitions(input: unknown, base: Overlays = this.root.overlays): Overlays {
    const pending = pendingListDefinitionsSchema.parse(input)
    const overlays = structuredClone(base)
    const taken = (numId: string) =>
      this.original.numbering.has(numId) ||
      overlays.numbering.some((definition) => definition.numId === numId) ||
      overlays.restartNums.some((restart) => restart.numId === numId)
    for (const definition of pending.newDefs) {
      const existing = overlays.numbering.find((item) => item.numId === definition.numId)
      if (existing) {
        // The view refines a definition it created this session (levels edited).
        if (existing.kind !== definition.kind) throw new Error('List definition identity reused')
        overlays.numbering = overlays.numbering.map((item) =>
          item === existing ? definition : item,
        )
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
    for (const edit of pending.levelEdits ?? []) {
      if (edit.abstractNumId.startsWith('pending-')) {
        const numId = edit.abstractNumId.slice('pending-'.length)
        const definition = overlays.numbering.find((item) => item.numId === numId)
        if (!definition) throw new Error('Unknown list definition')
        const levels = [...(definition.levels ?? [])]
        while (levels.length <= edit.ilvl)
          levels.push({
            numFmt: definition.kind === 'bullet' ? 'bullet' : 'decimal',
            lvlText: definition.kind === 'bullet' ? '•' : `%${levels.length + 1}.`,
            indentLeft: 720 * (levels.length + 1),
            hanging: 360,
          })
        levels[edit.ilvl] = edit.level
        overlays.numbering = overlays.numbering.map((item) =>
          item === definition ? { ...item, levels } : item,
        )
        continue
      }
      if (
        ![...this.original.numbering.values()].some(
          (definition) => definition.abstractNumId === edit.abstractNumId,
        )
      )
        throw new Error('Unknown list definition')
      overlays.levelEdits = [
        ...overlays.levelEdits.filter(
          (item) => !(item.abstractNumId === edit.abstractNumId && item.ilvl === edit.ilvl),
        ),
        edit,
      ]
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
    // The view rasterizes drawn strokes; restored file inks carry their own PNG.
    const inkSettings = this.root.overlays.settings.inks
    const rasters = new Map<
      InkStroke,
      NonNullable<NonNullable<typeof inkSettings>[number]['raster']>
    >()
    for (const ink of inkSettings ?? [])
      if (ink.stroke && ink.raster) rasters.set(ink.stroke, ink.raster)
    const inks = inkSettings
      ? buildInkImages(
          inkSettings.filter((ink) => !ink.stroke || ink.raster),
          plan.saveBlockIndexByDocx,
          (stroke) => rasters.get(stroke)!,
        )
      : this.original.inks?.length
        ? buildInkImages(annotationsFromParsed(this.original.inks), plan.saveBlockIndexByDocx)
        : undefined
    // Page setup of non-final sections lives in their closing paragraphs' sectPr.
    const rewrites = new Map<number, SectPrRewrite>()
    const sectionEdits = this.root.overlays.settings.sectionEdits ?? {}
    const originalSections = readSections(this.original)
    for (const [key, edit] of Object.entries(sectionEdits)) {
      const sec = originalSections.slice(0, -1).find((item) => String(item.lastBlockIndex) === key)
      const block =
        sec && this.original.blocks.find((item) => item.docxIndex === sec.lastBlockIndex)
      if (!sec || !block?.originalXml || !sec.sectPrXml) continue
      let sectPr = sec.sectPrXml
      if (edit.section) sectPr = applySectionSettings(sectPr, edit.section)
      if (edit.startType) sectPr = applySectionStartType(sectPr, edit.startType)
      if (edit.titlePg !== undefined) sectPr = applyTitlePg(sectPr, edit.titlePg)
      if (edit.pageNumbering)
        sectPr = applyPageNumType(sectPr, edit.pageNumbering.fmt, edit.pageNumbering.start)
      rewrites.set(sec.lastBlockIndex, {
        from: sec.sectPrXml,
        to: sectPr,
        originalXml: block.originalXml,
      })
    }
    const saveBlocks = rewrites.size
      ? applySectPrRewrites(plan.saveBlocks, plan.saveBlockIndexByDocx, rewrites)
      : plan.saveBlocks
    return saveDocx(this.original, saveBlocks, {
      newPictures: plan.newPictures,
      ...(this.root.overlays.numbering.length ||
      this.root.overlays.restartNums.length ||
      this.root.overlays.levelEdits.length
        ? {
            numbering: {
              newDefs: this.root.overlays.numbering,
              restartNums: this.root.overlays.restartNums,
              levelEdits: this.root.overlays.levelEdits,
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
      ...(this.root.overlays.settings.sectionStart
        ? { sectionStartType: this.root.overlays.settings.sectionStart }
        : {}),
      ...(this.root.overlays.settings.pageNumbering
        ? { pgNumType: this.root.overlays.settings.pageNumbering }
        : {}),
      ...(this.root.overlays.settings.titlePg !== undefined
        ? { titlePg: this.root.overlays.settings.titlePg }
        : {}),
      ...(this.root.overlays.settings.evenAndOddHeaders !== undefined
        ? { evenAndOddHeaders: this.root.overlays.settings.evenAndOddHeaders }
        : {}),
      ...(this.root.overlays.settings.comments
        ? { comments: this.root.overlays.settings.comments }
        : {}),
      ...(inks ? { inks } : {}),
      ...(this.root.overlays.settings.sources
        ? { sources: this.root.overlays.settings.sources }
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
      ...Object.fromEntries(
        (['headerFirst', 'footerFirst', 'headerEven', 'footerEven'] as const).flatMap((key) =>
          this.root.overlays.settings[key] ? [[key, this.root.overlays.settings[key]]] : [],
        ),
      ),
      ...(this.root.overlays.settings.sectionHfState ? hfSaveOptions(this.hfSectionState()) : {}),
      ...(this.root.overlays.settings.watermark !== undefined
        ? { watermark: this.root.overlays.settings.watermark }
        : {}),
    })
  }
}
