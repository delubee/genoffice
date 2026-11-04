import type { EditorJSON } from '../conversion'
import type {
  CommentInfo,
  CustomNumberingLevel,
  HeaderFooter,
  NoteInfo,
  StyleUpsert,
  NumberingDef,
  SectionSettings,
  SourceInfo,
  DocProtection,
  WriteProtection,
  ThemeColors,
  ThemeFonts,
  StyleInfo,
  DocDefaults,
} from '@genoffice/docx-engine'
/** Draw → ink annotation; `raster` is the view's rendering of a drawn stroke for the save. */
export interface InkSetting {
  id: string
  anchorIndex: number
  stroke?: {
    tool: 'pen' | 'highlighter'
    color: string
    width: number
    points: { x: number; y: number }[]
  }
  image?: {
    dataUrl: string
    offsetXPx: number
    offsetYPx: number
    widthPx: number
    heightPx: number
  }
  raster?: {
    base64: string
    widthPx: number
    heightPx: number
    offsetXPx: number
    offsetYPx: number
  }
}
/** One section as the view shows it (the final section's settings also appear at the top level). */
export interface SectionEntry {
  /** docxIndex of the paragraph that closes the section (the key of its edits) */
  blockIndex: number
  /** how the section starts */
  startType: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage'
  section: SectionSettings
  titlePg: boolean
  pageNumbering: { fmt?: string; start?: number }
  final: boolean
}
/** Edits to a non-final section of the opened file, keyed by the docxIndex of its closing paragraph. */
export interface SectionEdit {
  section?: SectionSettings
  startType?: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage'
  titlePg?: boolean
  pageNumbering?: { fmt?: string; start?: number }
}
export type HfPartKind = 'header' | 'footer'
export type HfPartVariant = 'default' | 'first' | 'even'
/** The strip a section shows (read-only output; present when the file has several sections). */
export interface SectionStrip {
  section: number
  kind: HfPartKind
  variant: HfPartVariant
  value: HeaderFooter | null
  /** the section whose part is shown (an inherited strip is owned by an earlier section) */
  owner: number
  /** Word's Link to Previous: null for the first section */
  linked: boolean | null
}
/** Document settings the Layout and Design tabs edit (final section, page color, watermark). */
export interface DocumentSettings {
  section?: SectionSettings
  /** every section of the document (read-only output) */
  sections?: SectionEntry[]
  /** edits to non-final sections */
  sectionEdits?: Record<string, SectionEdit>
  /** the header/footer strips of every section (read-only output) */
  sectionStrips?: SectionStrip[]
  /** write a strip of one section (a command: the session keeps the resulting per-part state) */
  sectionHfEdit?: { section: number; kind: HfPartKind; variant: HfPartVariant; hf: HeaderFooter }
  /** Link to Previous on or off for one section's strip (a command) */
  sectionHfLink?: { section: number; kind: HfPartKind; variant: HfPartVariant; on: boolean }
  /** session-held pending parts and links of the sections (set by the two commands above) */
  sectionHfState?: { edits: Record<string, HeaderFooter>; links: Record<string, true> }
  pageColor?: string | null
  mirrorMargins?: boolean
  /** Header & Footer → Different First Page / Different Odd & Even Pages */
  titlePg?: boolean
  /** how the final section starts (the type of the last Layout → Breaks entry) */
  sectionStart?: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage'
  /** Page Number Format of the (final) section: w:pgNumType fmt / start */
  pageNumbering?: { fmt?: string; start?: number }
  evenAndOddHeaders?: boolean
  watermark?: string | null
  /** default page header / footer of the document */
  header?: HeaderFooter
  footer?: HeaderFooter
  /** first-page and even-page parts (Header & Footer → Different First Page / Odd & Even) */
  headerFirst?: HeaderFooter
  footerFirst?: HeaderFooter
  headerEven?: HeaderFooter
  footerEven?: HeaderFooter
  /** the full comment list (threads, replies, resolved state) */
  comments?: CommentInfo[]
  footnotes?: NoteInfo[]
  endnotes?: NoteInfo[]
  /** Draw → ink annotations over the page */
  inks?: InkSetting[]
  /** References → Manage Sources (bibliography sources) */
  sources?: SourceInfo[]
  /** Review → Protect Document (applies to the saved file) */
  protection?: DocProtection | null
  writeProtection?: WriteProtection | null
  removePersonalInfo?: boolean
  /** Design → Fonts / Colors (null: the document has no theme part) */
  themeFonts?: ThemeFonts | null
  themeColors?: ThemeColors | null
  /** style definition edits (merged into the session's pending styles) */
  styles?: StyleUpsert[]
  /** list definition edits without document steps */
  numbering?: PendingListDefinitions
}
/** List definitions the view allocated for these steps (Ribbon bullets / numbering / multilevel). */
export interface PendingListDefinitions {
  newDefs: Array<{ numId: string; kind: 'bullet' | 'ordered'; levels?: CustomNumberingLevel[] }>
  restartNums: Array<{
    numId: string
    abstractNumId: string
    startOverrides: Record<number, number>
  }>
  /** one level of an existing (or session) list rewritten: Adjust List Indents */
  levelEdits?: Array<{ abstractNumId: string; ilvl: number; level: CustomNumberingLevel }>
}
export interface DocumentPresentation {
  css: string
  numbering: Array<[string, NumberingDef]>
  styles: Array<[string, StyleInfo]>
  docDefaults?: DocDefaults
  settings?: DocumentSettings
}
export interface SessionDescriptor {
  schemaVersion: 1
  sessionId: string
  stateId: string
  revision: number
  collabVersion: number
  epoch: number
  readonly: boolean
  dirty: boolean
  conflict: boolean
}
export interface StepRequest {
  sessionId: string
  epoch: number
  transactionId: string
  expectedRevision: number
  baseCollabVersion: number
  clientId: string
  steps: unknown[]
  numbering?: PendingListDefinitions
  /** the comment list after these steps, when they add or remove comment anchors */
  comments?: CommentInfo[]
  footnotes?: NoteInfo[]
  endnotes?: NoteInfo[]
  /** style definitions (Styles pane / gallery) the steps use */
  styles?: StyleUpsert[]
}
export interface Commit {
  descriptor: SessionDescriptor
  fromVersion: number
  steps: unknown[]
  clientIds: string[]
  changed: boolean
  presentation?: DocumentPresentation
}
export interface DocumentProjection {
  descriptor: SessionDescriptor
  content: EditorJSON
  presentation?: DocumentPresentation
}
