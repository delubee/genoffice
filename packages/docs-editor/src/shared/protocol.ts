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
/** Document settings the Layout and Design tabs edit (final section, page color, watermark). */
export interface DocumentSettings {
  section?: SectionSettings
  pageColor?: string | null
  mirrorMargins?: boolean
  /** Header & Footer → Different First Page / Different Odd & Even Pages */
  titlePg?: boolean
  /** Page Number Format of the (final) section: w:pgNumType fmt / start */
  pageNumbering?: { fmt?: string; start?: number }
  evenAndOddHeaders?: boolean
  watermark?: string | null
  /** default page header / footer of the document */
  header?: HeaderFooter
  footer?: HeaderFooter
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
