import type { EditorJSON } from '../conversion'
import type {
  CustomNumberingLevel,
  NumberingDef,
  SectionSettings,
  StyleInfo,
  DocDefaults,
} from '@genoffice/docx-engine'
/** Document settings the Layout and Design tabs edit (final section, page color, watermark). */
export interface DocumentSettings {
  section?: SectionSettings
  pageColor?: string | null
  mirrorMargins?: boolean
  watermark?: string | null
}
/** List definitions the view allocated for these steps (Ribbon bullets / numbering / multilevel). */
export interface PendingListDefinitions {
  newDefs: Array<{ numId: string; kind: 'bullet' | 'ordered'; levels?: CustomNumberingLevel[] }>
  restartNums: Array<{
    numId: string
    abstractNumId: string
    startOverrides: Record<number, number>
  }>
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
