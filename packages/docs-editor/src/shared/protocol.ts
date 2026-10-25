import type { EditorJSON } from '../conversion'
import type { NumberingDef, StyleInfo, DocDefaults } from '@genoffice/docx-engine'
export interface DocumentPresentation {
  css: string
  numbering: Array<[string, NumberingDef]>
  styles: Array<[string, StyleInfo]>
  docDefaults?: DocDefaults
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
