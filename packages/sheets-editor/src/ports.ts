export interface GridRange {
  readonly startRow: number
  readonly endRow: number
  readonly startColumn: number
  readonly endColumn: number
}

/** Browser ports contain data only; no extension or Node runtime handles. */
export interface EditorHost {
  action?(
    request: import('@genoffice/xlsx-gateway/shared/office-actions').OfficeAction,
  ): Promise<void>
  read(sheetId: string, range: GridRange, stateId: string): Promise<unknown>
  submit(intent: {
    transactionId: string
    expectedRevision: number
    operations: readonly unknown[]
  }): Promise<unknown>
  selection(sheetId: string, range: GridRange | null): void
  history(direction: 'undo' | 'redo'): Promise<void>
}

export interface CandidateResult {
  readonly beforeStateId: string
  readonly nextState: unknown
  readonly forward: unknown
  readonly inverse: unknown
  readonly affectedTargets: readonly string[]
}

export interface EditorProjection {
  install(stateId: string, revision: number, data: unknown): Promise<void>
  flushEditor(): Promise<boolean>
  dispose(): void
}
