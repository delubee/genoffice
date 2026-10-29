export interface GridRange {
  readonly startRow: number
  readonly endRow: number
  readonly startColumn: number
  readonly endColumn: number
}

/** Browser ports contain data only; no extension or Node runtime handles. */
export interface EditorHost {
  outline?(sheetId: string, stateId: string): Promise<unknown>
  goalSeek?(
    request: import('@genoffice/xlsx-gateway/shared/session-goal-seek').GoalSeekRequest,
  ): Promise<unknown>

  audit?(
    request: import('@genoffice/xlsx-gateway/shared/session-audit').WorkbookAuditRequest,
  ): Promise<unknown>
  media?(
    visualId: string,
    stateId: string,
  ): Promise<import('@genoffice/xlsx-gateway/shared/workbook-protocol').WorkbookMediaResult>
  search?(
    request: import('@genoffice/xlsx-gateway/shared/session-search').SessionSearchRequest,
  ): Promise<unknown>
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
  flushEditor(options?: { commitChartDraft?: boolean }): Promise<boolean>
  dispose(): void
}
