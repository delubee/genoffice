import { readCurrentOutline } from '@genoffice/workbook-session/read-outline'
import {
  outlineReadRequestSchema,
  outlineReadResultSchema,
} from '@genoffice/xlsx-gateway/domain/session-outline'
import { WorkbookGoalSeekController } from '@genoffice/workbook-session/goal-seek-controller'
import { auditWorkbook } from '@genoffice/workbook-session/audit'
import { readCurrentMedia } from '@genoffice/workbook-session/media'
import { randomUUID } from 'node:crypto'
import { WorkbookSearchController } from '@genoffice/workbook-session/search-controller'
import type { Webview } from 'vscode'
import { z } from 'zod'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import { JsonBridge, BridgeError } from '@genoffice/xlsx-gateway/shared/session-bridge'
import { prepareCandidateMaterial } from '@genoffice/workbook-session/candidate-material'
import { readCurrentRange } from '@genoffice/workbook-session/read'
import { workbookStateSchema, transactionInputSchema } from '@genoffice/workbook-session/schemas'
import type {
  DocumentDescriptor,
  WorkbookState,
  OriginalBase,
  PreparedChangeset,
  TransactionInput,
} from '@genoffice/workbook-session/domain'
import type { CandidateInterpreter, WorkbookSource } from '@genoffice/workbook-session/ports'

const sourceReadSchema = z.strictObject({
  stateId: z.string().min(1),
  logicalSheetId: z.string().min(1),
  range: z.string().regex(/^[A-Z]{1,3}[1-9][0-9]{0,6}(:[A-Z]{1,3}[1-9][0-9]{0,6})?$/),
})
const candidateResultSchema = z.strictObject({
  beforeStateId: z.string(),
  nextState: workbookStateSchema,
  forward: z.json(),
  inverse: z.json(),
  affectedTargets: z.array(z.string()),
  calculationState: z.enum(['ready', 'pending', 'error']),
})

/** Local URI/source handles stay on this side of the validated JSON bridge. */
export class WebviewCandidateBridge implements CandidateInterpreter {
  readonly bridge: JsonBridge
  private readonly goalSeek: WorkbookGoalSeekController
  private readonly search: WorkbookSearchController
  private notifyReady!: () => void
  readonly ready = new Promise<void>((resolve) => {
    this.notifyReady = resolve
  })
  constructor(
    private readonly input: {
      webview: Webview
      sessionId: string
      original: OriginalBase
      source: WorkbookSource
      getState(): WorkbookState
      getDescriptor(): DocumentDescriptor
      submit(intent: TransactionInput): Promise<unknown>
      submitStaged?(intent: Omit<TransactionInput, 'operations'>, stages: unknown): Promise<unknown>
      history?(direction: 'undo' | 'redo'): Promise<void>
    },
  ) {
    this.goalSeek = new WorkbookGoalSeekController({
      sessionId: input.sessionId,
      sourceSessionId: input.original.sidecarSessionId,
      source: input.source,
      getState: input.getState,
      getDescriptor: input.getDescriptor,
      submit: input.submit,
    })
    this.search = new WorkbookSearchController({
      source: input.source,
      sourceSessionId: input.original.sidecarSessionId,
      getState: () => input.getState(),
      replace: async (stateId, operations) => {
        const descriptor = input.getDescriptor()
        if (input.getState().stateId !== stateId || descriptor.stateId !== stateId)
          throw new BridgeError('revision_conflict', 'The document changed before replacement.')
        return input.submit({
          schemaVersion: 1,
          sessionId: input.sessionId,
          transactionId: randomUUID(),
          expectedRevision: descriptor.revision,
          origin: 'manual',
          summary: 'Replace search results',
          operations,
        })
      },
    })
    this.bridge = new JsonBridge(input.sessionId, {
      post: (message) => input.webview.postMessage(message),
      subscribe: (listener) => {
        const subscription = input.webview.onDidReceiveMessage(listener)
        return () => subscription.dispose()
      },
    })
    this.bridge.handle('source.outline', async (raw) => {
      const request = outlineReadRequestSchema.parse(raw)
      const state = input.getState()
      if (state.stateId !== request.stateId)
        throw new BridgeError('revision_conflict', 'The outline request is stale.')
      const result = await readCurrentOutline({
        state,
        source: input.source,
        sessionId: input.original.sidecarSessionId,
        sheetId: request.sheetId,
      })
      if (input.getState().stateId !== state.stateId)
        throw new BridgeError(
          'revision_conflict',
          'The document changed while reading outline metadata.',
        )
      return outlineReadResultSchema.parse(result)
    })
    this.bridge.handle('source.goalSeek', (raw) => this.goalSeek.handle(raw))
    this.bridge.handle('view.ready', () => {
      this.notifyReady()
      return { descriptor: input.getDescriptor(), state: input.getState() }
    })
    this.bridge.handle('view.resync', () => ({
      descriptor: input.getDescriptor(),
      state: input.getState(),
    }))
    this.bridge.handle('source.media', (request) =>
      readCurrentMedia({
        request,
        getState: input.getState,
        source: input.source,
        sourceSessionId: input.original.sidecarSessionId,
      }),
    )
    this.bridge.handle('source.audit', (request) =>
      auditWorkbook({
        request,
        getState: input.getState,
        source: input.source,
        sourceSessionId: input.original.sidecarSessionId,
      }),
    )
    this.bridge.handle('source.find', (raw) => this.search.handle(raw))
    this.bridge.handle('source.read', async (raw) => {
      const request = sourceReadSchema.parse(raw),
        state = input.getState()
      if (request.stateId !== state.stateId)
        throw new BridgeError('revision_conflict', 'The requested state is no longer current.')
      return readCurrentRange({
        state,
        sessionId: input.original.sidecarSessionId,
        source: input.source,
        logicalSheetId: request.logicalSheetId,
        range: request.range,
        includeMerges: true,
      })
    })
    this.bridge.handle('view.intent', async (raw) => {
      const intent = transactionInputSchema.parse(raw)
      if (intent.sessionId !== input.sessionId || intent.origin !== 'manual')
        throw new BridgeError('session_mismatch', 'Unexpected manual transaction source.')
      return input.submit(intent)
    })
    // Host-bounded stages (CSV import): one manual commit above the ordinary batch ceiling.
    this.bridge.handle('view.stagedIntent', async (raw) => {
      const request = z
        .strictObject({
          intent: transactionInputSchema.omit({ operations: true }),
          stages: z.unknown(),
        })
        .parse(raw)
      if (request.intent.sessionId !== input.sessionId || request.intent.origin !== 'manual')
        throw new BridgeError('session_mismatch', 'Unexpected manual transaction source.')
      if (!input.submitStaged)
        throw new BridgeError('unsupported_operation', 'Staged edits are unavailable in this host.')
      return input.submitStaged(request.intent, request.stages)
    })
    this.bridge.handle('history.request', async (raw) => {
      const request = z.strictObject({ direction: z.enum(['undo', 'redo']) }).parse(raw)
      if (!input.history)
        throw new BridgeError(
          'unsupported_operation',
          'Native history is unavailable in this host.',
        )
      await input.history(request.direction)
      return { completed: true }
    })
  }
  async prepare(
    state: Readonly<WorkbookState>,
    intent: TransactionInput,
    signal?: AbortSignal,
  ): Promise<PreparedChangeset> {
    const material = await prepareCandidateMaterial({
      state,
      operations: intent.operations,
      source: this.input.source,
      sourceSessionId: this.input.original.sidecarSessionId,
      signal,
    })
    // Inline picture roots can approach the transfer limit when repeated as
    // next/forward/inverse. Execute their shared interpreter authoritatively
    // here; only the accepted root is projected to the view.
    // Sheet passwords use Excel's 100,000-round hash: run it once, on the host.
    if (
      intent.operations.some((op) => op.op === 'add_picture' || op.op === 'protect_sheet') ||
      state.sheets.some((sheet) =>
        (sheet.features?.visuals as { kind?: string }[] | undefined)?.some(
          (v) => v.kind === 'image',
        ),
      )
    )
      return computeCandidateState(state, intent, material, { signal })
    const result = await this.bridge.request(
      'candidate.prepare',
      JSON.parse(JSON.stringify({ state, intent, material })),
    )
    if (signal?.aborted)
      throw new BridgeError('operation_cancelled', 'The candidate was cancelled before commit.')
    const parsed = candidateResultSchema.parse(result)
    if (parsed.beforeStateId !== state.stateId)
      throw new BridgeError('revision_conflict', 'The candidate has an obsolete baseline.')
    const verified = await computeCandidateState(state, intent, material, { signal })
    if (JSON.stringify(parsed) !== JSON.stringify(candidateResultSchema.parse(verified)))
      throw new BridgeError(
        'candidate_mismatch',
        'The view candidate differs from host execution of the authorized intent.',
      )
    return parsed
  }
  dispose() {
    this.goalSeek.dispose()
    this.search.dispose()
    this.bridge.dispose()
  }
}
