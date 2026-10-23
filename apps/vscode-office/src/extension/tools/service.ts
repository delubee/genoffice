import { z } from 'zod'
import { parseRange, rangeCellCount } from '@genoffice/xlsx-gateway/domain/cell-address'
import {
  LIMITS,
  WorkbookError,
  assertNotCancelled,
  type DocumentDescriptor,
  type WorkbookState,
  type TransactionInput,
} from '@genoffice/workbook-session/domain'
import { aggregateCurrentRange } from '@genoffice/workbook-session/aggregate'
import { assertFreshRevision } from '@genoffice/workbook-session/freshness'
import type { WorkbookBindings } from '@genoffice/workbook-session/bindings'
import type { readCurrentRange } from '@genoffice/workbook-session/read'
import { toolSchemas, type ToolKind } from './schemas'
import { SESSION_OPERATION_NAMES } from '@genoffice/xlsx-gateway/shared/session-operations'
import type { WorkbookExchangeTools } from './exchange-tools'

export interface ToolDocument {
  descriptor(): DocumentDescriptor
  state(): WorkbookState
  projectionError?(): string | null
  flush(signal?: AbortSignal): Promise<void>
  snapshot(): {
    descriptor: DocumentDescriptor
    state: WorkbookState
    read(sheetId: string, range: string, signal?: AbortSignal): ReturnType<typeof readCurrentRange>
    release(): Promise<void>
  }
  submit(intent: TransactionInput, signal?: AbortSignal): Promise<unknown>
  operationStatus(transactionId: string): Promise<{ status: string }>
}
export interface ToolHost {
  exchange?: Pick<
    WorkbookExchangeTools,
    'export' | 'prepareImport' | 'applyImport' | 'planStatus' | 'describePlan' | 'describeExport'
  >
  bindings: WorkbookBindings
  documents(): ToolDocument[]
  operationStatus(sessionId: string, transactionId: string): Promise<unknown>
}

export class WorkbookTools {
  constructor(private readonly host: ToolHost) {}
  private parse(kind: string, raw: unknown) {
    if (['export', 'prepareImport', 'applyImport'].includes(kind) && !this.host.exchange)
      throw new WorkbookError(
        'unsupported_operation',
        'File exchange tools are not available in this host.',
      )
    if (Buffer.byteLength(JSON.stringify(raw) ?? 'null') > LIMITS.messageBytes)
      throw new WorkbookError('input_limit', 'Tool input exceeds 1 MiB; shrink this request.')
    if (!Object.hasOwn(toolSchemas, kind))
      throw new WorkbookError('unsupported_operation', 'This workbook tool is not available.')
    return toolSchemas[kind as ToolKind].parse(raw)
  }
  private target(bindingId: string) {
    const binding = this.host.bindings.get(bindingId)
    const document = this.host
      .documents()
      .find(
        (item) =>
          item.descriptor().sessionId === binding.sessionId &&
          item.descriptor().documentUri === binding.documentUri,
      )
    if (!document || binding.status === 'invalidated')
      throw new WorkbookError(
        'session_closed',
        'The bound workbook session is closed; bind an open document again.',
      )
    return { binding, document }
  }
  prepare(kind: string, raw: unknown) {
    const input = this.parse(kind, raw)
    const target =
      'bindingId' in input && input.bindingId
        ? this.target(input.bindingId).binding.documentUri
        : 'documentUri' in input
          ? input.documentUri
          : undefined
    const effect =
      'operations' in input
        ? input.operations
            .slice(0, 12)
            .map((operation) => {
              const document =
                'bindingId' in input ? this.target(input.bindingId!).document : undefined
              const sheet =
                'sheetId' in operation
                  ? document?.state().sheets.find((item) => item.id === operation.sheetId)
                  : undefined
              const area =
                'address' in operation
                  ? operation.address
                  : 'range' in operation
                    ? operation.range
                    : 'start' in operation
                      ? `${operation.start}; ${'values' in operation ? operation.values.reduce((count, row) => count + row.length, 0) : 0} cells`
                      : 'source' in operation && 'target' in operation
                        ? `${operation.source} → ${operation.target}`
                        : 'dataRange' in operation
                          ? operation.dataRange
                          : 'row' in operation
                            ? `row ${operation.row}; count ${'count' in operation ? operation.count : 1}`
                            : 'column' in operation
                              ? `column ${operation.column}; count ${'count' in operation ? operation.count : 1}`
                              : 'name' in operation
                                ? operation.name
                                : 'workbook object'
              return `${operation.op}: ${sheet?.name ?? ('sheetId' in operation ? operation.sheetId : 'workbook')}; ${area}`
            })
            .join('\n')
        : ''
    let exchangeEffect = ''
    if (kind === 'export') {
      const options = toolSchemas.export.parse(raw)
      exchangeEffect = `\nWrites a complete typed snapshot in ${this.host.exchange!.describeExport(options.bindingId, options.workspaceUri)} under .genoffice/work/<snapshotId>/. Does not modify the workbook or execute Python.`
    } else if (kind === 'prepareImport') {
      const options = toolSchemas.prepareImport.parse(raw)
      exchangeEffect = `\nReads ${options.resultManifestUri} and creates a validated private immutable copy. No workbook cells are modified. Source declarations do not prove business correctness.`
    } else if (kind === 'applyImport') {
      const options = toolSchemas.applyImport.parse(raw),
        plan = this.host.exchange!.describePlan(options.bindingId, options.planId)
      exchangeEffect = `\nApplies batch cursor ${options.cursor} of ${plan.batchCount}, at most 1000 cells, to ${JSON.stringify(plan.target)}; shape ${plan.shape.rows}×${plan.shape.columns}; ${plan.overwriteCellCount} existing cells may be overwritten across this plan. Mode: ${plan.mode}. One undo step for this batch; no automatic save.`
    }
    return {
      confirmationRequired: ['apply', 'export', 'prepareImport', 'applyImport'].includes(kind),
      title:
        kind === 'apply' || kind === 'applyImport'
          ? 'Edit workbook'
          : kind === 'export'
            ? 'Export workbook snapshot'
            : kind === 'prepareImport'
              ? 'Validate workbook result'
              : 'Read workbook',
      message: `${kind}: ${target ?? ('sessionId' in input ? input.sessionId : 'explicit target required')}${'range' in input && input.range ? `; range ${input.range}` : ''}${'operations' in input ? `; ${input.operations.length} operation(s); revision ${input.expectedRevision}; unsaved, undoable workbook changes\n${effect}${input.operations.length > 12 ? '\nAdditional operations are included in the same batch.' : ''}` : exchangeEffect ? exchangeEffect : '; no AI workbook content changes'}`,
    }
  }
  private ok(document: ToolDocument | undefined, data: unknown, summary: string) {
    const descriptor = document?.descriptor()
    return {
      schemaVersion: 1,
      status: 'ok' as const,
      sessionId: descriptor?.sessionId ?? null,
      revision: descriptor?.revision ?? null,
      summary,
      data,
    }
  }
  async invoke(kind: string, raw: unknown, signal?: AbortSignal) {
    try {
      this.parse(kind, raw)
      if (kind === 'applyImport') {
        const result = await this.host.exchange!.applyImport(
          toolSchemas.applyImport.parse(raw),
          signal,
        )
        return {
          ...this.ok(
            undefined,
            result.data,
            'One import batch result; report its completed and remaining counts.',
          ),
          sessionId: result.sessionId,
          revision: result.revision,
        }
      }
      if (kind === 'apply') return await this.apply(toolSchemas.apply.parse(raw), signal)
      assertNotCancelled(signal)
      if (kind === 'export' || kind === 'prepareImport') {
        const result =
          kind === 'export'
            ? await this.host.exchange!.export(toolSchemas.export.parse(raw), signal)
            : await this.host.exchange!.prepareImport(toolSchemas.prepareImport.parse(raw), signal)
        return {
          ...this.ok(
            undefined,
            result.data,
            kind === 'export'
              ? 'Complete current-state snapshot exported to the workspace host.'
              : 'Result validated and staged; workbook content is unchanged.',
          ),
          sessionId: result.sessionId,
          revision: result.revision,
        }
      }
      if (kind === 'context') return await this.context(toolSchemas.context.parse(raw), signal)
      if (kind === 'operationStatus') {
        const input = toolSchemas.operationStatus.parse(raw)
        return {
          ...this.ok(
            undefined,
            'planId' in input
              ? (this.host.exchange?.planStatus(input.sessionId, input.planId) ??
                  (() => {
                    throw new WorkbookError(
                      'unsupported_operation',
                      'Import plans are not supported by this host.',
                    )
                  })())
              : await this.host.operationStatus(input.sessionId, input.transactionId),
            'Transaction status; no workbook changes.',
          ),
          sessionId: input.sessionId,
        }
      }
      const input = (kind === 'read' ? toolSchemas.read : toolSchemas.aggregate).parse(raw)
      const { document } = this.target(input.bindingId)
      assertFreshRevision(
        document.descriptor(),
        document.descriptor().sessionId,
        input.expectedRevision,
      )
      await document.flush(signal)
      assertNotCancelled(signal)
      assertFreshRevision(
        document.descriptor(),
        document.descriptor().sessionId,
        input.expectedRevision,
      )
      const snapshot = document.snapshot()
      try {
        const count = rangeCellCount(parseRange(input.range))
        const includeFormats = 'includeFormats' in input && input.includeFormats === true
        if (kind === 'read' && count > (includeFormats ? LIMITS.formatCells : LIMITS.readCells))
          throw new WorkbookError(
            'range_limit',
            includeFormats
              ? 'Read at most 200 cells with formats.'
              : 'Read at most 2,000 cells per request.',
          )
        const data =
          kind === 'read'
            ? await snapshot.read(input.sheetId, input.range, signal)
            : await aggregateCurrentRange({
                range: input.range,
                signal,
                read: async (range) => (await snapshot.read(input.sheetId, range, signal)).cells,
              })
        if (kind === 'read' && !includeFormats && 'cells' in data)
          data.cells = data.cells.map(
            ({ style: _style, numberFormat: _numberFormat, ...cell }) => cell,
          )
        if (Buffer.byteLength(JSON.stringify(data)) > LIMITS.messageBytes)
          throw new WorkbookError(
            'response_limit',
            'The typed result exceeds 1 MiB; read a smaller range. No records were truncated.',
          )
        if (
          document.descriptor().sessionId === snapshot.descriptor.sessionId &&
          document.descriptor().revision === snapshot.descriptor.revision
        )
          this.host.bindings.acknowledgeRead(input.bindingId, snapshot.descriptor.revision)
        return {
          ...this.ok(
            document,
            { ...data, projectionError: document.projectionError?.() ?? null },
            'Read the current immutable workbook state; no workbook changes.',
          ),
          revision: snapshot.descriptor.revision,
          sessionId: snapshot.descriptor.sessionId,
        }
      } finally {
        await snapshot.release()
      }
    } catch (error) {
      const failure =
        error instanceof z.ZodError
          ? new WorkbookError('invalid_input', 'Tool input does not match schemaVersion 1.', {
              issues: error.issues.map((issue) => ({
                path: issue.path.join('.'),
                message: issue.message,
              })),
            })
          : error instanceof WorkbookError
            ? error
            : new WorkbookError(
                'tool_failed',
                error instanceof Error ? error.message : String(error),
              )
      return {
        schemaVersion: 1,
        status: 'error' as const,
        code: failure.code,
        message: failure.message,
        details: failure.details,
        retryable: failure.retryable,
      }
    }
  }
  private async context(input: z.infer<typeof toolSchemas.context>, signal?: AbortSignal) {
    if (!input.bindingId && !input.documentUri)
      return this.ok(
        undefined,
        {
          requires_target: true,
          documents: this.host.documents().map((item) => item.descriptor()),
        },
        'Choose an explicit open document URI or binding; focus is not a target.',
      )
    let document: ToolDocument, binding
    if (input.bindingId) ({ document, binding } = this.target(input.bindingId))
    else {
      const found = this.host
        .documents()
        .find((item) => item.descriptor().documentUri === input.documentUri)
      if (!found)
        throw new WorkbookError(
          'session_closed',
          'The explicit workbook URI is not open in this window.',
        )
      document = found
    }
    await document.flush(signal)
    assertNotCancelled(signal)
    const descriptor = document.descriptor(),
      state = document.state()
    if (!binding)
      binding = this.host.bindings.bind(
        descriptor.sessionId,
        input.sheetId ?? state.sheets[0]!.id,
        input.range ?? null,
      )
    return this.ok(
      document,
      {
        binding,
        descriptor,
        projectionError: document.projectionError?.() ?? null,
        dateSystem: state.dateSystem,
        sheets: state.sheets.map((sheet) => ({
          logicalSheetId: sheet.id,
          name: sheet.name,
          rowCount: sheet.rows,
          columnCount: sheet.columns,
          readOnly: sheet.readOnly,
          protection: {
            protected:
              !sheet.originalSheetId || Object.hasOwn(sheet.features ?? {}, 'sheetProtection')
                ? sheet.protected
                : null,
            metadataVerified:
              !sheet.originalSheetId || Object.hasOwn(sheet.features ?? {}, 'sheetProtection'),
            verifiedBeforeEveryWrite: true,
          },
        })),
        capabilities: {
          operations: SESSION_OPERATION_NAMES,
          charts: ['column', 'bar', 'line', 'pie'],
          pythonExchange: !!this.host.exchange,
          automaticSave: false,
        },
        limits: LIMITS,
      },
      'Current workbook context; binding does not depend on editor focus.',
    )
  }
  private async apply(input: z.infer<typeof toolSchemas.apply>, signal?: AbortSignal) {
    const { document, binding } = this.target(input.bindingId)
    const intent: TransactionInput = {
      schemaVersion: 1,
      sessionId: binding.sessionId,
      transactionId: input.transactionId,
      expectedRevision: input.expectedRevision,
      origin: 'agent',
      summary: input.summary,
      operations: input.operations,
    }
    if ((await document.operationStatus(input.transactionId)).status === 'committed')
      return this.ok(
        document,
        await document.submit(intent, signal),
        'Previously committed transaction; not applied again.',
      )
    assertNotCancelled(signal)
    assertFreshRevision(document.descriptor(), binding.sessionId, input.expectedRevision)
    await document.flush(signal)
    assertNotCancelled(signal)
    this.host.bindings.authorizeWrite(
      input.bindingId,
      input.expectedRevision,
      input.usesFrozenSelection ?? true,
    )
    return this.ok(
      document,
      await document.submit(intent, signal),
      'Committed one undoable batch; the XLSX has not been saved.',
    )
  }
}
