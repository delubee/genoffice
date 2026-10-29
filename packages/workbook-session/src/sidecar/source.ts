import { calculationContextSchema } from '../calculation/context'
import { parseRange } from '@genoffice/xlsx-gateway/domain/cell-address'
import {
  workbookFileSchema,
  workbookMediaResultSchema,
  workbookFindCellsRequestSchema,
  workbookFindCellsResultSchema,
  type WorkbookFindCellsRequest,
  workbookRangeRequestSchema,
  workbookRangeResultSchema,
  workbookMergeRangesResultSchema,
  workbookRowOutlineResultSchema,
} from '@genoffice/xlsx-gateway/shared/workbook-protocol'
import type { WorkbookSource } from '../ports'
import { WorkbookError } from '../domain'
import { XlsxSidecarClient } from './client'

/** Validates every sidecar response at the shared protocol boundary. */
export class SidecarWorkbookSource implements WorkbookSource {
  private features = new Map<string, Record<string, unknown>>()
  private metadata = new Map<string, ReturnType<typeof workbookFileSchema.parse>>()
  private mergeFeatures = new Map<
    string,
    ReturnType<typeof workbookMergeRangesResultSchema.parse>
  >()
  constructor(readonly client: XlsxSidecarClient) {}
  async openSnapshot(path: string) {
    const metadata = workbookFileSchema
      .omit({ sha256: true, readOnly: true })
      .parse(await this.client.open(path))
    const file = workbookFileSchema.parse({ ...metadata, sha256: '0'.repeat(64), readOnly: false })
    this.metadata.set(file.sessionId, file)
    return { sessionId: file.sessionId, metadata: file }
  }
  async readMedia(sessionId: string, visualId: string) {
    if (!this.metadata.get(sessionId)?.visuals.some((visual) => visual.id === visualId))
      throw new WorkbookError('session_mismatch', 'This visual is not in the source session.')
    return workbookMediaResultSchema.parse(await this.client.readMedia({ sessionId, visualId }))
  }
  async readCalculationContext(sessionId: string) {
    const metadata = this.metadata.get(sessionId)
    if (!metadata)
      throw new WorkbookError('session_mismatch', 'The source session has been closed.')
    const context = calculationContextSchema.parse(
      await this.client.readCalculationContext(sessionId),
    )
    if (
      context.sessionId !== sessionId ||
      context.names.some(
        (name) =>
          name.sheetId !== undefined && !metadata.sheets.some((sheet) => sheet.id === name.sheetId),
      )
    )
      throw new WorkbookError(
        'session_mismatch',
        'Calculation resources belong to a different source session.',
      )
    return context
  }
  async findBaseCells(raw: WorkbookFindCellsRequest) {
    const request = workbookFindCellsRequestSchema.parse(raw)
    if (
      !request.sheetId ||
      !this.metadata.get(request.sessionId)?.sheets.some((sheet) => sheet.id === request.sheetId)
    )
      throw new WorkbookError('session_mismatch', 'This worksheet is not in the source session.')
    const result = workbookFindCellsResultSchema.parse(await this.client.findCells(request))
    if (
      result.matches.some((match) => match.sheetId !== request.sheetId) ||
      (result.nextCursor && result.nextCursor.sheetId !== request.sheetId)
    )
      throw new WorkbookError('session_mismatch', 'Search returned a different source worksheet.')
    return result
  }
  async readRowOutline(sessionId: string, sheetId: string) {
    if (!this.metadata.get(sessionId)?.sheets.some((sheet) => sheet.id === sheetId))
      throw new WorkbookError('session_mismatch', 'This worksheet is not in the source session.')
    return workbookRowOutlineResultSchema.parse(
      await this.client.readRowOutline({ sessionId, sheetId }),
    )
  }
  async readBaseRange(sessionId: string, sheetId: string, range: string) {
    const parsed = parseRange(range)
    const request = workbookRangeRequestSchema.parse({ sessionId, sheetId, range: parsed })
    if (!this.metadata.get(sessionId)?.sheets.some((sheet) => sheet.id === sheetId))
      throw new WorkbookError('session_mismatch', 'This worksheet is not in the source session.')
    const result = workbookRangeResultSchema.parse(await this.client.readRange(request))
    if (
      !result.indexingComplete &&
      (result.indexedThroughRow === null || result.indexedThroughRow < parsed.endRow)
    )
      throw new WorkbookError(
        'index_pending',
        'The source has not indexed the complete requested range yet.',
        {},
        true,
      )
    this.captureFeatures(sessionId, sheetId, result)
    return result
  }
  async readFeatures(sessionId: string, sheetId: string) {
    const file = this.metadata.get(sessionId),
      sheet = file?.sheets.find((item) => item.id === sheetId)
    if (!file || !sheet)
      throw new WorkbookError('session_mismatch', 'The source session has been closed.')
    const key = `${sessionId}:${sheetId}`
    if (!this.features.get(key)?.indexingComplete)
      this.captureFeatures(
        sessionId,
        sheetId,
        workbookRangeResultSchema.parse(
          await this.client.readRange({
            sessionId,
            sheetId,
            range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
          }),
        ),
      )
    if (!this.mergeFeatures.get(key)?.indexingComplete)
      this.mergeFeatures.set(
        key,
        workbookMergeRangesResultSchema.parse(
          await this.client.readMergeRanges({ sessionId, sheetId }),
        ),
      )
    const mergeFeatures = this.mergeFeatures.get(key)!
    return {
      ...this.features.get(key),
      ...mergeFeatures,
      indexingComplete: Boolean(
        this.features.get(key)?.indexingComplete && mergeFeatures.indexingComplete,
      ),
      sheet,
      styles: file.styles,
      nativeVisuals: file.visuals.filter((visual) => visual.sheetId === sheetId),
    }
  }
  private captureFeatures(
    sessionId: string,
    sheetId: string,
    result: ReturnType<typeof workbookRangeResultSchema.parse>,
  ) {
    const key = `${sessionId}:${sheetId}`,
      previous = this.features.get(key) ?? {},
      { cells: _cells, rows: _rows, ...features } = result
    const next: Record<string, unknown> = { ...previous, ...features }
    for (const name of ['merges', 'hyperlinks', 'conditionalRules', 'protectedRanges'])
      if (Array.isArray(features[name as keyof typeof features]))
        next[name] = [
          ...new Map(
            [
              ...((previous[name] as unknown[]) ?? []),
              ...(features[name as keyof typeof features] as unknown[]),
            ].map((value) => [JSON.stringify(value), value]),
          ).values(),
        ]
    this.features.set(key, next)
  }
  archiveManifest(path: string) {
    return this.client.archiveManifest(path)
  }
  async close(sessionId: string) {
    try {
      await this.client.close(sessionId)
    } finally {
      this.metadata.delete(sessionId)
      for (const key of this.features.keys())
        if (key.startsWith(`${sessionId}:`)) this.features.delete(key)
      for (const key of this.mergeFeatures.keys())
        if (key.startsWith(`${sessionId}:`)) this.mergeFeatures.delete(key)
    }
  }
}
