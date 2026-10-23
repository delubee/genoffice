import { parseRange } from '@genoffice/xlsx-gateway/domain/cell-address'
import {
  workbookFileSchema,
  workbookRangeRequestSchema,
  workbookRangeResultSchema,
} from '@genoffice/xlsx-gateway/shared/workbook-protocol'
import type { WorkbookSource } from '../ports'
import { WorkbookError } from '../domain'
import { XlsxSidecarClient } from './client'

/** Validates every sidecar response at the shared protocol boundary. */
export class SidecarWorkbookSource implements WorkbookSource {
  private features = new Map<string, Record<string, unknown>>()
  private metadata = new Map<string, ReturnType<typeof workbookFileSchema.parse>>()
  constructor(readonly client: XlsxSidecarClient) {}
  async openSnapshot(path: string) {
    const metadata = workbookFileSchema
      .omit({ sha256: true, readOnly: true })
      .parse(await this.client.open(path))
    const file = workbookFileSchema.parse({ ...metadata, sha256: '0'.repeat(64), readOnly: false })
    this.metadata.set(file.sessionId, file)
    return { sessionId: file.sessionId, metadata: file }
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
    return {
      ...this.features.get(key),
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
    }
  }
}
