import { randomUUID } from 'node:crypto'
import { join } from 'node:path'
import { formatAddress, parseRange } from '@genoffice/xlsx-gateway/domain/cell-address'
import {
  WorkbookError,
  assertNotCancelled,
  type DocumentDescriptor,
  type WorkbookState,
} from '../domain'
import { immutable } from '../state'
import { logicalSheet } from '../sheet-map'
import { writeTypedNdjson } from './ndjson'
import { toTypedCell } from './cells'
import { workspaceTaskDirectory, publishJson } from './paths'
import {
  shapeSchema,
  snapshotManifestSchema,
  type SnapshotManifest,
  type ExchangeRecord,
} from './schemas'
import type { readCurrentRange } from '../read'
import { captureReadView } from '../freshness'

export interface ExchangeReadView {
  descriptor: DocumentDescriptor
  state: WorkbookState
  read(sheetId: string, range: string, signal?: AbortSignal): ReturnType<typeof readCurrentRange>
}
export interface SnapshotAuthority {
  manifest: SnapshotManifest
  manifestSha256: string
  workspaceRoot: string
  directory: string
  manifestPath: string
  dataPath: string
}

export class WorkbookSnapshots {
  private records = new Map<string, SnapshotAuthority>()
  get(snapshotId: string): SnapshotAuthority {
    const record = this.records.get(snapshotId)
    if (!record)
      throw new WorkbookError(
        'source_mismatch',
        'This snapshot is unknown in the current host session.',
      )
    return structuredClone(record)
  }
  invalidateSession(sessionId: string) {
    for (const [id, record] of this.records)
      if (record.manifest.sessionId === sessionId) this.records.delete(id)
  }
  async export(input: {
    view: ExchangeReadView
    sheetId: string
    range: string
    workspaceRoot: string
    signal?: AbortSignal
  }) {
    assertNotCancelled(input.signal)
    captureReadView(input.view.descriptor, input.view.state)
    const { view } = input,
      sheet = logicalSheet(view.state, input.sheetId),
      bounds = parseRange(input.range)
    const parsedShape = shapeSchema.safeParse({
      rows: bounds.endRow - bounds.startRow + 1,
      columns: bounds.endColumn - bounds.startColumn + 1,
      cellCount:
        (bounds.endRow - bounds.startRow + 1) * (bounds.endColumn - bounds.startColumn + 1),
    })
    if (!parsedShape.success)
      throw new WorkbookError(
        'artifact_limit_exceeded',
        'Export a nonempty rectangle of at most 1,000,000 cells; split larger ranges.',
      )
    const shape = parsedShape.data
    if (bounds.endRow >= 1048576 || bounds.endColumn >= 16384)
      throw new WorkbookError('target_invalid', 'Snapshot range exceeds the Excel grid.')
    const snapshotId = randomUUID(),
      { root, directory } = await workspaceTaskDirectory(input.workspaceRoot, snapshotId),
      dataPath = join(directory, 'input.ndjson'),
      manifestPath = join(directory, 'input.manifest.json')
    let calculationState: SnapshotManifest['calculationState'] = 'ready'
    async function* rows(): AsyncGenerator<ExchangeRecord> {
      for (let row = bounds.startRow; row <= bounds.endRow;) {
        const rowEnd = Math.min(
          bounds.endRow,
          row + Math.max(1, Math.floor(2000 / shape.columns)) - 1,
        )
        for (let column = bounds.startColumn; column <= bounds.endColumn; column += 2000) {
          const columnEnd = Math.min(bounds.endColumn, column + 1999)
          const page = await view.read(
            input.sheetId,
            `${formatAddress(row, column)}:${formatAddress(rowEnd, columnEnd)}`,
            input.signal,
          )
          const width = columnEnd - column + 1
          if (
            page.stateId !== view.state.stateId ||
            page.cells.length !== (rowEnd - row + 1) * width
          )
            throw new WorkbookError(
              'source_mismatch',
              'Snapshot pages do not belong to one complete immutable state.',
            )
          for (let offset = 0; offset <= rowEnd - row; offset++) {
            const cells = page.cells.slice(offset * width, (offset + 1) * width).map(toTypedCell)
            for (const cell of cells)
              if (cell.kind === 'formula') {
                if (cell.calculationState === 'error') calculationState = 'error'
                else if (cell.calculationState === 'pending' && calculationState !== 'error')
                  calculationState = 'pending'
              }
            yield {
              rowOffset: row - bounds.startRow + offset,
              columnOffset: column - bounds.startColumn,
              cells,
            }
          }
        }
        row = rowEnd + 1
      }
    }
    const data = await writeTypedNdjson({
      path: dataPath,
      shape,
      records: rows(),
      signal: input.signal,
    })
    const manifest = snapshotManifestSchema.parse({
      schemaVersion: 1,
      kind: 'workbook-snapshot',
      snapshotId,
      sessionId: view.descriptor.sessionId,
      documentUri: view.descriptor.documentUri,
      sourceRevision: view.descriptor.revision,
      sourceStateId: view.state.stateId,
      sheet: { logicalSheetId: sheet.id, name: sheet.name },
      range: input.range,
      shape,
      dateSystem: view.state.dateSystem,
      calculationState,
      data,
      complete: true,
    })
    const manifestSha256 = await publishJson(manifestPath, manifest, input.signal)
    const record = immutable({
      manifest,
      manifestSha256,
      workspaceRoot: root,
      directory,
      manifestPath,
      dataPath,
    })
    this.records.set(snapshotId, record)
    return structuredClone(record)
  }
}
