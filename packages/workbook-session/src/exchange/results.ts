import { WorkbookError, type WorkbookState } from '../domain'
import { parseAddress } from '@genoffice/xlsx-gateway/domain/cell-address'
import { resultManifestSchema, type ResultManifest, type TypedCell } from './schemas'
import type { SnapshotAuthority } from './snapshots'

export function validateResultManifest(
  raw: unknown,
  source: SnapshotAuthority,
  state: WorkbookState,
  sessionId: string,
  revision: number,
): ResultManifest {
  if ((raw as { schemaVersion?: unknown })?.schemaVersion !== 1)
    throw new WorkbookError('schema_unsupported', 'Result schemaVersion must be 1.')
  const parsed = resultManifestSchema.safeParse(raw)
  if (!parsed.success)
    throw new WorkbookError('type_mismatch', 'Invalid, incomplete or empty result manifest.', {
      issues: parsed.error.issues.map((issue) => issue.message),
    })
  const result = parsed.data,
    manifest = source.manifest,
    target = result.target
  if (state.sheets.some((sheet) => sheet.readOnly))
    throw new WorkbookError(
      'target_invalid',
      'The bound workbook is read-only; Save As and reopen a writable copy.',
    )
  if (
    manifest.sessionId !== sessionId ||
    result.source.snapshotId !== manifest.snapshotId ||
    result.source.sourceManifestSha256 !== source.manifestSha256 ||
    result.source.sourceDataSha256 !== manifest.data.sha256 ||
    result.source.sourceRevision !== manifest.sourceRevision
  )
    throw new WorkbookError(
      'source_mismatch',
      'The result does not match the privately registered source snapshot.',
    )
  if (revision !== manifest.sourceRevision || state.stateId !== manifest.sourceStateId)
    throw new WorkbookError(
      'snapshot_stale',
      'The source workbook changed; export and calculate a new snapshot.',
      {},
      true,
    )
  const start = parseAddress(result.target.startCell)
  if (start.row + result.shape.rows > 1048576 || start.column + result.shape.columns > 16384)
    throw new WorkbookError(
      'target_invalid',
      'The complete result does not fit inside the Excel grid.',
    )
  if ('sheetId' in target) {
    const sheet = state.sheets.find((item) => item.id === target.sheetId)
    if (!sheet || sheet.readOnly || sheet.protected)
      throw new WorkbookError(
        'target_invalid',
        'The result target must be an editable logical sheet in the bound workbook.',
      )
  } else if (
    state.sheets.some((sheet) => sheet.name.toLowerCase() === target.newSheetName.toLowerCase())
  )
    throw new WorkbookError(
      'target_invalid',
      'The new worksheet name is already used; choose an explicit existing sheet or a new name.',
    )
  return result
}

export function validateResultCell(
  cell: TypedCell,
  result: ResultManifest,
  dateSystem: WorkbookState['dateSystem'],
) {
  if (cell.kind === 'formula' && result.mode !== 'formulas')
    throw new WorkbookError(
      'type_mismatch',
      'Formula cells require explicit formulas mode; string prefixes never imply formulas.',
    )
  const date = cell.kind === 'formula' ? cell.cached : cell
  if (date?.kind === 'date' && date.dateSystem !== dateSystem)
    throw new WorkbookError(
      'type_mismatch',
      'Result dateSystem must match the target workbook; no implicit date conversion is performed.',
    )
}
