import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import { stateChangeset } from '@genoffice/xlsx-gateway/shared/session-changeset'
import {
  workbookOperationSchema,
  type WorkbookOperation,
} from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { z } from 'zod'
import {
  WorkbookError,
  assertNotCancelled,
  type TransactionInput,
  type WorkbookState,
} from './domain'
import { prepareCandidateMaterial } from './candidate-material'
import type { WorkbookSource } from './ports'

/** Upper bound on cells a host-built staged transaction may write (desktop CSV/merge scale). */
export const STAGED_WRITE_CELLS = 50_000

/**
 * Host-built stages: each one is an ordinary bounded transaction, so the
 * per-batch write ceiling and structural/content separation still hold.
 * Never accepted from a model-facing tool.
 */
export const stagedOperationsSchema = z
  .array(z.array(workbookOperationSchema).min(1).max(2000))
  .min(1)
  .max(200)

/** Stages may target a sheet an earlier stage adds, before its ID exists. */
export const NEW_SHEET_PREFIX = 'new-sheet:'

function resolveNewSheets(
  operations: readonly WorkbookOperation[],
  before: WorkbookState,
  current: WorkbookState,
): WorkbookOperation[] {
  return operations.map((op) => {
    if (!('sheetId' in op) || !op.sheetId.startsWith(NEW_SHEET_PREFIX)) return op
    const name = op.sheetId.slice(NEW_SHEET_PREFIX.length)
    const added = current.sheets.find(
      (sheet) => sheet.name === name && !before.sheets.some((item) => item.id === sheet.id),
    )
    if (!added)
      throw new WorkbookError('target_invalid', `No earlier stage added the sheet "${name}".`)
    return { ...op, sheetId: added.id }
  })
}

function stagedWriteCells(stages: readonly (readonly WorkbookOperation[])[]) {
  let cells = 0
  for (const stage of stages)
    for (const op of stage)
      if (op.op === 'paste_cells') cells += op.cells.length * op.cells[0]!.length
      else if (op.op === 'set_range') cells += op.values.length * op.values[0]!.length
      else cells += 1
  return cells
}

/** Applies each stage to a private root; only the combined root enters host commit. */
export async function prepareStagedCandidate(input: {
  state: WorkbookState
  intent: TransactionInput
  stages: readonly (readonly WorkbookOperation[])[]
  source: WorkbookSource
  sourceSessionId: string
  signal?: AbortSignal
}) {
  const stages = stagedOperationsSchema.parse(input.stages)
  if (stagedWriteCells(stages) > STAGED_WRITE_CELLS)
    throw new WorkbookError(
      'expanded_write_limit',
      `A staged transaction may write at most ${STAGED_WRITE_CELLS} cells.`,
    )
  let current = input.state,
    pending = false
  const affected: string[] = []
  for (const stage of stages) {
    assertNotCancelled(input.signal)
    const operations = resolveNewSheets(stage, input.state, current)
    const material = await prepareCandidateMaterial({
      state: current,
      operations,
      source: input.source,
      sourceSessionId: input.sourceSessionId,
      signal: input.signal,
    })
    const changes = await computeCandidateState(
      current,
      { ...input.intent, operations },
      material,
      { signal: input.signal },
    )
    if (changes.calculationState === 'pending') pending = true
    affected.push(...changes.affectedTargets)
    current = changes.nextState
  }
  assertNotCancelled(input.signal)
  const changes = stateChangeset(input.state, current, [...new Set(affected)])
  return pending ? { ...changes, calculationState: 'pending' as const } : changes
}
