import { z } from 'zod'
import { workbookStateSchema } from './session-state'

export type CandidateState = z.infer<typeof workbookStateSchema>

import { assertSerializable } from './json-data'
export { assertSerializable } from './json-data'

export function stateChangeset(
  before: CandidateState,
  after: CandidateState,
  affectedTargets: readonly string[],
) {
  assertSerializable(before)
  assertSerializable(after)
  const previous = workbookStateSchema.parse(before),
    next = workbookStateSchema.parse(after)
  return {
    beforeStateId: previous.stateId,
    nextState: next,
    forward: { schemaVersion: 1, fromStateId: previous.stateId, state: next },
    inverse: { schemaVersion: 1, fromStateId: next.stateId, state: previous },
    affectedTargets: [...new Set(affectedTargets)],
    calculationState: next.sheets.some((sheet) =>
      Object.values(sheet.cells).some((cell) => cell.formula),
    )
      ? ('pending' as const)
      : ('ready' as const),
  }
}
