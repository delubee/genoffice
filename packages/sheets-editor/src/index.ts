/** Version of the browser-to-host workbook message protocol. */
export const EDITOR_PROTOCOL_VERSION = 1 as const
export { IsolatedCandidateInterpreter } from './candidate'
export type { CandidateContext, CandidateMaterial } from './candidate'
export { stateChangeset } from './changeset'
export type { CandidateState } from './changeset'
export type * from './ports'
