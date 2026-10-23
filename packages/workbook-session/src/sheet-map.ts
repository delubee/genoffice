import { WorkbookError, type WorkbookState } from './domain'

export function logicalSheet(state: Readonly<WorkbookState>, logicalId: string) {
  const sheet = state.sheets.find((candidate) => candidate.id === logicalId)
  if (!sheet)
    throw new WorkbookError('sheet_not_found', 'The logical worksheet no longer exists.', {
      logicalSheetId: logicalId,
    })
  return sheet
}

export function originalSheet(state: Readonly<WorkbookState>, originalId: string) {
  const sheet = state.sheets.find((candidate) => candidate.originalSheetId === originalId)
  if (!sheet)
    throw new WorkbookError('sheet_not_found', 'The original worksheet is not in this state.')
  return sheet
}
