import { sessionDefinedNames } from '@genoffice/xlsx-gateway/domain/session-defined-names'
import { shiftedSessionFormula } from '@genoffice/xlsx-gateway/domain/session-formula'
import {
  rewriteNameReferences,
  type NameReferenceChange,
} from '@genoffice/xlsx-gateway/domain/defined-name-references'
import { WorkbookError, type WorkbookState } from '../domain'
import type { CalculationContext } from './context'
import type { CalculationNameDefinition } from './names'

/** Compose source-only definitions and current action-time journal definitions.
 * Replaying the full history over a newly authored reference would rebind it.
 * This data stays private; hidden definition bodies never enter canonical state.
 */
export function composeCalculationDefinitions(
  state: Readonly<WorkbookState>,
  context: CalculationContext,
): CalculationNameDefinition[] {
  const journal = state.features?.definedNamesEdits ? sessionDefinedNames(state) : undefined
  const preserve = new Set(journal?.preserveNames ?? [])
  const plan = {
    changes: (state.features?.definedNameReferences ?? []) as unknown as NameReferenceChange[],
    sheets: state.sheets.map(({ id, name }) => ({ id, name })),
  }
  const definitions: CalculationNameDefinition[] = []
  for (const entry of context.names) {
    if (
      journal &&
      !entry.hidden &&
      entry.formula.trim() &&
      !entry.name.toLowerCase().startsWith('_xlnm') &&
      !preserve.has(entry.name)
    )
      continue
    const owner =
      entry.sheetId === undefined
        ? undefined
        : state.sheets.find((sheet) => sheet.originalSheetId === entry.sheetId)
    if (entry.sheetId !== undefined && !owner) continue
    const formula = rewriteNameReferences(
      shiftedSessionFormula(state, owner?.id ?? '', entry.formula),
      owner?.id ?? '',
      plan,
    )
    definitions.push({
      name: entry.name,
      formula,
      hidden: entry.hidden,
      ...(owner ? { sheetId: owner.id } : {}),
    })
  }
  for (const entry of journal?.names ?? []) {
    if (
      entry.scopeSheetId !== undefined &&
      !state.sheets.some((sheet) => sheet.id === entry.scopeSheetId)
    )
      throw new WorkbookError(
        'invalid_document',
        'A current name refers to a missing logical worksheet.',
      )
    definitions.push({
      name: entry.name,
      formula: entry.formula,
      ...(entry.scopeSheetId === undefined ? {} : { sheetId: entry.scopeSheetId }),
    })
  }
  return definitions
}
