import { randomUUID } from 'node:crypto'
import { z } from 'zod'
import { workbookStateSchema } from './schemas'
import type { WorkbookState } from './domain'

const metadataSchema = z
  .object({
    date1904: z.boolean().optional(),
    readOnly: z.boolean().optional(),
    sheets: z
      .array(
        z
          .object({
            id: z.string(),
            name: z.string(),
            rowCount: z.number().int().nonnegative(),
            columnCount: z.number().int().nonnegative(),
            protected: z.boolean().optional(),
          })
          .passthrough(),
      )
      .min(1),
  })
  .passthrough()

export function immutable<T>(value: T): T {
  if (value && typeof value === 'object' && !Object.isFrozen(value)) {
    Object.freeze(value)
    for (const child of Object.values(value)) immutable(child)
  }
  return value
}

export function initialState(rawMetadata: unknown): WorkbookState {
  const metadata = metadataSchema.parse(rawMetadata)
  const { sheets: _sheets, ...workbookMetadata } = metadata
  const state = workbookStateSchema.parse({
    schemaVersion: 1,
    stateId: randomUUID(),
    dateSystem: metadata.date1904 ? '1904' : '1900',
    structuralLog: [],
    features: { originalMetadata: JSON.parse(JSON.stringify(workbookMetadata)) },
    sheets: metadata.sheets.map((sheet) => ({
      id: randomUUID(),
      originalSheetId: sheet.id,
      name: sheet.name,
      rows: sheet.rowCount,
      columns: sheet.columnCount,
      cells: {},
      readOnly: metadata.readOnly ?? false,
      protected: sheet.protected ?? false,
      features: { originalMetadata: JSON.parse(JSON.stringify(sheet)) },
    })),
  })
  for (const sheet of state.sheets)
    state.sheetMap[sheet.id] = { originalSheetId: sheet.originalSheetId!, originalName: sheet.name }
  return immutable(state)
}
