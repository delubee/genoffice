import { z } from 'zod'
const address = z.string().regex(/^[A-Z]{1,3}[1-9][0-9]{0,6}$/)
export const workbookAuditRequestSchema = z.strictObject({
  stateId: z.string().min(1),
  kind: z.enum(['dependents', 'errors']),
  sheetId: z.string().min(1),
  address,
})
export const workbookAuditResultSchema = z.strictObject({
  stateId: z.string().min(1),
  valuesPending: z.boolean(),
  hits: z
    .array(
      z.strictObject({
        sheetId: z.string().min(1),
        address,
        value: z.string().max(512).optional(),
      }),
    )
    .max(100_000),
})
export type WorkbookAuditRequest = z.infer<typeof workbookAuditRequestSchema>
export type WorkbookAuditResult = z.infer<typeof workbookAuditResultSchema>
export const FORMULA_ERROR_VALUE =
  /^#(?:REF!|DIV\/0!|VALUE!|NAME\?|N\/A|NUM!|NULL!|SPILL!|CALC!|GETTING_DATA)$/i
