import { z } from 'zod'

/** Host-only: hidden definition bodies must never enter public workbook metadata
 * or renderer/AI bridge responses. Scope IDs refer to immutable source sheets.
 */
export const calculationContextSchema = z
  .object({
    sessionId: z.string().uuid(),
    names: z.array(
      z
        .object({
          name: z.string().min(1),
          formula: z.string(),
          hidden: z.boolean(),
          sheetId: z.string().min(1).optional(),
        })
        .strict(),
    ),
    calculationProperties: z.record(z.string(), z.string()),
    externalRelationshipIds: z.array(z.string().min(1)),
  })
  .strict()
export type CalculationContext = z.infer<typeof calculationContextSchema>
