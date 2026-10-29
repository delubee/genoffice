import { z } from 'zod'
import { workbookStyleEditSchema } from './edit-schemas'
const scalar = z.union([z.string(), z.number().finite(), z.boolean(), z.null()])
const calculation = z.enum(['ready', 'pending', 'error'])

export const cellRecordSchema = z
  .strictObject({
    kind: z.enum(['blank', 'string', 'number', 'boolean', 'date', 'error']),
    rawValue: scalar,
    formula: z.string().startsWith('=').optional(),
    calculationState: calculation.optional(),
    cachedValue: scalar.optional(),
    numberFormat: z.string().optional(),
    dateSystem: z.enum(['1900', '1904']).optional(),
    style: workbookStyleEditSchema.optional(),
    /** Replace the original cell format with xf 0 before applying this style. */
    styleReset: z.boolean().optional(),
  })
  .superRefine((cell, ctx) => {
    const valid =
      cell.kind === 'blank'
        ? cell.rawValue === null
        : cell.kind === 'number' || cell.kind === 'date'
          ? typeof cell.rawValue === 'number'
          : cell.kind === 'boolean'
            ? typeof cell.rawValue === 'boolean'
            : typeof cell.rawValue === 'string'
    if (!valid) ctx.addIssue({ code: 'custom', message: 'Cell kind does not match its raw value' })
    if (cell.kind === 'date' && !cell.dateSystem)
      ctx.addIssue({ code: 'custom', message: 'Dates require an explicit date system' })
  })
