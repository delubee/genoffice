import { z } from 'zod'
import { addressSchema } from './session-state'
export const goalSeekRequestSchema = z.strictObject({
  stateId: z.string().min(1),
  revision: z.number().int().nonnegative(),
  sheetId: z.string().min(1),
  setCell: addressSchema,
  byCell: addressSchema,
  toValue: z.number().finite(),
})
export const goalSeekResultSchema = z.strictObject({
  stateId: z.string().min(1),
  revision: z.number().int().nonnegative(),
  found: z.boolean(),
  solution: z.number().finite(),
  reached: z.number().finite().nullable(),
  iterations: z.number().int().min(1).max(100),
})
export type GoalSeekRequest = z.infer<typeof goalSeekRequestSchema>
