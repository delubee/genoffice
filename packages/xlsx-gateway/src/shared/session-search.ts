import { z } from 'zod'

export const sessionFindQuerySchema = z.strictObject({
  stateId: z.string().min(1),
  query: z.string().trim().min(1).max(10_000),
  sheetIds: z.array(z.string().min(1)).min(1).max(500).optional(),
  matchCase: z.boolean().default(false),
  matchEntireCell: z.boolean().default(false),
  lookIn: z.enum(['values', 'formulas']).default('values'),
  direction: z.enum(['row', 'column']).default('row'),
})
export const sessionSearchRequestSchema = z.discriminatedUnion('action', [
  z.strictObject({
    action: z.literal('replace'),
    stateId: z.string().min(1),
    searchId: z.string().uuid(),
    mode: z.enum(['one', 'all']),
    index: z.number().int().nonnegative().optional(),
    replacement: z.string().max(32_767),
  }),
  sessionFindQuerySchema.extend({ action: z.literal('start'), queryId: z.string().uuid() }),
  z.strictObject({
    action: z.literal('cancel'),
    stateId: z.string().min(1),
    queryId: z.string().uuid(),
  }),
  z.strictObject({
    action: z.literal('page'),
    stateId: z.string().min(1),
    searchId: z.string().uuid(),
    offset: z.number().int().nonnegative(),
  }),
  z.strictObject({
    action: z.literal('release'),
    stateId: z.string().min(1),
    searchId: z.string().uuid(),
  }),
])
export type SessionFindQuery = z.infer<typeof sessionFindQuerySchema>
export type SessionSearchRequest = z.infer<typeof sessionSearchRequestSchema>

export const sessionSearchPageSchema = z.strictObject({
  stateId: z.string().min(1),
  searchId: z.string().uuid(),
  offset: z.number().int().nonnegative(),
  total: z.number().int().min(0).max(100_000),
  valuesPending: z.boolean(),
  nextOffset: z.number().int().nonnegative().optional(),
  hits: z
    .array(
      z.strictObject({
        sheetId: z.string().min(1),
        address: z.string().regex(/^[A-Z]{1,3}[1-9][0-9]{0,6}$/),
        text: z.string().max(512),
        truncated: z.boolean(),
        isFormula: z.boolean(),
        replaceable: z.boolean(),
      }),
    )
    .max(200),
})
export type SessionSearchPage = z.infer<typeof sessionSearchPageSchema>
