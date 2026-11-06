import { z } from 'zod'
import { MAX_GRID_COLUMNS, MAX_GRID_ROWS } from '../shared/grid-bounds'
import type { StructuralOp } from '../gateway/xlsx-structure'

export const setOutlineSchema = z
  .object({
    op: z.literal('set_outline'),
    sheetId: z.string().min(1),
    axis: z.enum(['rows', 'cols']),
    start: z.number().int().min(1).max(MAX_GRID_ROWS),
    count: z.number().int().min(1).max(10_000),
    level: z.number().int().min(0).max(7),
    collapsed: z.boolean().optional(),
  })
  .refine(
    (value) =>
      value.start + value.count - 1 <= (value.axis === 'rows' ? MAX_GRID_ROWS : MAX_GRID_COLUMNS),
    'Outline span exceeds the worksheet grid.',
  )
export const setOutlineSettingsSchema = z.object({
  op: z.literal('set_outline_settings'),
  sheetId: z.string().min(1),
  summaryBelow: z.boolean(),
  summaryRight: z.boolean(),
})
export const outlineEditSchema = z.discriminatedUnion('op', [
  setOutlineSchema,
  setOutlineSettingsSchema,
])
export function outlineStructuralOp(edit: z.infer<typeof outlineEditSchema>): StructuralOp {
  if (edit.op === 'set_outline_settings')
    return {
      kind: 'set-outline-pr',
      summaryBelow: edit.summaryBelow,
      summaryRight: edit.summaryRight,
    }
  return {
    kind: edit.axis === 'rows' ? 'set-rows-outline' : 'set-cols-outline',
    start: edit.start - 1,
    end: edit.start + edit.count - 2,
    level: edit.level,
    ...(edit.collapsed === undefined ? {} : { collapsed: edit.collapsed }),
  }
}

export const outlineReadRequestSchema = z.strictObject({
  stateId: z.string().min(1),
  sheetId: z.string().min(1),
})
const entrySchema = z.strictObject({
  index: z.number().int().nonnegative(),
  level: z.number().int().min(0).max(7),
  collapsed: z.boolean(),
  hidden: z.boolean(),
})
export const outlineReadResultSchema = z.strictObject({
  stateId: z.string().min(1),
  sheetId: z.string().min(1),
  summaryBelow: z.boolean(),
  summaryRight: z.boolean(),
  rows: z
    .array(
      entrySchema.extend({
        index: z
          .number()
          .int()
          .min(0)
          .max(MAX_GRID_ROWS - 1),
      }),
    )
    .max(MAX_GRID_ROWS),
  cols: z
    .array(
      entrySchema.extend({
        index: z
          .number()
          .int()
          .min(0)
          .max(MAX_GRID_COLUMNS - 1),
      }),
    )
    .max(MAX_GRID_COLUMNS),
})
