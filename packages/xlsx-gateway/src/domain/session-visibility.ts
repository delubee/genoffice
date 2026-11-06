import { z } from 'zod'
import type { StructuralOp } from '../gateway/xlsx-structure'

/** Current-coordinate axis visibility. Kept separately from cell content and sizes. */
export const visibilityEditSchema = z
  .object({
    axis: z.enum(['rows', 'cols']),
    start: z.number().int().nonnegative(),
    count: z.number().int().positive().max(10_000),
    hidden: z.boolean(),
  })
  .refine((edit) => edit.start + edit.count <= (edit.axis === 'rows' ? 1048576 : 16384))
export const visibilityEditsSchema = z.array(visibilityEditSchema)
export type VisibilityEdit = z.infer<typeof visibilityEditSchema>

/** Undefined means preserve the immutable source/filter visibility. */
export function axisHiddenOverride(
  edits: readonly VisibilityEdit[],
  axis: 'rows' | 'cols',
  index: number,
): boolean | undefined {
  for (let i = edits.length - 1; i >= 0; i--) {
    const edit = edits[i]!
    if (edit.axis === axis && index >= edit.start && index < edit.start + edit.count)
      return edit.hidden
  }
}

export function visibilityStructuralOp(edit: VisibilityEdit): StructuralOp {
  return {
    kind: edit.axis === 'rows' ? 'set-rows-hidden' : 'set-cols-hidden',
    start: edit.start,
    end: edit.start + edit.count - 1,
    hidden: edit.hidden,
  }
}
