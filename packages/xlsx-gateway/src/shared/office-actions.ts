import { z } from 'zod'

/** Explicit menu intents; the host supplies the target resource and capabilities. */
export const officeActionSchema = z.discriminatedUnion('action', [
  z.object({ action: z.literal('open') }).strict(),
  z.object({ action: z.literal('save') }).strict(),
  z.object({ action: z.literal('saveAs') }).strict(),
  z
    .object({
      action: z.literal('chat'),
      instruction: z.string().trim().max(4000).default(''),
      scope: z.enum(['document', 'selection']).default('document'),
    })
    .strict(),
])
export type OfficeAction = z.infer<typeof officeActionSchema>
export interface OfficeActionHost {
  action(request: OfficeAction): Promise<void>
}
