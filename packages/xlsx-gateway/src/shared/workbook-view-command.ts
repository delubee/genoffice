import { z } from 'zod'

/** Host-owned presentation commands; no URI, operations or arbitrary command ID. */
export const workbookViewCommandSchema = z.strictObject({
  command: z.enum(['find', 'replace', 'goto']),
})
export type WorkbookViewCommand = z.infer<typeof workbookViewCommandSchema>['command']
