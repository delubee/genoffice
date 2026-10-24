export { JsonBridge, BridgeError } from '@genoffice/xlsx-gateway/shared/session-bridge'
export type { BridgeMethod, MessageTransport } from '@genoffice/xlsx-gateway/shared/session-bridge'

import { z } from 'zod'
import { JsonBridge } from '@genoffice/xlsx-gateway/shared/session-bridge'
import {
  workbookStateSchema,
  transactionInputSchema,
} from '@genoffice/xlsx-gateway/shared/session-state'
import { IsolatedCandidateInterpreter } from './candidate'

const candidateRequestSchema = z.strictObject({
  state: workbookStateSchema,
  intent: transactionInputSchema,
  material: z.strictObject({
    state: workbookStateSchema,
    completeSheetIds: z.array(z.string()),
    referenceRanges: z.record(z.string(), z.array(z.string())).optional(),
    writableSheetIds: z.array(z.string()),
    allocation: z.strictObject({
      stateId: z.string().min(1),
      sheetIds: z.record(z.string(), z.string().min(1)),
    }),
  }),
})

/** Installs only data handlers; code, paths and process execution are absent. */
export function attachCandidateHandler(bridge: JsonBridge) {
  const active = new Set<IsolatedCandidateInterpreter>()
  const remove = bridge.handle('candidate.prepare', async (raw) => {
    const request = candidateRequestSchema.parse(raw)
    if (request.intent.sessionId !== bridge.sessionId) throw new Error('session_mismatch')
    const interpreter = new IsolatedCandidateInterpreter({
      async hydrate() {
        return request.material
      },
    })
    active.add(interpreter)
    try {
      return await interpreter.prepare(request.state, request.intent)
    } finally {
      interpreter.dispose()
      active.delete(interpreter)
    }
  })
  return () => {
    remove()
    for (const interpreter of active) interpreter.dispose()
    active.clear()
  }
}
