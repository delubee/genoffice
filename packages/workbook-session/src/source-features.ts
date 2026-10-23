import { z } from 'zod'
import { WorkbookError, assertNotCancelled } from './domain'
import type { WorkbookSource } from './ports'

const featuresSchema = z
  .object({
    indexingComplete: z.boolean(),
    sheetProtection: z
      .object({ protected: z.boolean().optional() })
      .passthrough()
      .nullable()
      .optional(),
  })
  .passthrough()

export async function indexedFeatures(
  source: WorkbookSource,
  sessionId: string,
  sheetId: string,
  signal?: AbortSignal,
) {
  const deadline = Date.now() + 10_000
  for (;;) {
    assertNotCancelled(signal)
    const features = featuresSchema.parse(await source.readFeatures(sessionId, sheetId))
    if (features.indexingComplete) return features
    if (Date.now() >= deadline)
      throw new WorkbookError(
        'index_pending',
        'Wait for worksheet indexing; merge and protection metadata is not complete yet.',
        {},
        true,
      )
    await new Promise((resolve) => setTimeout(resolve, 25))
  }
}
