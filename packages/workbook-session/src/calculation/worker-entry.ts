import { parentPort } from 'node:worker_threads'
import { calculateSourceWorkbook } from './evaluate-source'

/** Worker thread entry: one calculation per message, so the extension host stays responsive. */
parentPort!.on('message', async (message: { id: number; source: never; options: never }) => {
  try {
    const result = await calculateSourceWorkbook(message.source, message.options)
    parentPort!.postMessage({ id: message.id, result })
  } catch (error) {
    parentPort!.postMessage({
      id: message.id,
      error: error instanceof Error ? error.message : String(error),
    })
  }
})
