import { Worker } from 'node:worker_threads'
import { existsSync } from 'node:fs'
import type { CalculationSourceData } from './source'
import { calculateSourceWorkbook } from './evaluate-source'

type Calculated = Awaited<ReturnType<typeof calculateSourceWorkbook>>

let workerPath: string | undefined
let shared: { worker: Worker } | undefined
let queue: Promise<unknown> = Promise.resolve()

/** Where the bundled worker lives; unset (or missing) means calculations run in-process. */
export function setCalculationWorkerPath(path: string | undefined) {
  workerPath = path
}

function start(path: string) {
  const worker = new Worker(path)
  // A crashed or terminated worker is replaced on the next calculation.
  worker.once('exit', () => {
    if (shared?.worker === worker) shared = undefined
  })
  worker.unref()
  return worker
}

async function run(
  path: string,
  source: CalculationSourceData,
  options: { signal?: AbortSignal; sheetId?: string; maxResultCells?: number },
): Promise<Calculated> {
  const worker = (shared ??= { worker: start(path) }).worker
  const { signal, ...cloneable } = options
  return new Promise<Calculated>((resolve, reject) => {
    const id = Math.random()
    const finish = () => {
      worker.off('message', onMessage)
      worker.off('error', onError)
      worker.off('exit', onExit)
      signal?.removeEventListener('abort', onAbort)
    }
    const onMessage = (message: { id: number; result?: Calculated; error?: string }) => {
      if (message.id !== id) return
      finish()
      if (message.error !== undefined) reject(new Error(message.error))
      else resolve(message.result!)
    }
    const onError = (error: Error) => {
      finish()
      reject(error)
    }
    const onExit = () => {
      finish()
      reject(new Error('The calculation worker stopped.'))
    }
    // An abort cannot interrupt the engine, so the worker itself is stopped.
    const onAbort = () => {
      finish()
      if (shared?.worker === worker) shared = undefined
      void worker.terminate()
      reject(signal?.reason instanceof Error ? signal.reason : new Error('Calculation cancelled.'))
    }
    if (signal?.aborted) return onAbort()
    worker.on('message', onMessage)
    worker.on('error', onError)
    worker.on('exit', onExit)
    signal?.addEventListener('abort', onAbort, { once: true })
    worker.postMessage({ id, source, options: cloneable })
  })
}

/**
 * Evaluate a loaded root off the calling thread when a worker is available. Calculations are
 * queued one at a time, as the engine keeps module-level state.
 */
export function calculateOffThread(
  source: CalculationSourceData,
  options: { signal?: AbortSignal; sheetId?: string; maxResultCells?: number } = {},
): Promise<Calculated> {
  if (!workerPath || !existsSync(workerPath)) return calculateSourceWorkbook(source, options)
  const path = workerPath
  const work = queue.catch(() => undefined).then(() => run(path, source, options))
  queue = work
  return work
}
