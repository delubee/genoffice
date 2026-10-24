import { randomUUID, createHash } from 'node:crypto'
import { z } from 'zod'
import { assertSerializable } from '@genoffice/xlsx-gateway/shared/json-data'
import {
  LIMITS,
  WorkbookError,
  type HistoryEntry,
  type PreparedChangeset,
  type TransactionInput,
  type WorkbookState,
} from './domain'
import { workbookStateSchema } from './schemas'
import type { ResourceStore } from './ports'
import { resourceManifestSchema } from './store'
import { assertHistoryBudget } from './history-budget'

const changeSchema = z.strictObject({
  schemaVersion: z.literal(1),
  fromStateId: z.string().min(1),
  state: workbookStateSchema,
})
const manifestPayloadSchema = z.strictObject({
  kind: z.literal('history-state'),
  fromStateId: z.string(),
  stateId: z.string(),
  serializedBytes: z.number().int().nonnegative(),
  chunks: z.array(z.string()),
})

export async function prepareHistoryRecord(input: {
  store: ResourceStore
  before: WorkbookState
  prepared: PreparedChangeset
  origin: TransactionInput['origin']
  label: string
  retainedBytes?: number
  retainedEntries?: readonly HistoryEntry[]
  transactionByteLimit?: number
  historyByteLimit?: number
}) {
  const { prepared, before, store } = input
  assertSerializable(prepared.forward)
  assertSerializable(prepared.inverse)
  assertSerializable(prepared.nextState)
  const forward = changeSchema.parse(prepared.forward),
    inverse = changeSchema.parse(prepared.inverse)
  if (
    prepared.beforeStateId !== before.stateId ||
    forward.fromStateId !== before.stateId ||
    inverse.fromStateId !== prepared.nextState.stateId ||
    JSON.stringify(inverse.state) !== JSON.stringify(before) ||
    JSON.stringify(forward.state) !== JSON.stringify(prepared.nextState)
  )
    throw new WorkbookError(
      'invalid_changeset',
      'The bidirectional changeset does not match the canonical roots.',
    )
  const forwardBytes = Buffer.from(JSON.stringify(forward)),
    inverseBytes = Buffer.from(JSON.stringify(inverse))
  const entryId = randomUUID(),
    forwardRef = `history-${entryId}-forward`,
    inverseRef = `history-${entryId}-inverse`
  function manifestFor(bytes: Buffer, fromStateId: string, stateId: string) {
    const references: string[] = []
    for (let offset = 0; offset < bytes.length; offset += LIMITS.chunkBytes)
      references.push(
        createHash('sha256')
          .update(bytes.subarray(offset, offset + LIMITS.chunkBytes))
          .digest('hex'),
      )
    return {
      schemaVersion: 1,
      references,
      payload: {
        kind: 'history-state' as const,
        fromStateId,
        stateId,
        serializedBytes: bytes.length,
        chunks: references,
      },
    }
  }
  const forwardManifest = manifestFor(forwardBytes, before.stateId, prepared.nextState.stateId)
  const inverseManifest = manifestFor(inverseBytes, prepared.nextState.stateId, before.stateId)
  const retained = input.retainedEntries
    ? await retainedHistoryStorage(store, input.retainedEntries)
    : { bytes: input.retainedBytes ?? 0, chunks: new Map<string, number>() }
  let addition = 0
  for (const manifest of [forwardManifest, inverseManifest]) {
    addition += Buffer.byteLength(JSON.stringify(manifest))
    for (const [index, reference] of manifest.references.entries()) {
      const length = Math.min(
        LIMITS.chunkBytes,
        manifest.payload.serializedBytes - index * LIMITS.chunkBytes,
      )
      if (!retained.chunks.has(reference)) {
        addition += length
        retained.chunks.set(reference, length)
      }
    }
  }
  const serializedBytes = assertHistoryBudget({
    forwardBytes: forwardBytes.length,
    inverseBytes: inverseBytes.length,
    retainedBytes: retained.bytes,
    retainedAdditionBytes: input.retainedEntries ? addition : undefined,
    transactionByteLimit: input.transactionByteLimit,
    historyByteLimit: input.historyByteLimit,
  })
  for (const [name, bytes, manifest] of [
    [forwardRef, forwardBytes, forwardManifest],
    [inverseRef, inverseBytes, inverseManifest],
  ] as const) {
    for (let offset = 0; offset < bytes.length; offset += LIMITS.chunkBytes)
      await store.writeChunk(bytes.subarray(offset, offset + LIMITS.chunkBytes))
    await store.publishManifest(name, manifest)
  }
  return {
    entryId,
    beforeStateId: before.stateId,
    afterStateId: prepared.nextState.stateId,
    origin: input.origin,
    label: input.label,
    forwardRef,
    inverseRef,
    serializedBytes,
  } satisfies HistoryEntry
}

export async function loadHistoryState(
  store: ResourceStore,
  entry: HistoryEntry,
  direction: 'forward' | 'inverse',
) {
  const manifest = resourceManifestSchema.parse(
    await store.readManifest(direction === 'forward' ? entry.forwardRef : entry.inverseRef),
  )
  const payload = manifestPayloadSchema.parse(manifest.payload)
  const expectedStateId = direction === 'forward' ? entry.afterStateId : entry.beforeStateId
  const expectedFromId = direction === 'forward' ? entry.beforeStateId : entry.afterStateId
  if (
    payload.stateId !== expectedStateId ||
    payload.fromStateId !== expectedFromId ||
    payload.serializedBytes > LIMITS.transactionBytes ||
    JSON.stringify(manifest.references) !== JSON.stringify(payload.chunks)
  )
    throw new WorkbookError('history_corrupt', 'The history root has invalid references.')
  const pieces = []
  let bytes = 0
  for (const reference of payload.chunks) {
    const chunk = await store.readChunk(reference)
    bytes += chunk.length
    if (bytes > payload.serializedBytes)
      throw new WorkbookError('history_corrupt', 'The history data exceeds its declared size.')
    pieces.push(chunk)
  }
  if (bytes !== payload.serializedBytes)
    throw new WorkbookError('history_corrupt', 'The history state is incomplete.')
  const change = changeSchema.parse(JSON.parse(Buffer.concat(pieces).toString('utf8')))
  if (change.state.stateId !== expectedStateId || change.fromStateId !== payload.fromStateId)
    throw new WorkbookError('history_corrupt', 'The history state does not match the entry.')
  return change.state
}

export async function historyReferences(store: ResourceStore, entries: readonly HistoryEntry[]) {
  const references = new Set<string>()
  for (const entry of entries)
    for (const name of [entry.forwardRef, entry.inverseRef])
      for (const reference of resourceManifestSchema.parse(
        await (store.readManifestMetadata?.(name) ?? store.readManifest(name)),
      ).references)
        references.add(reference)
  return [...references]
}

/** Counts actual serialized files, charging shared content-addressed chunks once. */
export async function retainedHistoryStorage(
  store: ResourceStore,
  entries: readonly HistoryEntry[],
) {
  const chunks = new Map<string, number>()
  let bytes = 0
  for (const entry of entries)
    for (const name of [entry.forwardRef, entry.inverseRef]) {
      const manifest = resourceManifestSchema.parse(
        await (store.readManifestMetadata?.(name) ?? store.readManifest(name)),
      )
      const payload = manifestPayloadSchema.parse(manifest.payload)
      if (
        JSON.stringify(manifest.references) !== JSON.stringify(payload.chunks) ||
        payload.serializedBytes > LIMITS.transactionBytes
      )
        throw new WorkbookError('history_corrupt', 'Retained history has inconsistent references.')
      bytes += Buffer.byteLength(JSON.stringify(manifest))
      for (const [index, reference] of payload.chunks.entries()) {
        const length = Math.min(
          LIMITS.chunkBytes,
          payload.serializedBytes - index * LIMITS.chunkBytes,
        )
        if (length <= 0 || (chunks.has(reference) && chunks.get(reference) !== length))
          throw new WorkbookError('history_corrupt', 'Retained chunk lengths are inconsistent.')
        if (!chunks.has(reference)) {
          chunks.set(reference, length)
          bytes += length
        }
      }
    }
  return { bytes, chunks }
}
