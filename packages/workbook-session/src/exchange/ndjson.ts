import { createHash, randomUUID } from 'node:crypto'
import { createReadStream } from 'node:fs'
import { open, rename, rm } from 'node:fs/promises'
import { basename } from 'node:path'
import { LIMITS, WorkbookError, assertNotCancelled } from '../domain'
import {
  exchangeRecordSchema,
  shapeSchema,
  type ExchangeRecord,
  type ExchangeShape,
  type DataFile,
} from './schemas'

function coverage(shape: ExchangeShape) {
  let row = 0,
    column = 0
  return {
    accept(record: ExchangeRecord) {
      if (
        record.rowOffset !== row ||
        record.columnOffset !== column ||
        row >= shape.rows ||
        column + record.cells.length > shape.columns
      )
        throw new WorkbookError(
          'artifact_incomplete',
          'Records must cover the rectangle exactly, in order, without overlaps or gaps.',
        )
      column += record.cells.length
      if (column === shape.columns) {
        row++
        column = 0
      }
    },
    finish() {
      if (row !== shape.rows || column !== 0)
        throw new WorkbookError(
          'artifact_incomplete',
          'The data does not cover the complete declared rectangle.',
        )
    },
  }
}

export async function writeTypedNdjson(input: {
  path: string
  shape: ExchangeShape
  records: AsyncIterable<ExchangeRecord> | Iterable<ExchangeRecord>
  signal?: AbortSignal
  byteLimit?: number
}): Promise<DataFile> {
  const shape = shapeSchema.parse(input.shape),
    check = coverage(shape),
    temporary = `${input.path}.${randomUUID()}.tmp`,
    hash = createHash('sha256')
  const limit = Math.min(input.byteLimit ?? LIMITS.artifactBytes, LIMITS.artifactBytes)
  let bytes = 0,
    recordCount = 0,
    handle
  try {
    handle = await open(temporary, 'wx', 0o600)
    for await (const raw of input.records) {
      assertNotCancelled(input.signal)
      const record = exchangeRecordSchema.parse(raw)
      check.accept(record)
      let column = record.columnOffset,
        parts: string[] = [],
        length = 0
      const flush = async () => {
        if (!parts.length) return
        const encoded = Buffer.from(
          `{"rowOffset":${record.rowOffset},"columnOffset":${column},"cells":[${parts.join(',')}]}\n`,
        )
        if (encoded.length - 1 > LIMITS.recordBytes || bytes + encoded.length > limit)
          throw new WorkbookError(
            'artifact_limit_exceeded',
            'The typed artifact exceeds its record or 64 MiB data budget; split the range.',
          )
        await handle!.writeFile(encoded)
        hash.update(encoded)
        bytes += encoded.length
        recordCount++
        column += parts.length
        parts = []
        length = 0
      }
      for (const cell of record.cells) {
        const encoded = JSON.stringify(cell),
          size = Buffer.byteLength(encoded)
        // Account for the complete bounded record envelope before adding a cell.
        const envelope = Buffer.byteLength(
          `{"rowOffset":${record.rowOffset},"columnOffset":${column},"cells":[]}`,
        )
        if (parts.length && envelope + length + parts.length + size > LIMITS.recordBytes)
          await flush()
        parts.push(encoded)
        length += size
      }
      await flush()
    }
    check.finish()
    assertNotCancelled(input.signal)
    await handle.sync()
    await handle.close()
    handle = undefined
    await rename(temporary, input.path)
    return {
      fileName: basename(input.path),
      sha256: hash.digest('hex'),
      bytes,
      recordCount,
      encoding: 'utf-8',
      format: 'typed-ndjson-v1',
    }
  } finally {
    await handle?.close()
    await rm(temporary, { force: true }).catch(() => {})
  }
}

export async function* readTypedNdjson(input: {
  path: string
  shape: ExchangeShape
  expected: Pick<DataFile, 'sha256' | 'bytes' | 'recordCount'>
  signal?: AbortSignal
  onRecordPosition?(record: ExchangeRecord, byteOffset: number): void
}): AsyncGenerator<ExchangeRecord> {
  const shape = shapeSchema.parse(input.shape),
    check = coverage(shape),
    hash = createHash('sha256'),
    decoder = new TextDecoder('utf-8', { fatal: true })
  let pending = Buffer.alloc(0),
    bytes = 0,
    records = 0,
    position = 0
  const stream = createReadStream(input.path, { highWaterMark: 64 * 1024 })
  try {
    for await (const chunk of stream) {
      assertNotCancelled(input.signal)
      const buffer = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk)
      bytes += buffer.length
      if (bytes > LIMITS.artifactBytes || bytes > input.expected.bytes)
        throw new WorkbookError(
          'artifact_limit_exceeded',
          'Artifact exceeds its declared byte budget.',
        )
      hash.update(buffer)
      pending = Buffer.concat([pending, buffer])
      for (;;) {
        const end = pending.indexOf(10)
        if (end < 0) break
        if (!end || end > LIMITS.recordBytes)
          throw new WorkbookError('artifact_limit_exceeded', 'Invalid or oversized NDJSON record.')
        const line = pending.subarray(0, end)
        pending = pending.subarray(end + 1)
        let record: ExchangeRecord
        try {
          record = exchangeRecordSchema.parse(JSON.parse(decoder.decode(line)))
        } catch (error) {
          throw new WorkbookError(
            'type_mismatch',
            error instanceof Error ? error.message : String(error),
          )
        }
        check.accept(record)
        records++
        input.onRecordPosition?.(record, position)
        position += end + 1
        yield record
      }
      if (pending.length > LIMITS.recordBytes)
        throw new WorkbookError('artifact_limit_exceeded', 'NDJSON record exceeds 256 KiB.')
    }
    assertNotCancelled(input.signal)
    if (pending.length)
      throw new WorkbookError(
        'artifact_incomplete',
        'NDJSON must end with a complete newline-terminated record.',
      )
    check.finish()
    if (
      bytes !== input.expected.bytes ||
      records !== input.expected.recordCount ||
      hash.digest('hex') !== input.expected.sha256
    )
      throw new WorkbookError(
        'artifact_hash_mismatch',
        'Artifact hash, bytes or record count does not match its manifest.',
      )
  } finally {
    stream.destroy()
  }
}

/** Reads only an indexed batch from a previously verified private immutable copy. */
export async function readPrivateBatch(input: {
  path: string
  byteOffset: number
  cellOffset: number
  count: number
  signal?: AbortSignal
}) {
  const stream = createReadStream(input.path, { start: input.byteOffset, highWaterMark: 64 * 1024 })
  const decoder = new TextDecoder('utf-8', { fatal: true })
  let pending = Buffer.alloc(0),
    skip = input.cellOffset
  const cells: ExchangeRecord['cells'] = []
  try {
    for await (const chunk of stream) {
      assertNotCancelled(input.signal)
      pending = Buffer.concat([pending, Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk)])
      for (;;) {
        const end = pending.indexOf(10)
        if (end < 0) break
        if (!end || end > LIMITS.recordBytes)
          throw new WorkbookError('artifact_incomplete', 'Private result record is corrupt.')
        const record = exchangeRecordSchema.parse(
          JSON.parse(decoder.decode(pending.subarray(0, end))),
        )
        pending = pending.subarray(end + 1)
        for (const cell of record.cells.slice(skip)) {
          cells.push(cell)
          if (cells.length === input.count) return cells
        }
        skip = 0
      }
      if (pending.length > LIMITS.recordBytes)
        throw new WorkbookError('artifact_limit_exceeded', 'Private record exceeds 256 KiB.')
    }
    throw new WorkbookError('artifact_incomplete', 'Private result batch is incomplete.')
  } finally {
    stream.destroy()
  }
}
