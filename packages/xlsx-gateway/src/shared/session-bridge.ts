import { z } from 'zod'
import { protocolMessageSchema } from './session-state'
import { assertSerializable } from './json-data'

const MAX_MESSAGE = 1024 * 1024,
  MAX_TRANSFER = 64 * 1024 * 1024,
  CHUNK_BYTES = 128 * 1024
const methodSchema = protocolMessageSchema.shape.method
export type BridgeMethod = z.infer<typeof methodSchema>
export interface MessageTransport {
  post(message: unknown): void | PromiseLike<unknown>
  subscribe(listener: (message: unknown) => void): () => void
}
const errorSchema = z.strictObject({
  code: z.string().max(128),
  message: z.string().max(4096),
  details: z.record(z.string(), z.json()).default({}),
  retryable: z.boolean().default(false),
})
const responseSchema = z.discriminatedUnion('ok', [
  z.strictObject({
    protocolVersion: z.literal(1),
    requestId: z.string().min(1).max(256),
    sessionId: z.string().min(1).max(256),
    ok: z.literal(true),
    result: z.json(),
  }),
  z.strictObject({
    protocolVersion: z.literal(1),
    requestId: z.string().min(1).max(256),
    sessionId: z.string().min(1).max(256),
    ok: z.literal(false),
    error: errorSchema,
  }),
])
const chunkSchema = z.strictObject({
  protocolVersion: z.literal(1),
  requestId: z.string().min(1).max(256),
  sessionId: z.string().min(1).max(256),
  kind: z.literal('chunk'),
  transferId: z.string().uuid(),
  method: methodSchema,
  direction: z.enum(['request', 'response']),
  index: z.number().int().min(0).max(511),
  count: z.number().int().min(1).max(512),
  digest: z.string().regex(/^[a-f0-9]{64}$/),
  data: z.string().max(180_000),
})
const encoder = new TextEncoder(),
  decoder = new TextDecoder('utf-8', { fatal: true })
async function checksum(bytes: Uint8Array) {
  return [...new Uint8Array(await crypto.subtle.digest('SHA-256', Uint8Array.from(bytes)))]
    .map((byte) => byte.toString(16).padStart(2, '0'))
    .join('')
}
function encode(bytes: Uint8Array) {
  let binary = ''
  for (let offset = 0; offset < bytes.length; offset += 8192)
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 8192))
  return btoa(binary)
}
function decode(value: string) {
  return Uint8Array.from(atob(value), (character) => character.charCodeAt(0))
}
export class BridgeError extends Error {
  constructor(
    readonly code: string,
    message: string,
    readonly details: Record<string, unknown> = {},
    readonly retryable = false,
  ) {
    super(message)
    this.name = 'BridgeError'
  }
}
interface Pending {
  resolve(value: unknown): void
  reject(error: Error): void
  timer: ReturnType<typeof setTimeout>
}
interface Assembly {
  frame: z.infer<typeof chunkSchema>
  parts: Map<number, Uint8Array>
  size: number
  timer: ReturnType<typeof setTimeout>
}

/** JSON-only transport used identically on the host and in a Webview. */
export class JsonBridge {
  private pending = new Map<string, Pending>()
  private assemblies = new Map<string, Assembly>()
  private handlers = new Map<BridgeMethod, (payload: unknown) => unknown | Promise<unknown>>()
  private closed = false
  private release: () => void
  constructor(
    readonly sessionId: string,
    private readonly transport: MessageTransport,
    private readonly timeoutMs = 30_000,
  ) {
    this.release = transport.subscribe((message) => {
      void this.accept(message).catch(() => {
        /* Invalid peer data has no side effects. */
      })
    })
  }
  handle(method: BridgeMethod, handler: (payload: unknown) => unknown | Promise<unknown>) {
    this.handlers.set(methodSchema.parse(method), handler)
    return () => this.handlers.delete(method)
  }
  async request(method: BridgeMethod, payload: unknown): Promise<unknown> {
    if (this.closed) throw new BridgeError('view_unavailable', 'The document view has closed.')
    methodSchema.parse(method)
    assertSerializable(payload)
    const requestId = crypto.randomUUID()
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(requestId)
        reject(
          new BridgeError(
            'request_timeout',
            'The view response timed out; query transaction status before retrying a write.',
          ),
        )
      }, this.timeoutMs)
      this.pending.set(requestId, { resolve, reject, timer })
      void this.send(
        { protocolVersion: 1, requestId, sessionId: this.sessionId, method, payload },
        method,
        'request',
      ).catch((error) => this.fail(requestId, error))
    })
  }
  private fail(requestId: string, error: unknown) {
    const pending = this.pending.get(requestId)
    if (pending) {
      clearTimeout(pending.timer)
      this.pending.delete(requestId)
      pending.reject(error instanceof Error ? error : new Error(String(error)))
    }
  }
  private async send(message: unknown, method: BridgeMethod, direction: 'request' | 'response') {
    if (this.closed) throw new BridgeError('view_unavailable', 'The document view has closed.')
    assertSerializable(message)
    const bytes = encoder.encode(JSON.stringify(message))
    if (bytes.length > MAX_TRANSFER)
      throw new BridgeError('transfer_limit', 'The transfer exceeds 64 MiB.')
    if (bytes.length <= MAX_MESSAGE) {
      await this.transport.post(message)
      return
    }
    const { requestId } = message as { requestId: string },
      transferId = crypto.randomUUID(),
      digest = await checksum(bytes),
      count = Math.ceil(bytes.length / CHUNK_BYTES)
    for (let index = 0; index < count; index++) {
      if (this.closed) throw new BridgeError('view_unavailable', 'The document view has closed.')
      await this.transport.post({
        protocolVersion: 1,
        requestId,
        sessionId: this.sessionId,
        kind: 'chunk',
        transferId,
        method,
        direction,
        index,
        count,
        digest,
        data: encode(bytes.subarray(index * CHUNK_BYTES, (index + 1) * CHUNK_BYTES)),
      })
    }
  }
  async accept(raw: unknown): Promise<void> {
    if (this.closed) return
    assertSerializable(raw)
    if (encoder.encode(JSON.stringify(raw)).length > MAX_MESSAGE)
      throw new BridgeError('message_limit', 'A control message exceeds 1 MiB.')
    if (
      raw === null ||
      typeof raw !== 'object' ||
      (raw as { sessionId?: unknown }).sessionId !== this.sessionId
    )
      throw new BridgeError('session_mismatch', 'Unexpected document session.')
    if ((raw as { kind?: string }).kind === 'chunk') {
      const parsed = chunkSchema.safeParse(raw)
      if (!parsed.success) throw new BridgeError('invalid_chunk', 'Invalid transfer frame.')
      const frame = parsed.data
      if (
        frame.index >= frame.count ||
        (frame.direction === 'response' && !this.pending.has(frame.requestId))
      )
        throw new BridgeError('invalid_chunk', 'Unexpected transfer index or response.')
      let assembly = this.assemblies.get(frame.transferId)
      if (!assembly) {
        if (this.assemblies.size >= 8)
          throw new BridgeError('transfer_limit', 'Too many concurrent transfers.')
        assembly = {
          frame,
          parts: new Map(),
          size: 0,
          timer: setTimeout(() => {
            this.assemblies.delete(frame.transferId)
            this.fail(
              frame.requestId,
              new BridgeError('request_timeout', 'Incomplete chunk transfer.'),
            )
          }, this.timeoutMs),
        }
        this.assemblies.set(frame.transferId, assembly)
      }
      if (
        assembly.frame.requestId !== frame.requestId ||
        assembly.frame.count !== frame.count ||
        assembly.frame.digest !== frame.digest ||
        assembly.frame.method !== frame.method ||
        assembly.frame.direction !== frame.direction
      ) {
        this.drop(frame.transferId)
        throw new BridgeError('invalid_chunk', 'Transfer headers changed.')
      }
      const bytes = decode(frame.data)
      if (bytes.length > CHUNK_BYTES) {
        this.drop(frame.transferId)
        throw new BridgeError('chunk_limit', 'The chunk exceeds its byte limit.')
      }
      const existing = assembly.parts.get(frame.index)
      if (existing) {
        if (encode(existing) !== frame.data) {
          this.drop(frame.transferId)
          throw new BridgeError('invalid_chunk', 'A chunk index was reused with different data.')
        }
        return
      }
      assembly.parts.set(frame.index, bytes)
      assembly.size += bytes.length
      if ([...this.assemblies.values()].reduce((sum, item) => sum + item.size, 0) > MAX_TRANSFER) {
        this.drop(frame.transferId)
        throw new BridgeError('transfer_limit', 'Buffered transfers exceed 64 MiB.')
      }
      if (assembly.parts.size !== frame.count) return
      this.drop(frame.transferId)
      const complete = new Uint8Array(assembly.size)
      let offset = 0
      for (let index = 0; index < frame.count; index++) {
        const part = assembly.parts.get(index)!
        complete.set(part, offset)
        offset += part.length
      }
      if ((await checksum(complete)) !== frame.digest) {
        this.fail(
          frame.requestId,
          new BridgeError('digest_mismatch', 'The completed transfer has the wrong digest.'),
        )
        throw new BridgeError('digest_mismatch', 'The completed transfer has the wrong digest.')
      }
      const message = JSON.parse(decoder.decode(complete))
      if (
        message.sessionId !== this.sessionId ||
        message.requestId !== frame.requestId ||
        (frame.direction === 'request' && message.method !== frame.method)
      )
        throw new BridgeError('invalid_chunk', 'The assembled request does not match the transfer.')
      await this.dispatch(message)
    } else await this.dispatch(raw)
  }
  private async dispatch(raw: unknown) {
    const response = responseSchema.safeParse(raw)
    if (response.success) {
      const pending = this.pending.get(response.data.requestId)
      if (!pending) return
      clearTimeout(pending.timer)
      this.pending.delete(response.data.requestId)
      if (response.data.ok) pending.resolve(response.data.result)
      else
        pending.reject(
          new BridgeError(
            response.data.error.code,
            response.data.error.message,
            response.data.error.details,
            response.data.error.retryable,
          ),
        )
      return
    }
    const request = protocolMessageSchema.safeParse(raw)
    if (!request.success)
      throw new BridgeError('invalid_message', 'Unknown message fields, version or method.')
    const { requestId, method, payload } = request.data
    try {
      const handler = this.handlers.get(method)
      if (!handler)
        throw new BridgeError('unsupported_method', `No handler registered for ${method}.`)
      const result = await handler(payload)
      await this.send(
        {
          protocolVersion: 1,
          requestId,
          sessionId: this.sessionId,
          ok: true,
          result: result ?? null,
        },
        method,
        'response',
      )
    } catch (error) {
      const structured = errorSchema.safeParse(
        error && typeof error === 'object' && 'code' in error
          ? {
              code: String(error.code),
              message: error instanceof Error ? error.message : String(error),
              details: 'details' in error ? error.details : {},
              retryable: 'retryable' in error ? error.retryable : false,
            }
          : null,
      )
      const failure = structured.success
        ? new BridgeError(
            structured.data.code,
            structured.data.message,
            structured.data.details,
            structured.data.retryable,
          )
        : new BridgeError('request_failed', error instanceof Error ? error.message : String(error))
      await this.send(
        {
          protocolVersion: 1,
          requestId,
          sessionId: this.sessionId,
          ok: false,
          error: {
            code: failure.code,
            message: failure.message.slice(0, 4096),
            details: failure.details,
            retryable: failure.retryable,
          },
        },
        method,
        'response',
      )
    }
  }
  private drop(transferId: string) {
    const assembly = this.assemblies.get(transferId)
    if (assembly) clearTimeout(assembly.timer)
    this.assemblies.delete(transferId)
  }
  dispose() {
    if (this.closed) return
    this.closed = true
    this.release()
    for (const id of this.assemblies.keys()) this.drop(id)
    for (const requestId of this.pending.keys())
      this.fail(requestId, new BridgeError('view_unavailable', 'The document view has closed.'))
    this.handlers.clear()
  }
}
