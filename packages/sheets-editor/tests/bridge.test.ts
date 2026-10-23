import { expect, it } from 'vitest'
import { JsonBridge } from '../src/bridge'

function pair() {
  let left: (message: unknown) => void = () => {},
    right: (message: unknown) => void = () => {}
  const a = new JsonBridge('session', {
    post: (value) => {
      queueMicrotask(() => right(value))
    },
    subscribe: (listener) => {
      left = listener
      return () => {
        left = () => {}
      }
    },
  })
  const b = new JsonBridge('session', {
    post: (value) => {
      queueMicrotask(() => left(value))
    },
    subscribe: (listener) => {
      right = listener
      return () => {
        right = () => {}
      }
    },
  })
  return { a, b }
}
it('transfers chunked JSON only after complete digest validation', async () => {
  const { a, b } = pair()
  try {
    b.handle('candidate.prepare', (payload) => ({ echo: payload }))
    const value = { text: '汉字🙂'.repeat(100_000) }
    expect(await a.request('candidate.prepare', value)).toEqual({ echo: value })
  } finally {
    a.dispose()
    b.dispose()
  }
})
it('rejects unknown sessions, methods, extra fields, executable JSON and disposed views', async () => {
  const { a, b } = pair()
  try {
    await expect(a.request('source.read', { path: '/etc/passwd' })).rejects.toMatchObject({
      code: 'unsupported_method',
    })
    await expect(a.request('source.read', { run() {} })).rejects.toThrow()
    b.dispose()
    await expect(b.request('source.read', {})).rejects.toMatchObject({ code: 'view_unavailable' })
  } finally {
    a.dispose()
    b.dispose()
  }
})

it('rejects wrong-session/unknown fields and corrupted complete transfers before running handlers', async () => {
  const { a, b } = pair()
  try {
    await expect(
      b.accept({
        protocolVersion: 1,
        requestId: 'foreign',
        sessionId: 'another-session',
        method: 'source.read',
        payload: {},
      }),
    ).rejects.toMatchObject({ code: 'session_mismatch' })
    await expect(
      b.accept({
        protocolVersion: 1,
        requestId: 'extra',
        sessionId: 'session',
        method: 'source.read',
        payload: {},
        exec: 'python',
      }),
    ).rejects.toMatchObject({ code: 'invalid_message' })
  } finally {
    a.dispose()
    b.dispose()
  }
  const frames: Record<string, any>[] = []
  let ready: () => void = () => {}
  const complete = new Promise<void>((resolve) => {
    ready = resolve
  })
  const sender = new JsonBridge('session', {
    post(value) {
      const frame = value as Record<string, any>
      frames.push(frame)
      if (frame.index + 1 === frame.count) ready()
    },
    subscribe() {
      return () => {}
    },
  })
  const receiver = new JsonBridge('session', {
    post() {},
    subscribe() {
      return () => {}
    },
  })
  let calls = 0
  receiver.handle('candidate.prepare', () => {
    calls++
    return null
  })
  const pending = sender
    .request('candidate.prepare', { text: 'x'.repeat(1_100_000) })
    .catch(() => {})
  try {
    await complete
    frames[0]!.data = (frames[0]!.data[0] === 'A' ? 'B' : 'A') + frames[0]!.data.slice(1)
    for (const frame of frames.slice(0, -1)) await receiver.accept(frame)
    await expect(receiver.accept(frames.at(-1))).rejects.toMatchObject({ code: 'digest_mismatch' })
    expect(calls).toBe(0)
  } finally {
    sender.dispose()
    receiver.dispose()
    await pending
  }
})
