import { expect, it, vi } from 'vitest'
import { requestHostHistory } from '../src/history-bridge'
import type { EditorHost } from '../src/ports'

it('waits for pending input to flush and issues exactly one host history request', async () => {
  let finish!: (ready: boolean) => void
  const flushed = new Promise<boolean>((resolve) => {
      finish = resolve
    }),
    history = vi.fn(async () => {})
  const request = requestHostHistory({ history } as unknown as EditorHost, 'undo', () => flushed)
  await Promise.resolve()
  expect(history).not.toHaveBeenCalled()
  finish(true)
  await request
  expect(history).toHaveBeenCalledExactlyOnceWith('undo')
})

it('preserves IME or failed input without issuing an undo/redo command', async () => {
  const history = vi.fn(async () => {})
  await expect(
    requestHostHistory({ history } as unknown as EditorHost, 'redo', async () => false),
  ).rejects.toThrow('editor_busy')
  expect(history).not.toHaveBeenCalled()
})
