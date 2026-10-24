import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, vi } from 'vitest'

export function useSessionTestResources() {
  const directories = new Set<string>()
  afterEach(async () => {
    vi.useRealTimers()
    await Promise.all([...directories].map((path) => rm(path, { recursive: true, force: true })))
    directories.clear()
  })
  return {
    async directory() {
      const path = await mkdtemp(join(tmpdir(), 'genoffice-session-test-'))
      directories.add(path)
      return path
    },
    freezeTime(iso = '2026-10-21T00:00:00.000Z') {
      vi.useFakeTimers({ toFake: ['Date'] })
      vi.setSystemTime(new Date(iso))
    },
  }
}

export function failAt(expectedPoint: string) {
  return (point: string) => {
    if (point === expectedPoint) throw new Error(`Injected failure: ${point}`)
  }
}
