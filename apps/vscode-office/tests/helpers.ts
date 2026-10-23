import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

export async function withTestWorkspace<T>(run: (path: string) => Promise<T>): Promise<T> {
  const path = await mkdtemp(join(tmpdir(), 'genoffice-workspace-test-'))
  try {
    return await run(path)
  } finally {
    await rm(path, { recursive: true, force: true })
  }
}
