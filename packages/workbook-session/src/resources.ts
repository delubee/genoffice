import { createHash } from 'node:crypto'
import { createReadStream } from 'node:fs'
import { mkdir, mkdtemp, rm } from 'node:fs/promises'
import { join } from 'node:path'

export async function hashFile(path: string): Promise<string> {
  const hash = createHash('sha256')
  for await (const chunk of createReadStream(path)) hash.update(chunk)
  return hash.digest('hex')
}

export async function privateDirectory(parent: string, prefix: string) {
  await mkdir(parent, { recursive: true, mode: 0o700 })
  return mkdtemp(join(parent, prefix))
}

/** Every timeline/backup owner must release its reference before cleanup. */
export class RetainedResource {
  private references = 1
  private disposed = false
  private cleanup: Promise<void> | undefined
  constructor(
    readonly directory: string,
    private readonly close: () => Promise<void>,
  ) {}
  retain() {
    if (this.references === 0) throw new Error('Resource has already been released')
    this.references++
    let released = false
    return async () => {
      if (!released) {
        released = true
        await this.release()
      }
    }
  }
  async dispose() {
    if (!this.disposed) {
      this.disposed = true
      await this.release()
    }
  }
  private async release() {
    if (--this.references === 0)
      this.cleanup = this.close().finally(() =>
        rm(this.directory, { recursive: true, force: true }),
      )
    await this.cleanup
  }
}
