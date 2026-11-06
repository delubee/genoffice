import { realpath, readlink, stat } from 'node:fs/promises'
import { win32 } from 'node:path'

const unsupported = new Set(['UNKNOWN', 'ENOSYS', 'ENOTSUP', 'EINVAL'])
const code = (error: unknown) => (error as NodeJS.ErrnoException)?.code ?? ''

/**
 * User-selected documents can live on Windows virtual filesystems which allow
 * reads but cannot answer realpath. This fallback is NOT a containment check:
 * workspace exchange and recovery paths must keep their strict realpath rules.
 */
export async function resolveDocumentPath(path: string, platform = process.platform) {
  try {
    return await realpath(path)
  } catch (error) {
    if (platform !== 'win32' || !unsupported.has(code(error))) throw error
  }
  const requested = win32.resolve(path)
  // Do not turn missing files, inaccessible mounts or broken junctions into
  // apparently valid paths just because normalization itself is possible.
  await stat(requested)
  let pending = requested
  const visited = new Set<string>()
  for (let hops = 0; hops < 40; hops++) {
    const key = pending.toLowerCase()
    if (visited.has(key)) break
    visited.add(key)
    const root = win32.parse(pending).root
    const parts = pending.slice(root.length).split(win32.sep).filter(Boolean)
    let current = root,
      redirected = false
    for (const [index, part] of parts.entries()) {
      current = win32.join(current, part)
      let target: string
      try {
        target = await readlink(current)
      } catch (error) {
        if (unsupported.has(code(error))) continue
        if (code(error) === 'ENOENT' && pending !== requested) {
          await stat(requested)
          return requested
        }
        throw error
      }
      pending = win32.resolve(win32.dirname(current), target, ...parts.slice(index + 1))
      redirected = true
      break
    }
    if (!redirected) {
      try {
        await stat(pending)
        return pending
      } catch {
        // Some mount points expose an NT device target which Node cannot use
        // as a path. Recheck the usable user path instead of inventing one.
        await stat(requested)
        return requested
      }
    }
  }
  throw Object.assign(new Error('Too many links while resolving the document path.'), {
    code: 'ELOOP',
  })
}
