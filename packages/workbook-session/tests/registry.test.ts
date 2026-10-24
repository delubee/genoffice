import { expect, it } from 'vitest'
import { writeFile, link, rename } from 'node:fs/promises'
import { join } from 'node:path'
import { pathToFileURL } from 'node:url'
import { WorkbookRegistry } from '../src/registry'
import { captureFileBaseline } from '../src/publish'
import { useSessionTestResources } from './helpers'
const resources = useSessionTestResources()

it('locks real hard-link aliases while allowing independent same-name workbooks', async () => {
  const directory = await resources.directory(),
    a = join(directory, 'sales.xlsx'),
    b = join(directory, 'alias.xlsx'),
    other = join(await resources.directory(), 'sales.xlsx')
  await writeFile(a, 'A')
  await link(a, b)
  await writeFile(other, 'B')
  const registry = new WorkbookRegistry(),
    lease = registry.reserve(pathToFileURL(a).href, await captureFileBaseline(a))
  lease.bind('session-a')
  expect(() =>
    registry.reserve(pathToFileURL(b).href, {
      ...lease.baseline,
      requestedPath: b,
      canonicalPath: b,
    }),
  ).toThrow('already open')
  const independent = registry.reserve(pathToFileURL(other).href, await captureFileBaseline(other))
  independent.bind('session-b')
  expect(registry.session('session-a')?.documentUri).toBe(pathToFileURL(a).href)
  lease.release()
  const alias = registry.reserve(pathToFileURL(b).href, await captureFileBaseline(b))
  alias.bind('new-session')
  expect(registry.session('session-a')).toBeUndefined()
  expect(registry.session('new-session')?.epoch).not.toBe(lease.epoch)
})

it('updates physical ownership after atomic saves and invalidates the prior session on revert', async () => {
  const directory = await resources.directory(),
    path = join(directory, 'book.xlsx'),
    oldAlias = join(directory, 'old-alias.xlsx'),
    staged = join(directory, 'next.xlsx')
  await writeFile(path, 'old')
  await link(path, oldAlias)
  const registry = new WorkbookRegistry(),
    lease = registry.reserve(pathToFileURL(path).href, await captureFileBaseline(path))
  lease.bind('old-session')
  await writeFile(staged, 'new')
  await rename(staged, path)
  lease.update(await captureFileBaseline(path))
  const old = registry.reserve(pathToFileURL(oldAlias).href, await captureFileBaseline(oldAlias))
  old.bind('independent-old-file')
  lease.bind('reverted-session')
  expect(registry.session('old-session')).toBeUndefined()
  expect(registry.session('reverted-session')?.documentUri).toBe(pathToFileURL(path).href)
  lease.release()
  lease.release()
  expect(registry.session('independent-old-file')).toBeDefined()
})
