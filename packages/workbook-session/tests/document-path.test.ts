import { beforeEach, expect, it, vi } from 'vitest'
import { realpath, readlink, stat } from 'node:fs/promises'
import { resolveDocumentPath } from '../src/document-path'

vi.mock('node:fs/promises', () => ({ realpath: vi.fn(), readlink: vi.fn(), stat: vi.fn() }))
const error = (code: string) => Object.assign(new Error(code), { code })
beforeEach(() => {
  vi.resetAllMocks()
  vi.mocked(realpath).mockRejectedValue(error('UNKNOWN'))
  vi.mocked(readlink).mockRejectedValue(error('EINVAL'))
  vi.mocked(stat).mockResolvedValue({} as Awaited<ReturnType<typeof stat>>)
})

it('keeps native canonical paths when available', async () => {
  vi.mocked(realpath).mockResolvedValue('R:\\Reports\\book.xlsx')
  expect(await resolveDocumentPath('E:\\work\\cloud\\book.xlsx', 'win32')).toBe(
    'R:\\Reports\\book.xlsx',
  )
  expect(readlink).not.toHaveBeenCalled()
})
it('resolves a junction to its accessible virtual-drive target when realpath is unsupported', async () => {
  vi.mocked(readlink).mockImplementation(async (path) => {
    if (path === 'E:\\work\\cloud') return 'R:\\Reports'
    throw error('EINVAL')
  })
  expect(await resolveDocumentPath('E:\\work\\cloud\\book.xlsx', 'win32')).toBe(
    'R:\\Reports\\book.xlsx',
  )
  expect(stat).toHaveBeenCalledWith('R:\\Reports\\book.xlsx')
})
it('accepts an accessible virtual-drive path without junctions', async () => {
  expect(await resolveDocumentPath('R:\\Reports\\book.xlsx', 'win32')).toBe(
    'R:\\Reports\\book.xlsx',
  )
})
it('does not suppress permission, missing-file or non-Windows realpath errors', async () => {
  for (const code of ['ENOENT', 'EACCES', 'EPERM', 'EIO', 'ELOOP']) {
    vi.mocked(realpath).mockRejectedValue(error(code))
    await expect(resolveDocumentPath('E:\\work\\book.xlsx', 'win32')).rejects.toMatchObject({
      code,
    })
  }
  vi.mocked(realpath).mockRejectedValue(error('UNKNOWN'))
  await expect(resolveDocumentPath('/work/book.xlsx', 'linux')).rejects.toMatchObject({
    code: 'UNKNOWN',
  })
  vi.mocked(stat).mockRejectedValue(error('ENOENT'))
  await expect(resolveDocumentPath('R:\\book.xlsx', 'win32')).rejects.toMatchObject({
    code: 'ENOENT',
  })
})
it('rejects cyclic junctions instead of returning an unverified path', async () => {
  vi.mocked(readlink).mockImplementation(async (path) => {
    if (path === 'E:\\a') return 'E:\\b'
    if (path === 'E:\\b') return 'E:\\a'
    throw error('EINVAL')
  })
  await expect(resolveDocumentPath('E:\\a\\book.xlsx', 'win32')).rejects.toMatchObject({
    code: 'ELOOP',
  })
})

it('keeps the accessible user path when a mount exposes an unusable device target', async () => {
  vi.mocked(readlink).mockImplementation(async (path) => {
    if (path === 'E:\\work\\cloud') return '\\\\?\\Volume{opaque}\\'
    if (String(path).startsWith('\\\\?\\')) throw error('ENOENT')
    throw error('EINVAL')
  })
  expect(await resolveDocumentPath('E:\\work\\cloud\\book.xlsx', 'win32')).toBe(
    'E:\\work\\cloud\\book.xlsx',
  )
})
