import { expect, it } from 'vitest'
import { nativePath, validateHeader, checkNative } from '../scripts/native-resource.mjs'

it('accepts only the three packaged host formats and architectures', () => {
  const linux = Buffer.alloc(64)
  linux.set([127, 69, 76, 70, 2, 1])
  linux.writeUInt16LE(62, 18)
  expect(() => validateHeader(linux, 'linux-x64')).not.toThrow()
  expect(() => validateHeader(linux, 'darwin-arm64')).toThrow()
  linux.writeUInt16LE(183, 18)
  expect(() => validateHeader(linux, 'linux-x64')).toThrow()
  const mac = Buffer.alloc(64)
  mac.writeUInt32LE(0xfeedfacf)
  mac.writeUInt32LE(0x0100000c, 4)
  expect(() => validateHeader(mac, 'darwin-arm64')).not.toThrow()
  const win = Buffer.alloc(128)
  win.write('MZ')
  win.writeUInt32LE(64, 60)
  win.write('PE\0\0', 64)
  win.writeUInt16LE(0x8664, 68)
  expect(() => validateHeader(win, 'win32-x64')).not.toThrow()
  win.writeUInt16LE(0xaa64, 68)
  expect(() => validateHeader(win, 'win32-x64')).toThrow()
  expect(() => validateHeader(Buffer.alloc(0), 'linux-x64')).toThrow()
  expect(() => nativePath('/extension', 'linux-arm64')).toThrow()
})

it('validates the actual bundled executable on this host', async () => {
  await expect(
    checkNative(process.cwd(), `${process.platform}-${process.arch}`),
  ).resolves.toContain('xlsx-sidecar')
})
