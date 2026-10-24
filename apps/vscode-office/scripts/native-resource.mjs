import { open, access } from 'node:fs/promises'
import { constants } from 'node:fs'
import { join } from 'node:path'

export const targets = ['linux-x64', 'win32-x64', 'darwin-arm64']
export function nativePath(root, target) {
  if (!targets.includes(target)) throw new Error(`Unsupported extension host: ${target}`)
  return join(root, 'native', target, target === 'win32-x64' ? 'xlsx-sidecar.exe' : 'xlsx-sidecar')
}
export function validateHeader(bytes, target) {
  if (!targets.includes(target)) throw new Error(`Unsupported extension host: ${target}`)
  let valid = false
  if (target === 'linux-x64')
    valid =
      bytes.length >= 20 &&
      bytes.subarray(0, 4).equals(Buffer.from([127, 69, 76, 70])) &&
      bytes[4] === 2 &&
      bytes[5] === 1 &&
      bytes.readUInt16LE(18) === 62
  if (target === 'darwin-arm64')
    valid =
      bytes.length >= 8 &&
      bytes.readUInt32LE(0) === 0xfeedfacf &&
      bytes.readUInt32LE(4) === 0x0100000c
  if (target === 'win32-x64' && bytes.length >= 64 && bytes.toString('ascii', 0, 2) === 'MZ') {
    const offset = bytes.readUInt32LE(60)
    valid =
      offset + 6 <= bytes.length &&
      bytes.toString('ascii', offset, offset + 4) === 'PE\0\0' &&
      bytes.readUInt16LE(offset + 4) === 0x8664
  }
  if (!valid)
    throw new Error(
      `The packaged xlsx-sidecar has an invalid ${target} binary header. Reinstall the matching platform VSIX.`,
    )
}
export async function checkNative(root, target) {
  const path = nativePath(root, target)
  const file = await open(path, 'r')
  try {
    if (!(await file.stat()).isFile())
      throw new Error('The packaged sidecar is not a regular file.')
    const buffer = Buffer.alloc(65536)
    const { bytesRead } = await file.read(buffer, 0, buffer.length, 0)
    validateHeader(buffer.subarray(0, bytesRead), target)
  } finally {
    await file.close()
  }
  await access(path, target.startsWith('win32') ? constants.R_OK : constants.X_OK)
  return path
}
