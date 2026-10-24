import { execFileSync } from 'node:child_process'
import { chmod, copyFile, mkdir } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { checkNative, nativePath } from './native-resource.mjs'

const root = fileURLToPath(new URL('../', import.meta.url))
const target = process.argv[2] ?? `${process.platform}-${process.arch}`
const output = nativePath(root, target)
if (target !== `${process.platform}-${process.arch}`)
  throw new Error(
    'Build each platform VSIX on its matching host; cross-packaging is not supported.',
  )
const engine = join(root, '../sheets/native/xlsx-engine')
execFileSync('cargo', ['build', '--locked', '--release', '--config', '.cargo/config.toml'], {
  cwd: engine,
  stdio: 'inherit',
})
await mkdir(dirname(output), { recursive: true })
await copyFile(
  join(engine, 'target/release', target.startsWith('win32') ? 'xlsx-sidecar.exe' : 'xlsx-sidecar'),
  output,
)
await chmod(output, 0o755)
await checkNative(root, target)
console.log(`Validated bundled sidecar: ${target}`)
