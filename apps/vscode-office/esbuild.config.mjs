import { build } from 'esbuild'
import { mkdir, readdir } from 'node:fs/promises'

await mkdir('test-workspace', { recursive: true })
await build({
  entryPoints: ['src/extension/tools/manifest.ts'],
  outfile: 'dist/tool-schema.cjs',
  bundle: true,
  platform: 'node',
  target: 'node22',
  format: 'cjs',
})

await build({
  entryPoints: ['src/extension/activate.ts'],
  outfile: 'dist/extension.cjs',
  bundle: true,
  platform: 'node',
  target: 'node22',
  format: 'cjs',
  external: ['vscode'],
  sourcemap: true,
  metafile: true,
}).then(async (result) => {
  const { writeFile } = await import('node:fs/promises')
  await writeFile('dist/extension.meta.json', JSON.stringify(result.metafile))
})

const tests = await readdir('tests/integration').catch(() => [])
const entries = tests
  .filter((name) => name.endsWith('.test.ts'))
  .map((name) => `tests/integration/${name}`)
if (entries.length)
  await build({
    entryPoints: entries,
    outdir: 'dist/tests/integration',
    outExtension: { '.js': '.cjs' },
    bundle: true,
    platform: 'node',
    target: 'node22',
    format: 'cjs',
    external: ['vscode', 'mocha'],
    sourcemap: true,
  })

await build({
  entryPoints: ['tests/fixtures/api-probe/index.ts'],
  outfile: 'tests/fixtures/api-probe/index.cjs',
  bundle: true,
  platform: 'node',
  format: 'cjs',
  external: ['vscode'],
})

await build({
  entryPoints: ['tests/proofs/normal-recovery.extension.ts', 'tests/proofs/package.extension.ts'],
  outdir: 'dist/tests/proofs',
  outExtension: { '.js': '.cjs' },
  bundle: true,
  platform: 'node',
  target: 'node22',
  format: 'cjs',
  external: ['vscode'],
})

const { copyFile, chmod } = await import('node:fs/promises')
const { join } = await import('node:path')
const nativeDirectory = join('native', `${process.platform}-${process.arch}`)
const binary = process.platform === 'win32' ? 'xlsx-sidecar.exe' : 'xlsx-sidecar'
await mkdir(nativeDirectory, { recursive: true })
await copyFile(
  join('..', 'sheets', 'native', 'xlsx-engine', 'target', 'release', binary),
  join(nativeDirectory, binary),
)
await chmod(join(nativeDirectory, binary), 0o755)
