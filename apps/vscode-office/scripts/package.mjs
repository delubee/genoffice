import { execFileSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import { copyFile, mkdir, readFile, readdir, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import JSZip from 'jszip'
import { checkNative, validateHeader } from './native-resource.mjs'

const root = fileURLToPath(new URL('../', import.meta.url))
const target = process.argv[2] ?? `${process.platform}-${process.arch}`
if (target !== `${process.platform}-${process.arch}`)
  throw new Error('Package on the matching platform host.')
await checkNative(root, target)
const nativeDirs = await readdir(join(root, 'native'))
if (nativeDirs.some((entry) => entry !== target))
  throw new Error('Remove other platform native resources before packaging.')
await copyFile(join(root, '../../LICENSE'), join(root, 'LICENSE'))
await copyFile(
  join(root, '../shell/build/THIRD-PARTY-NOTICES.txt'),
  join(root, 'THIRD-PARTY-NOTICES.txt'),
)
const manifest = JSON.parse(await readFile(join(root, 'package.json'), 'utf8'))
await mkdir(join(root, 'artifacts'), { recursive: true })
const name = `genoffice-workbook-${manifest.version}-${target}.vsix`
const output = join(root, 'artifacts', name)
const cli = fileURLToPath(import.meta.resolve('@vscode/vsce/vsce'))
execFileSync(
  process.execPath,
  [
    cli,
    'package',
    '--no-dependencies',
    '--allow-missing-repository',
    '--skip-license',
    '--baseContentUrl',
    'https://github.com/genspark-ai/genoffice/blob/main/apps/vscode-office/',
    '--baseImagesUrl',
    'https://github.com/genspark-ai/genoffice/raw/main/apps/vscode-office/',
    '--target',
    target,
    '--out',
    output,
  ],
  { cwd: root, stdio: 'inherit' },
)
const bytes = await readFile(output),
  zip = await JSZip.loadAsync(bytes)
const files = Object.keys(zip.files).filter((name) => !zip.files[name].dir)
for (const required of [
  'package.json',
  'dist/extension.cjs',
  'dist/webview/index.html',
  'skills/genoffice-workbook/SKILL.md',
  'skills/genoffice-workbook/scripts/analyze.py',
  'LICENSE.txt',
  'THIRD-PARTY-NOTICES.txt',
])
  if (!zip.file(`extension/${required}`)) throw new Error(`Missing packaged resource: ${required}`)
if (
  files.some((name) =>
    /(?:node_modules|tests|test-results|test-workspace|\.vscode-test|\.env)(?:\/|$)/.test(name),
  )
)
  throw new Error('Development/private resources leaked into the VSIX.')
const binary = zip.file(
  `extension/native/${target}/${target.startsWith('win32') ? 'xlsx-sidecar.exe' : 'xlsx-sidecar'}`,
)
if (!binary) throw new Error('Missing packaged sidecar.')
validateHeader(await binary.async('nodebuffer'), target)
if (!target.startsWith('win32') && (Number(binary.unixPermissions) & 0o111) === 0)
  throw new Error('Packaged sidecar lost executable permissions.')
const digest = createHash('sha256').update(bytes).digest('hex')
await writeFile(`${output}.sha256`, `${digest}  ${name}\n`)
await writeFile(
  join(root, 'artifacts/package-proof.json'),
  JSON.stringify(
    { target, version: manifest.version, sha256: digest, bytes: bytes.length, files },
    null,
    2,
  ),
)
console.log(`Verified ${name}: ${digest}`)
