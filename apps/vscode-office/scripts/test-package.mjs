import { spawn, execFileSync } from 'node:child_process'
import { createWriteStream } from 'node:fs'
import { copyFile, mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { join, delimiter } from 'node:path'
import { fileURLToPath } from 'node:url'
import { createHash } from 'node:crypto'
import {
  downloadAndUnzipVSCode,
  resolveCliArgsFromVSCodeExecutablePath,
} from '@vscode/test-electron'

const app = fileURLToPath(new URL('../', import.meta.url))
const manifest = JSON.parse(await readFile(join(app, 'package.json'), 'utf8'))
const target = `${process.platform}-${process.arch}`
const vsix = join(app, 'artifacts', `genoffice-workbook-${manifest.version}-${target}.vsix`)
const root = await mkdtemp(join(app, '.vscode-test', 'package-proof-'))
const extensions = join(root, 'extensions'),
  profile = join(root, 'profile'),
  workspace = join(root, 'workspace')
await Promise.all([mkdir(extensions), mkdir(profile), mkdir(workspace)])
const executable = await downloadAndUnzipVSCode({
  version: process.argv[2] ?? '1.109.5',
  cachePath: join(app, '.vscode-test'),
})
const [cli, ...args] = resolveCliArgsFromVSCodeExecutablePath(executable, {
  reuseMachineInstall: true,
})
const env = {
  ...process.env,
  PATH: (process.env.PATH ?? '')
    .split(delimiter)
    .filter((path) => !/cargo|rustup/i.test(path))
    .join(delimiter),
  GENOFFICE_PACKAGE_PROOF_ROOT: root,
}
delete env.ELECTRON_RUN_AS_NODE
execFileSync(
  cli,
  [
    ...args,
    '--no-sandbox',
    '--user-data-dir',
    profile,
    '--extensions-dir',
    extensions,
    '--install-extension',
    vsix,
    '--force',
  ],
  { env, stdio: 'inherit', ...(process.platform === 'win32' ? { shell: true } : {}) },
)
const probe = join(root, 'probe-source')
await mkdir(probe)
await copyFile(join(app, 'dist/tests/proofs/package.extension.cjs'), join(probe, 'index.cjs'))
await writeFile(
  join(probe, 'package.json'),
  JSON.stringify({
    name: 'package-probe',
    publisher: 'genoffice-test',
    version: '0.1.0',
    engines: { vscode: '^1.109.5' },
    main: './index.cjs',
    activationEvents: ['onStartupFinished'],
    extensionDependencies: ['genoffice.genoffice-workbook'],
  }),
)
await writeFile(join(probe, 'README.md'), '# Isolated installed-package test helper\n')
const probeVsix = join(root, 'probe.vsix')
execFileSync(
  process.execPath,
  [
    fileURLToPath(import.meta.resolve('@vscode/vsce/vsce')),
    'package',
    '--no-dependencies',
    '--skip-license',
    '--allow-missing-repository',
    '--out',
    probeVsix,
  ],
  { cwd: probe, stdio: 'inherit' },
)
execFileSync(
  cli,
  [
    ...args,
    '--no-sandbox',
    '--user-data-dir',
    profile,
    '--extensions-dir',
    extensions,
    '--install-extension',
    probeVsix,
    '--force',
  ],
  { env, stdio: 'inherit', ...(process.platform === 'win32' ? { shell: true } : {}) },
)
await new Promise((resolve, reject) => {
  const child = spawn(
    executable,
    [
      '--no-sandbox',
      '--disable-gpu',
      '--disable-workspace-trust',
      '--skip-welcome',
      '--skip-release-notes',
      '--disable-updates',
      '--user-data-dir',
      profile,
      '--extensions-dir',
      extensions,
      workspace,
    ],
    { env, stdio: ['ignore', 'pipe', 'pipe'] },
  )
  const log = createWriteStream(join(root, 'host.log'))
  child.stdout.pipe(log, { end: false })
  child.stderr.pipe(log, { end: false })
  const timer = setTimeout(() => {
    child.kill()
    reject(new Error(`Packaged test timed out: ${root}`))
  }, 120000)
  child.once('error', (error) => {
    clearTimeout(timer)
    log.end()
    reject(error)
  })
  child.once('exit', (code) => {
    clearTimeout(timer)
    log.end()
    if (code === 0) resolve()
    else reject(new Error(`Packaged host exited ${code}: ${root}`))
  })
})
const result = JSON.parse(await readFile(join(root, 'result.json'), 'utf8'))
if (!result.passed) throw new Error(`Packaged test failed: ${JSON.stringify(result)}`)
await writeFile(
  join(app, 'test-results', `package-${target}-proof.json`),
  JSON.stringify(
    {
      ...result,
      vsixSha256: createHash('sha256')
        .update(await readFile(vsix))
        .digest('hex'),
      root,
      checkedAt: new Date().toISOString(),
    },
    null,
    2,
  ),
)
console.log(`Actual installed VSIX smoke test passed: ${target}`)
