import { spawn } from 'node:child_process'
import { createWriteStream } from 'node:fs'
import { chmod, cp, mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

if (process.platform !== 'linux' || process.arch !== 'x64')
  throw new Error('This recovery proof currently verifies Linux x64 only.')
const app = fileURLToPath(new URL('..', import.meta.url)),
  freshCrash = process.argv.includes('--fresh-crash'),
  cache = join(app, '.vscode-test'),
  root = await mkdtemp(join(cache, 'normal-recovery-')),
  extensions = join(root, 'extensions'),
  profile = join(root, 'profile'),
  workspace = join(root, 'workspace')
await Promise.all([
  mkdir(extensions, { recursive: true }),
  mkdir(join(profile, 'User'), { recursive: true }),
  mkdir(workspace, { recursive: true }),
])
await writeFile(
  join(profile, 'User', 'settings.json'),
  JSON.stringify({
    'files.hotExit': 'onExitAndWindowClose',
    'files.autoSave': 'off',
    'window.restoreWindows': 'all',
  }),
)
const product = join(extensions, 'genoffice.genoffice-workbook-0.1.0'),
  probe = join(extensions, 'genoffice-test.recovery-probe-0.1.0')
await Promise.all([mkdir(product), mkdir(probe)])
await Promise.all(
  ['package.json', 'dist', 'native'].map((name) =>
    cp(join(app, name), join(product, name), { recursive: true }),
  ),
)
await chmod(join(product, 'native', 'linux-x64', 'xlsx-sidecar'), 0o755)
await cp(join(app, 'dist/tests/proofs/normal-recovery.extension.cjs'), join(probe, 'index.cjs'))
await writeFile(
  join(probe, 'package.json'),
  JSON.stringify({
    name: 'recovery-probe',
    publisher: 'genoffice-test',
    version: '0.1.0',
    engines: { vscode: '^1.109.5' },
    main: './index.cjs',
    activationEvents: ['onStartupFinished'],
    extensionDependencies: ['genoffice.genoffice-workbook'],
  }),
)
const executable = join(cache, 'vscode-linux-x64-1.109.5', 'code')
const run = (phase) =>
  new Promise((resolve, reject) => {
    const log = createWriteStream(join(root, `${phase}.log`)),
      env = {
        ...process.env,
        GENOFFICE_RECOVERY_PROOF_ROOT: root,
        GENOFFICE_RECOVERY_PROOF_PHASE: phase,
        GENOFFICE_RECOVERY_PROOF_FRESH: freshCrash ? '1' : '0',
      }
    delete env.ELECTRON_RUN_AS_NODE
    const child = spawn(
      executable,
      [
        '--no-sandbox',
        '--disable-gpu',
        '--disable-workspace-trust',
        '--skip-welcome',
        '--skip-release-notes',
        '--disable-updates',
        `--user-data-dir=${profile}`,
        `--extensions-dir=${extensions}`,
        workspace,
      ],
      { env, stdio: ['ignore', 'pipe', 'pipe'] },
    )
    child.stdout.pipe(log, { end: false })
    child.stderr.pipe(log, { end: false })
    const timeout = setTimeout(() => {
      child.kill('SIGKILL')
      reject(new Error(`${phase} timed out; inspect ${root}`))
    }, 60_000)
    child.once('error', (error) => {
      clearTimeout(timeout)
      log.end()
      reject(error)
    })
    child.once('exit', (code, signal) => {
      clearTimeout(timeout)
      log.end()
      resolve({ code, signal })
    })
  })
const seeded = await run('seed'),
  seed = JSON.parse(await readFile(join(root, 'result.json'), 'utf8'))
if (!seed.seeded || (freshCrash ? seeded.signal !== 'SIGKILL' : seeded.code !== 0))
  throw new Error(`Recovery seed failed: ${JSON.stringify(seed)}; inspect ${root}`)
console.log('Native backup persisted; the isolated seed exited as expected.')
if (!freshCrash) {
  const crashed = await run('crash'),
    crash = JSON.parse(await readFile(join(root, 'result.json'), 'utf8'))
  if (!crash.seeded || crashed.signal !== 'SIGKILL')
    throw new Error(`Crash seed failed: ${JSON.stringify(crash)}; inspect ${root}`)
  console.log(
    'A newer native backup persisted; the isolated test application exited with expected SIGKILL.',
  )
}
const restored = await run(freshCrash ? 'recover' : 'restore'),
  result = JSON.parse(await readFile(join(root, 'result.json'), 'utf8'))
if (!result.passed || restored.code !== 0)
  throw new Error(`Recovery verification failed: ${JSON.stringify(result)}; inspect ${root}`)
await mkdir(join(app, 'test-results'), { recursive: true })
await writeFile(
  join(
    app,
    freshCrash ? 'test-results/recovery-fresh-proof.json' : 'test-results/recovery-proof.json',
  ),
  JSON.stringify(
    {
      ...result,
      platform: `${process.platform}-${process.arch}`,
      runRoot: root,
      checkedAt: new Date().toISOString(),
    },
    null,
    2,
  ),
)
console.log(`Actual native crash recovery passed; evidence: ${root}`)
