import { spawn, execFileSync } from 'node:child_process'
import { createWriteStream } from 'node:fs'
import { copyFile, mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { join, delimiter } from 'node:path'
import { fileURLToPath } from 'node:url'
import { createHash } from 'node:crypto'
import { chromium } from '@playwright/test'
import { editInstalledDocx } from './installed-docx-editor.mjs'
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
      '--remote-debugging-address=127.0.0.1',
      '--remote-debugging-port=0',
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
  // Drive only this isolated installed-product profile, through its real editor.
  // Native tool writes retain their production confirmation behavior.
  let browser
  const driveWord = async () => {
    const deadline = Date.now() + 90000
    while (
      (await readFile(join(root, 'word-ready.json'), 'utf8').catch(() => '')) !== '{"ready":true}'
    ) {
      if (Date.now() > deadline) throw new Error('Installed DOCX did not become ready')
      await new Promise((resolve) => setTimeout(resolve, 100))
    }
    const port = Number(
      (await readFile(join(profile, 'DevToolsActivePort'), 'utf8')).split('\n')[0],
    )
    if (!Number.isInteger(port) || port <= 0) throw new Error('Invalid isolated debugger port')
    browser = await chromium.connectOverCDP(`http://127.0.0.1:${port}`)
    let edited = false
    const frameDeadline = Date.now() + 15000
    while (!edited && Date.now() < frameDeadline) {
      edited = await editInstalledDocx(browser, 'Installed DOCX works')
      if (!edited) await new Promise((resolve) => setTimeout(resolve, 100))
    }
    if (!edited) {
      const session = await browser.newBrowserCDPSession()
      const targets = await session.send('Target.getTargets')
      await writeFile(
        join(root, 'editor-targets.json'),
        JSON.stringify(
          {
            pages: browser.contexts().flatMap((context) =>
              context.pages().map((page) => ({
                url: page.url().split('?')[0],
                frames: page.frames().map((frame) => frame.url().split('?')[0]),
              })),
            ),
            targets: targets.targetInfos.map((target) => ({
              type: target.type,
              url: target.url.split('?')[0],
            })),
          },
          null,
          2,
        ),
      )
      throw new Error('Installed DOCX editor frame was not found')
    }
    await writeFile(join(root, 'word-edited.json'), '{"edited":true}')
  }
  void driveWord().catch((error) => {
    child.kill()
    reject(error)
  })
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
    void browser?.close().catch(() => {})
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
