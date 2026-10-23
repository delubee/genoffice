import { execFileSync } from 'node:child_process'
import { fileURLToPath } from 'node:url'
const root = fileURLToPath(new URL('../', import.meta.url))
const cli = fileURLToPath(new URL('./bin.mjs', import.meta.resolve('@vscode/test-cli')))
execFileSync(
  process.execPath,
  [
    cli,
    '--label',
    process.argv[2] ?? 'minimum',
    '--grep',
    'Workbook performance samples',
    '--fail-zero',
  ],
  {
    cwd: root,
    stdio: 'inherit',
    env: { ...process.env, GENOFFICE_BENCHMARK: '1' },
  },
)
