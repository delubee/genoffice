import { defineConfig } from '@vscode/test-cli'

const cdpPort = process.env.GENOFFICE_TEST_CDP_PORT
if (cdpPort && (!/^\d+$/.test(cdpPort) || Number(cdpPort) < 1024 || Number(cdpPort) > 65535))
  throw new Error('Invalid test CDP port')
const cdpArgs = cdpPort
  ? [`--remote-debugging-port=${cdpPort}`, '--remote-debugging-address=127.0.0.1']
  : []

export default defineConfig([
  {
    label: 'minimum',
    version: '1.109.5',
    extensionDevelopmentPath: ['.', 'tests/fixtures/api-probe'],
    files: 'dist/tests/integration/**/*.test.cjs',
    workspaceFolder: './test-workspace',
    launchArgs: ['--disable-gpu', '--no-sandbox', ...cdpArgs],
    mocha: { timeout: 30_000 },
  },
  {
    label: 'stable',
    version: 'stable',
    extensionDevelopmentPath: ['.', 'tests/fixtures/api-probe'],
    files: 'dist/tests/integration/**/*.test.cjs',
    workspaceFolder: './test-workspace',
    launchArgs: ['--disable-gpu', '--no-sandbox', ...cdpArgs],
    mocha: { timeout: 30_000 },
  },
])
