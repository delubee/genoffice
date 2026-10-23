import { defineConfig } from '@vscode/test-cli'

export default defineConfig([
  {
    label: 'minimum',
    version: '1.109.5',
    extensionDevelopmentPath: ['.', 'tests/fixtures/api-probe'],
    files: 'dist/tests/integration/**/*.test.cjs',
    workspaceFolder: './test-workspace',
    launchArgs: ['--disable-gpu', '--no-sandbox'],
    mocha: { timeout: 30_000 },
  },
  {
    label: 'stable',
    version: 'stable',
    extensionDevelopmentPath: ['.', 'tests/fixtures/api-probe'],
    files: 'dist/tests/integration/**/*.test.cjs',
    workspaceFolder: './test-workspace',
    launchArgs: ['--disable-gpu', '--no-sandbox'],
    mocha: { timeout: 30_000 },
  },
])
