import { defineConfig } from 'vitest/config'

export default defineConfig({
  test: {
    include: ['tests/**/*.test.ts'],
    exclude: [
      'tests/integration/**',
      'tests/python-full-output.test.ts',
      'tests/docx-browser.test.ts',
    ],
    environment: 'node',
    testTimeout: 20_000,
  },
})
