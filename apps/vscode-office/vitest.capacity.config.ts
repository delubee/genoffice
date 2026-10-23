import { defineConfig } from 'vitest/config'

export default defineConfig({
  test: {
    include: ['tests/python-full-output.test.ts'],
    environment: 'node',
    testTimeout: 1_200_000,
  },
})
