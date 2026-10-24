import { defineConfig } from '@playwright/test'

export default defineConfig({
  testDir: './tests', testMatch: '**/*.spec.ts', fullyParallel: false,
  outputDir: './test-results/browser', reporter: 'list',
  use: { headless: true },
})
