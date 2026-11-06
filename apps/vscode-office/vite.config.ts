import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import { fileURLToPath } from 'node:url'
// @ts-expect-error Build-only JavaScript plugin.
import { browserBoundary } from '../../packages/sheets-editor/browser-boundary.mjs'

export default defineConfig({
  root: fileURLToPath(new URL('./src/webview', import.meta.url)),
  base: './',
  plugins: [browserBoundary(), react()],
  build: {
    outDir: '../../dist/webview',
    emptyOutDir: true,
    manifest: true,
    rollupOptions: {
      input: {
        index: fileURLToPath(new URL('./src/webview/index.html', import.meta.url)),
        docx: fileURLToPath(new URL('./src/webview/docx.html', import.meta.url)),
      },
    },
  },
})
