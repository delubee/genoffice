import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import { fileURLToPath } from 'node:url'
// @ts-expect-error Build-only JavaScript plugin.
import { browserBoundary } from './browser-boundary.mjs'

export default defineConfig({
  plugins: [browserBoundary(), react()],
  build: {
    lib: { entry: fileURLToPath(new URL('./src/index.ts', import.meta.url)), formats: ['es'], fileName: 'index' },
    rollupOptions: { external: ['react', 'react-dom', 'react/jsx-runtime'] },
  },
})
