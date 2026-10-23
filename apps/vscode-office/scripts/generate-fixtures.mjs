import { build } from 'esbuild'
import { fileURLToPath } from 'node:url'
await build({
  entryPoints: [fileURLToPath(new URL('./generate-fixtures.ts', import.meta.url))],
  outfile: fileURLToPath(new URL('../dist/generate-fixtures.mjs', import.meta.url)),
  bundle: true,
  platform: 'node',
  target: 'node22',
  format: 'esm',
  banner: {
    js: 'import { createRequire } from "node:module"; const require = createRequire(import.meta.url);',
  },
})
await import('../dist/generate-fixtures.mjs')
