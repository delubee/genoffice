import { build } from 'esbuild'
import { fileURLToPath } from 'node:url'
import { expect, it } from 'vitest'
// @ts-expect-error Build-only JavaScript plugin shared with Vite.
import { browserBoundary } from '../../../packages/sheets-editor/browser-boundary.mjs'

it.each([
  ['Sheets', '../../sheets/src/renderer/SheetsRibbon.tsx', 'SheetsRibbon'],
  ['Docs', '../../docs/src/renderer/components/Ribbon.tsx', 'Ribbon'],
])(
  'bundles the desktop %s ribbon for an isolated browser without AI or grid services',
  async (_, entry, exported) => {
    const boundary = browserBoundary()
    const result = await build({
      entryPoints: [fileURLToPath(new URL(entry, import.meta.url))],
      outdir: '/unused-ribbon-bundle',
      bundle: true,
      platform: 'browser',
      format: 'esm',
      loader: { '.png': 'dataurl' },
      write: false,
      metafile: true,
      plugins: [
        {
          name: 'ribbon-browser-boundary',
          setup(builder) {
            builder.onResolve({ filter: /.*/ }, ({ path }) => {
              boundary.resolveId(path)
              return undefined
            })
          },
        },
      ],
    })
    const inputs = Object.keys(result.metafile!.inputs)
    expect(result.outputFiles!.some((file) => file.text.includes(exported))).toBe(true)
    // The ribbon is presentation + host callbacks. Importing a menu constant
    // must not initialize a second grid, clipboard service or embedded AI chat.
    expect(
      inputs.filter((file) => /(?:AiChatPanel|AiComposer|sheets-ui\/|engine-formula\/)/.test(file)),
    ).toEqual([])
  },
  30_000,
)
