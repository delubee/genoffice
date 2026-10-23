import { readFile, writeFile } from 'node:fs/promises'
import { toolContributions } from '../dist/tool-schema.cjs'

const path = new URL('../package.json', import.meta.url)
const manifest = JSON.parse(await readFile(path, 'utf8'))
if (process.argv.includes('--write')) {
  manifest.contributes.languageModelTools = toolContributions
  await writeFile(path, `${JSON.stringify(manifest, null, 2)}\n`)
} else if (
  JSON.stringify(manifest.contributes.languageModelTools) !== JSON.stringify(toolContributions)
) {
  throw new Error(
    'Native tool schemas are stale. Run node scripts/generate-tools.mjs --write after the backend build.',
  )
}
