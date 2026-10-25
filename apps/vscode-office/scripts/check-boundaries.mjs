import { readFile, readdir } from 'node:fs/promises'
import { resolve } from 'node:path'

const root = resolve('../..')
const forbidden =
  /^(?:node:|electron(?:\/|$)|@genoffice\/(?:agent-core|ai-provider|ai-search|workbook-session)(?:\/|$))/
async function walk(directory) {
  for (const entry of await readdir(directory, { withFileTypes: true })) {
    const path = resolve(directory, entry.name)
    if (entry.isDirectory()) await walk(path)
    else if (/\.[cm]?[jt]sx?$/.test(entry.name)) {
      const source = await readFile(path, 'utf8')
      for (const match of source.matchAll(
        /(?:\bfrom\s*|\bimport\s*(?:\(\s*)?|\brequire\s*\(\s*)['"]([^'"]+)/g,
      )) {
        if (
          !['@genoffice/agent-core/redact-range', '@genoffice/agent-core/redact-core'].includes(
            match[1],
          ) &&
          forbidden.test(match[1])
        )
          throw new Error(`Forbidden browser import ${match[1]} in ${path}`)
      }
    }
  }
}
await walk(resolve(root, 'packages/sheets-editor/src'))
await walk(resolve(root, 'packages/docs-editor/src'))
await walk(resolve(root, 'apps/vscode-office/src/webview'))
const extensionGraph = JSON.parse(
  await readFile(resolve(root, 'apps/vscode-office/dist/extension.meta.json'), 'utf8'),
)
if (
  Object.keys(extensionGraph.inputs).some((path) =>
    /(?:packages\/(?:docx-engine|docx-session)\/|apps\/docs\/src\/renderer\/)/.test(path),
  )
)
  throw new Error('DOCX implementation must remain outside the eager extension bundle.')
