import { readFile, readdir, writeFile } from 'node:fs/promises'
import { createHash } from 'node:crypto'
import path from 'node:path'
import ts from 'typescript'

const root = path.resolve(import.meta.dirname, '..')
const docs = 'apps/docs/src/renderer/components'
const sheets = 'apps/sheets/src/renderer'
const files = []
for (const [document, directory] of [
  ['Docs', docs],
  ['Sheets', sheets],
]) {
  for (const name of (await readdir(path.join(root, directory))).sort()) {
    if (!/\.(tsx|ts)$/.test(name)) continue
    if (
      /^(Ribbon|SheetsRibbon|ExcelShell|FunctionLibraryRibbon|SortDialog|GoalSeekDialog|RemoveDuplicatesDialog|AdvancedFilterDialog|SubtotalDialog|ConsolidateDialog)\.tsx$/.test(
        name,
      ) ||
      /^ribbon-.*\.tsx$/.test(name) ||
      /context.?menu/i.test(name)
    )
      files.push({ document, file: `${directory}/${name}` })
  }
}
const sources = []
const entries = []
for (const source of files) {
  const content = await readFile(path.join(root, source.file), 'utf8')
  const ast = ts.createSourceFile(
    source.file,
    content,
    ts.ScriptTarget.Latest,
    true,
    source.file.endsWith('tsx') ? ts.ScriptKind.TSX : ts.ScriptKind.TS,
  )
  sources.push({ ...source, sha256: createHash('sha256').update(content).digest('hex') })
  const add = (node, kind, detail) => {
    const start = ast.getLineAndCharacterOfPosition(node.getStart(ast))
    const end = ast.getLineAndCharacterOfPosition(node.getEnd())
    entries.push({
      id: `${source.file}:${start.line + 1}:${start.character + 1}:${kind}`,
      ...source,
      line: start.line + 1,
      endLine: end.line + 1,
      kind,
      ...detail,
    })
  }
  const visit = (node) => {
    if (ts.isJsxOpeningElement(node) || ts.isJsxSelfClosingElement(node)) {
      const attributes = node.attributes.properties.filter(ts.isJsxAttribute)
      const actions = attributes.filter((attribute) => /^on[A-Z]/.test(attribute.name.getText(ast)))
      if (actions.length)
        add(node, 'actionable-element', {
          element: node.tagName.getText(ast),
          labels: attributes
            .filter((attribute) =>
              /^(label|title|aria-label|data-tip|name)$/.test(attribute.name.getText(ast)),
            )
            .map((attribute) => attribute.getText(ast)),
          actions: actions.map((attribute) => ({
            name: attribute.name.getText(ast),
            expression: attribute.initializer?.getText(ast) ?? '',
          })),
          constraints: attributes
            .filter((attribute) =>
              /^(disabled|hidden|visible|checked|value)$/.test(attribute.name.getText(ast)),
            )
            .map((attribute) => attribute.getText(ast)),
        })
    }
    if (ts.isObjectLiteralExpression(node)) {
      const properties = node.properties.filter(ts.isPropertyAssignment)
      const value = properties.find((property) =>
        /^(value|command|action|id)$/.test(property.name.getText(ast)),
      )
      const label = properties.find((property) =>
        /^(label|title|text)$/.test(property.name.getText(ast)),
      )
      if (value && label)
        add(node, 'menu-option-definition', {
          value: value.initializer.getText(ast),
          label: label.initializer.getText(ast),
          dynamic: !ts.isStringLiteralLike(value.initializer),
        })
    }
    if (ts.isCallExpression(node) && /(?:^|\.)onCommand$/.test(node.expression.getText(ast)))
      add(node, 'command-dispatch', {
        expression: node.getText(ast),
        dynamic: !node.arguments[0] || !ts.isStringLiteralLike(node.arguments[0]),
      })
    if (
      ts.isCallExpression(node) &&
      /^(largeMenu|smallMenu|menu)$/.test(node.expression.getText(ast))
    )
      add(node, 'menu-helper-call', { expression: node.getText(ast) })
    ts.forEachChild(node, visit)
  }
  visit(ast)
}
const artifact = {
  schemaVersion: 1,
  scope:
    'Source syntax inventory of ribbon modules and context-menu modules; entries are evidence locations, not accepted commands.',
  limitations: [
    'Dynamic variants need expansion against their providers.',
    'Dialogs, supporting panes, keyboard routes and imported helpers need separate review.',
    'No plugin adapter or persistence is inferred from a matching command string.',
  ],
  sources,
  entries,
}
const json = `${JSON.stringify(artifact, null, 2)}\n`
const destination = path.join(root, 'specs/003-vscode-desktop-menus/control-inventory.json')
if (process.argv.includes('--check')) {
  if ((await readFile(destination, 'utf8')) !== json)
    throw new Error(
      'Desktop control inventory is stale; run node scripts/inventory-desktop-menus.mjs',
    )
} else await writeFile(destination, json)
console.log(
  `${sources.length} source modules; ${entries.length} evidence entries; ${entries.filter((entry) => entry.dynamic).length} dynamic dispatch/option entries require review.`,
)
