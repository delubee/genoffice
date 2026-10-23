import { builtinModules } from 'node:module'

const builtins = new Set(builtinModules.map((name) => name.replace(/^node:/, '')))

/** Reject forbidden imports anywhere in the browser dependency graph. */
export function browserBoundary() {
  return {
    name: 'genoffice-browser-boundary',
    enforce: 'pre',
    resolveId(source) {
      if (
        source.startsWith('node:') || builtins.has(source) ||
        /^electron(?:\/|$)/.test(source) ||
        /^@genoffice\/(?:agent-core|ai-provider|ai-search|workbook-session)(?:\/|$)/.test(source)
      ) throw new Error(`Browser bundle cannot import ${source}`)
      return null
    },
  }
}
