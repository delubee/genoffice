import { builtinModules } from 'node:module'

const builtins = new Set(builtinModules.map((name) => name.replace(/^node:/, '')))

/** Reject forbidden imports anywhere in the browser dependency graph. */
export function browserBoundary() {
  return {
    name: 'genoffice-browser-boundary',
    enforce: 'pre',
    resolveId(source) {
      // Pure constants/range redaction helpers shared by DOCX parsing; this
      // explicit subpath imports no agent runtime, provider, process or network.
      // The ai-provider `browser` entry is its renderer-safe settings surface (types and
      // pure helpers; its own header keeps Node-backed transports out).
      if (
        [
          '@genoffice/agent-core/redact-range',
          '@genoffice/agent-core/redact-core',
          '@genoffice/ai-provider/browser',
        ].includes(source)
      )
        return null
      if (
        source.startsWith('node:') ||
        builtins.has(source) ||
        /^electron(?:\/|$)/.test(source) ||
        /^@genoffice\/(?:agent-core|ai-provider|ai-search|workbook-session)(?:\/|$)/.test(source)
      )
        throw new Error(`Browser bundle cannot import ${source}`)
      return null
    },
  }
}
