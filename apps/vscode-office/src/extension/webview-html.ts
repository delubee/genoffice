import { relative, resolve, isAbsolute } from 'node:path'
import { WorkbookError } from '@genoffice/workbook-session/domain'

export function workbookHtml(input: {
  html: string
  root: string
  nonce: string
  cspSource: string
  language: string
  sessionId: string
  resource(path: string): string
}) {
  if (!/^[A-Za-z0-9+/=]+$/.test(input.nonce) || !/^[A-Za-z0-9-]+$/.test(input.sessionId))
    throw new Error('Invalid trusted HTML identity.')
  const csp = `default-src 'none'; img-src ${input.cspSource} data: blob:; style-src ${input.cspSource} 'unsafe-inline'; font-src ${input.cspSource} data:; script-src 'nonce-${input.nonce}' ${input.cspSource}; worker-src blob:; connect-src 'none';`
  return input.html
    .replace(/(src|href)="([^"]+)"/g, (_full, attribute: string, path: string) => {
      const absolute = resolve(input.root, path),
        child = relative(input.root, absolute)
      if (child.startsWith('..') || isAbsolute(child) || /^[a-z][a-z0-9+.-]*:/i.test(path))
        throw new WorkbookError(
          'invalid_resource',
          'Webview assets must be inside the packaged resource root.',
        )
      return `${attribute}="${input.resource(absolute)}"`
    })
    .replace(
      '<html ',
      `<html lang="${input.language.replace(/[^a-zA-Z0-9-]/g, '')}" data-session-id="${input.sessionId}" `,
    )
    .replace('<head>', `<head><meta http-equiv="Content-Security-Policy" content="${csp}">`)
    .replace(/<script /g, `<script nonce="${input.nonce}" `)
}
