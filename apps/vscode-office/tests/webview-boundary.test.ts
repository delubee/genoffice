import { expect, it } from 'vitest'
import { resolve } from 'node:path'
import { workbookHtml } from '../src/extension/webview-html'

const options = {
  root: resolve('/extension/dist/webview'),
  nonce: 'abc123',
  cspSource: 'vscode-webview://local',
  language: 'zh-CN',
  sessionId: 'session-123',
  resource: (path: string) => `resource:${path}`,
}
it('renders only packaged assets with nonce CSP and escaped locale', () => {
  const html = workbookHtml({
    ...options,
    language: 'zh-CN" onload="bad',
    html: '<html data-test="1"><head></head><script type="module" src="./assets/app.js"></script>',
  })
  expect(html).toContain("default-src 'none'")
  expect(html).toContain("connect-src 'none'")
  expect(html).toContain("script-src 'nonce-abc123'")
  expect(html).toContain('<script nonce="abc123"')
  expect(html).not.toContain('onload="bad')
  expect(html).toContain('data-workbook-appearance="excel"')
  expect(
    workbookHtml({ ...options, html: '<html lang="en"><head></head>', appearance: 'vscode' }),
  ).toContain('data-workbook-appearance="vscode"')
})
it('rejects escaping and network resources and invalid session attributes', () => {
  for (const path of [
    '../outside.js',
    '/outside.js',
    'https://example.com/app.js',
    'file:///outside',
  ])
    expect(() =>
      workbookHtml({ ...options, html: `<head><script src="${path}"></script>` }),
    ).toThrow()
  expect(() => workbookHtml({ ...options, sessionId: 'bad"', html: '<head>' })).toThrow()
})
