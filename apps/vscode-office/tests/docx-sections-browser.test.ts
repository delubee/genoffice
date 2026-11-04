import { it, expect } from 'vitest'
import { chromium } from '@playwright/test'
import { createServer } from 'node:http'
import { readFile } from 'node:fs/promises'
import { resolve, extname } from 'node:path'
import { DocxSession } from '@genoffice/docx-session/session'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { JsonBridge } from '@genoffice/xlsx-gateway/shared/session-bridge'
import type { StepRequest } from '@genoffice/docs-editor/shared/protocol'

const PG_MAR =
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>'

it('edits the page setup and page numbering of the section under the cursor', async () => {
  const session = await DocxSession.open(
    await buildDocx({
      bodyXml:
        '<w:p><w:r><w:t>First section</w:t></w:r></w:p>' +
        `<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/>${PG_MAR}</w:sectPr></w:pPr></w:p>` +
        '<w:p><w:r><w:t>Second section</w:t></w:r></w:p>',
    }),
  )
  const root = resolve('dist/webview')
  const server = createServer(async (request, response) => {
    try {
      const path = resolve(root, '.' + new URL(request.url!, 'http://localhost').pathname)
      if (!path.startsWith(root + '/')) throw new Error('Invalid resource')
      const bytes = await readFile(path)
      response.setHeader(
        'content-type',
        (
          { '.js': 'text/javascript', '.css': 'text/css', '.html': 'text/html' } as Record<
            string,
            string
          >
        )[extname(path)] ?? 'application/octet-stream',
      )
      response.end(
        extname(path) === '.html'
          ? bytes
              .toString()
              .replace('<html ', `<html data-session-id="${session.sessionId}" lang="en" `)
              .replace(
                '<head>',
                '<head><script>window.acquireVsCodeApi=()=>({postMessage:m=>window.postToHost(m)})</script>',
              )
          : bytes,
      )
    } catch {
      response.statusCode = 404
      response.end()
    }
  })
  await new Promise<void>((done) => server.listen(0, '127.0.0.1', done))
  const browser = await chromium.launch({ headless: true })
  const page = await browser.newPage()
  const errors: string[] = []
  page.on('pageerror', (error) => errors.push(String(error)))
  let incoming: (value: unknown) => void = () => {}
  const bridge = new JsonBridge(session.sessionId, {
    post: (message) =>
      page.evaluate(
        (value) => window.dispatchEvent(new MessageEvent('message', { data: value })),
        message,
      ),
    subscribe(listener) {
      incoming = listener
      return () => {}
    },
  })
  const projection = () => ({
    descriptor: session.descriptor(),
    content: session.content(),
    presentation: session.presentation(),
  })
  bridge.handle('view.ready', projection)
  bridge.handle('view.resync', projection)
  bridge.handle('view.presentation', () => ({ ready: true }))
  bridge.handle('view.intent', (raw) => session.submit(raw as StepRequest))
  bridge.handle('docx.settings', (raw) =>
    session.applySettings(
      raw as { transactionId: string; expectedRevision: number; settings: unknown },
    ),
  )
  await page.exposeFunction('postToHost', (message: unknown) => incoming(message))
  try {
    await page.goto(`http://127.0.0.1:${(server.address() as { port: number }).port}/docx.html`)
    await page.locator('.office-docx-page').waitFor()
    expect(session.presentation().settings!.sections).toHaveLength(2)
    // Put the caret in the first section and narrow its margins: the earlier section changes, not the last.
    await page.getByText('First section', { exact: true }).click()
    await page
      .locator('.ribbon-tab', { hasText: /^Layout$/ })
      .first()
      .click()
    await page.locator('[data-tip="Margins"]').first().click()
    await page.locator('[data-rb-panel] button', { hasText: 'Narrow' }).first().click()
    expect(await bridge.request('view.flush', {})).toEqual({ ready: true })
    await expect
      .poll(() => session.presentation().settings!.sections![0]!.section.marginLeft)
      .toBe(720)
    const settings = session.presentation().settings!
    expect(settings.sections![1]!.section.marginLeft).toBe(1440)
    expect(settings.section!.marginLeft).toBe(1440)
    expect(errors).toEqual([])
  } finally {
    await browser.close()
    await new Promise<void>((done) => server.close(() => done()))
  }
}, 60000)
