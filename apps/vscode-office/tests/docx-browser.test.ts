import { it, expect } from 'vitest'
import { chromium } from '@playwright/test'
import { createServer } from 'node:http'
import { readFile } from 'node:fs/promises'
import { resolve, extname } from 'node:path'
import { DocxSession } from '@genoffice/docx-session/session'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { JsonBridge } from '@genoffice/xlsx-gateway/shared/session-bridge'
import type { StepRequest } from '@genoffice/docs-editor/shared/protocol'

it('opens the built DOCX page, acknowledges typing/splits and preserves shared history across save', async () => {
  const session = await DocxSession.open(
    await buildDocx({ bodyXml: '<w:p><w:r><w:t>Hello</w:t></w:r></w:p>' }),
  )
  const initial = session.content()
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
  await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve))
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
  bridge.handle('docx.operations', (raw) =>
    session.applyOperations(
      raw as { transactionId: string; expectedRevision: number; operations: unknown },
    ),
  )
  bridge.handle('history.request', async () => {
    session.history('undo')
    await bridge.request('view.resync', projection())
    return { ready: true }
  })
  await page.exposeFunction('postToHost', (message: unknown) => incoming(message))
  try {
    const address = server.address() as { port: number }
    await page.goto(`http://127.0.0.1:${address.port}/docx.html`)
    const editor = page.locator('.tiptap')
    await editor.waitFor()
    expect(await editor.textContent()).toBe('Hello')
    await editor.click()
    await page.keyboard.press('End')
    await page.keyboard.type(' world')
    expect(await bridge.request('view.flush', {})).toEqual({ ready: true })
    expect(session.read().blocks[0]!.text).toBe('Hello world')
    await page.keyboard.press('Enter')
    await page.keyboard.type('Second')
    expect(await bridge.request('view.flush', {})).toEqual({ ready: true })
    expect(session.read().blocks.map((block) => block.text)).toEqual(['Hello world', 'Second'])
    const saved = await session.save()
    expect((await DocxSession.open(saved)).read().blocks.map((block) => block.text)).toEqual([
      'Hello world',
      'Second',
    ])
    await page.keyboard.press('Control+Home')
    await page.keyboard.press('Control+Shift+ArrowRight')
    await page.getByTitle('Bold', { exact: true }).click()
    expect(await bridge.request('view.flush', {})).toEqual({ ready: true })
    expect(
      session
        .content()
        .content![0]!.content!.some((node) => node.marks?.some((mark) => mark.type === 'bold')),
    ).toBe(true)
    const beforeFontText = await editor.textContent()
    const sizeInput = page.getByLabel('Font Size', { exact: true })
    await sizeInput.focus()
    await page.keyboard.press('Control+A')
    await page.keyboard.type('14')
    expect(await sizeInput.inputValue()).toBe('14')
    await page.keyboard.press('Enter')
    expect(await bridge.request('view.flush', {})).toEqual({ ready: true })
    const fontInput = page.getByLabel('Font', { exact: true })
    await fontInput.focus()
    await page.keyboard.press('Control+A')
    await page.keyboard.type('Times New Roman')
    expect(await fontInput.inputValue()).toBe('Times New Roman')
    expect(await editor.textContent()).toBe(beforeFontText)
    await page.keyboard.press('Enter')
    expect(await bridge.request('view.flush', {})).toEqual({ ready: true })
    expect(
      session
        .content()
        .content![0]!.content!.some((node) =>
          node.marks?.some(
            (mark) =>
              mark.type === 'docTextStyle' &&
              mark.attrs?.sizeHalfPoints === 28 &&
              mark.attrs?.font === 'Times New Roman',
          ),
        ),
    ).toBe(true)
    await page.getByLabel('Paragraph', { exact: true }).selectOption('1')
    expect(await bridge.request('view.flush', {})).toEqual({ ready: true })
    expect(session.content().content![0]!.type).toBe('docHeading')
    await page.getByLabel('Paragraph', { exact: true }).selectOption('0')
    expect(await bridge.request('view.flush', {})).toEqual({ ready: true })
    await page.getByTitle('Insert a table', { exact: true }).click()
    expect(await bridge.request('view.flush', {})).toEqual({ ready: true })
    await page.locator('.tiptap td p').first().click()
    await page.getByTitle('Insert Below', { exact: true }).click()
    expect(await bridge.request('view.flush', {})).toEqual({ ready: true })
    await page.getByTitle('Insert Right', { exact: true }).click()
    expect(await bridge.request('view.flush', {})).toEqual({ ready: true })
    const table = session.content().content!.find((node) => node.type === 'docTable')!
    expect(table.content).toHaveLength(3)
    expect(table.content![0]!.content).toHaveLength(3)
    const beforeTheme = session.content()
    await page.evaluate(() =>
      document.documentElement.style.setProperty('--vscode-editor-background', 'rgb(20, 20, 20)'),
    )
    await bridge.request('appearance.changed', { appearance: 'vscode' })
    expect(
      await page
        .locator('.office-docx')
        .evaluate((element) => getComputedStyle(element).backgroundColor),
    ).toBe('rgb(20, 20, 20)')
    expect(
      await page
        .locator('.office-docx-page')
        .evaluate((element) => getComputedStyle(element).backgroundColor),
    ).toBe('rgb(255, 255, 255)')
    expect(session.content()).toEqual(beforeTheme)
    while (session.descriptor().revision < 60) {
      try {
        session.history('undo')
      } catch (error) {
        expect(String(error)).toContain('Nothing to undo')
        break
      }
    }
    expect(session.content()).toEqual(initial)
    await bridge.request('view.resync', projection())
    expect(await editor.textContent()).toBe('Hello')
    expect(errors).toEqual([])
    expect(await page.locator('.office-docx-error').count()).toBe(0)
  } finally {
    bridge.dispose()
    await browser.close()
    await new Promise<void>((resolve) => server.close(() => resolve()))
  }
}, 60000)
