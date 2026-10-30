import { it, expect } from 'vitest'
import { chromium } from '@playwright/test'
import { createServer } from 'node:http'
import { readFile } from 'node:fs/promises'
import { resolve, extname } from 'node:path'
import { DocxSession } from '@genoffice/docx-session/session'
import { DocumentBindings } from '@genoffice/docx-session/bindings'
import { officeActionHandler } from '../src/extension/office-action-handler'
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
  bridge.handle('docx.settings', (raw) =>
    session.applySettings(
      raw as { transactionId: string; expectedRevision: number; settings: unknown },
    ),
  )
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
  const bindings = new DocumentBindings()
  let menuSave: Uint8Array | undefined
  const chats: { query: string; partial: boolean; scope: string; bindingId: string }[] = []
  let capturedScope = 'document',
    capturedBinding = ''
  bridge.handle(
    'office.action',
    officeActionHandler({
      available: () => true,
      active: () => true,
      flush: async () => {
        expect(await bridge.request('view.flush', {})).toEqual({ ready: true })
      },
      open: async () => {
        throw new Error('File picker belongs to the native host')
      },
      save: async () => {
        menuSave = await session.save()
      },
      bind: async (scope) => {
        const uri = 'file:///workspace/browser.docx'
        const view = (await bridge.request('view.selection', {})) as {
          stateId: string
          revision: number
          from: number
          to: number
        }
        const selection = scope === 'selection' ? session.selection(view) : undefined
        capturedScope = scope
        capturedBinding = bindings.bind(uri, session.sessionId, selection)
        return {
          kind: 'docx',
          scope,
          documentUri: uri,
          bindingId: capturedBinding,
          revision: session.descriptor().revision,
        }
      },
      chat: async (query, partial) => {
        chats.push({ query, partial, scope: capturedScope, bindingId: capturedBinding })
      },
    }),
  )
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
    await page.getByTitle('Save', { exact: true }).click()
    await page.waitForFunction(
      () => !document.querySelector<HTMLButtonElement>('button[title="Save"]')?.disabled,
    )
    expect(menuSave).toBeDefined()
    expect((await DocxSession.open(menuSave!)).read().blocks[0]!.text).toBe('Hello world')
    await page.getByTitle('VS Code Chat', { exact: true }).click()
    await page.waitForFunction(
      () => !document.querySelector<HTMLButtonElement>('button[title="VS Code Chat"]')?.disabled,
    )
    expect(chats[0]).toMatchObject({ partial: true, scope: 'document' })
    expect(bindings.resolve(chats[0]!.bindingId).selection).toBeUndefined()
    await page.keyboard.press('Control+Home')
    await page.keyboard.press('Control+Shift+ArrowRight')
    await page.getByTitle('Ask about selection', { exact: true }).click()
    await page.waitForFunction(
      () =>
        !document.querySelector<HTMLButtonElement>('button[title="Ask about selection"]')?.disabled,
    )
    expect(chats[1]).toMatchObject({ partial: true, scope: 'selection' })
    expect(bindings.resolve(chats[1]!.bindingId).selection?.blocks.length).toBeGreaterThan(0)
    expect(chats[1]!.query).not.toContain('Hello world')
    await editor.evaluate((element) =>
      (
        element as unknown as { editor: { commands: { focus(at: string): void } } }
      ).editor.commands.focus('end'),
    )
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
    await page.getByRole('button', { name: 'Bold (Ctrl+B)', exact: true }).click()
    expect(await bridge.request('view.flush', {})).toEqual({ ready: true })
    expect(
      session
        .content()
        .content![0]!.content!.some((node) => node.marks?.some((mark) => mark.type === 'bold')),
    ).toBe(true)
    const beforeDialog = session.content()
    await page.keyboard.press('Control+KeyD')
    let fontDialog = page.getByRole('dialog', { name: 'Font', exact: true })
    await fontDialog.getByRole('tab', { name: 'Advanced', exact: true }).click()
    await fontDialog.getByLabel('Scale (%)', { exact: true }).fill('80')
    await fontDialog.getByRole('button', { name: 'Cancel', exact: true }).click()
    expect(await bridge.request('view.flush', {})).toEqual({ ready: true })
    expect(session.content()).toEqual(beforeDialog)
    await page.keyboard.press('Control+KeyD')
    fontDialog = page.getByRole('dialog', { name: 'Font', exact: true })
    await fontDialog.getByRole('tab', { name: 'Advanced', exact: true }).click()
    await fontDialog.getByLabel('Scale (%)', { exact: true }).fill('80')
    await fontDialog.getByRole('button', { name: 'Spacing', exact: true }).click()
    await page.getByRole('option', { name: 'Expanded', exact: true }).click()
    await fontDialog.getByLabel('By (pt)', { exact: true }).first().fill('1.5')
    await page.screenshot({
      path: resolve('../../test-results/desktop-menus/docx-font-advanced.png'),
    })
    await fontDialog.getByRole('button', { name: 'OK', exact: true }).click()
    expect(await bridge.request('view.flush', {})).toEqual({ ready: true })
    const hasDialogFormatting = (content: ReturnType<DocxSession['content']>) =>
      content.content![0]!.content!.some((node) =>
        node.marks?.some(
          (mark) =>
            mark.type === 'docTextStyle' &&
            mark.attrs?.charScalePct === 80 &&
            mark.attrs?.charSpacingTwips === 30,
        ),
      )
    expect(hasDialogFormatting(session.content())).toBe(true)
    const reopenedDialog = await DocxSession.open(await session.save())
    expect(hasDialogFormatting(reopenedDialog.content())).toBe(true)
    // The entire dialog is one native undo entry, including derived scale attrs.
    session.history('undo')
    expect(session.content()).toEqual(beforeDialog)
    session.history('redo')
    await bridge.request('view.resync', projection())
    const beforeFontText = await editor.textContent()
    const sizeInput = page.getByLabel('Font Size', { exact: true }).first()
    await sizeInput.focus()
    await page.keyboard.press('Control+A')
    await page.keyboard.type('14')
    expect(await sizeInput.inputValue()).toBe('14')
    await page.keyboard.press('Enter')
    expect(await bridge.request('view.flush', {})).toEqual({ ready: true })
    const fontInput = page.getByLabel('Font', { exact: true }).first()
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
              mark.attrs?.fontAscii === 'Times New Roman',
          ),
        ),
    ).toBe(true)
    const beforeParagraph = session.content()
    await page.keyboard.press('Control+Alt+KeyM')
    const paragraphDialog = page.getByRole('dialog', { name: 'Paragraph', exact: true })
    await paragraphDialog.getByLabel(/^Space before/).fill('12')
    await paragraphDialog.getByRole('button', { name: 'Line spacing', exact: true }).click()
    await page.getByRole('option', { name: '1.5 lines', exact: true }).click()
    await paragraphDialog.getByRole('tab', { name: 'Line and Page Breaks' }).click()
    await paragraphDialog.getByLabel('Keep with next', { exact: true }).check()
    await paragraphDialog.getByRole('button', { name: 'OK', exact: true }).click()
    expect(await bridge.request('view.flush', {})).toEqual({ ready: true })
    expect(session.content().content![0]!.attrs).toMatchObject({
      spaceBefore: 240,
      lineSpacing: 1.5,
      keepNext: true,
    })
    expect(
      (await DocxSession.open(await session.save())).content().content![0]!.attrs,
    ).toMatchObject({
      spaceBefore: 240,
      lineSpacing: 1.5,
      keepNext: true,
    })
    session.history('undo')
    expect(session.content()).toEqual(beforeParagraph)
    session.history('redo')
    await bridge.request('view.resync', projection())
    await page.locator('.style-gallery').getByText('Heading 1', { exact: true }).click()
    expect(await bridge.request('view.flush', {})).toEqual({ ready: true })
    expect(session.content().content![0]!.type).toBe('docHeading')
    await page.locator('.style-gallery').getByText('Normal', { exact: true }).click()
    expect(await bridge.request('view.flush', {})).toEqual({ ready: true })
    await page.locator('.ribbon-tab', { hasText: /^Insert$/ }).click()
    await page.locator('[data-tip="Insert a table"]').click()
    await page.locator('.table-picker-grid .table-cell').nth(11).click()
    expect(await bridge.request('view.flush', {})).toEqual({ ready: true })
    await page.locator('.tiptap td p').first().click()
    await page.locator('.ribbon-tab', { hasText: /^Table Layout$/ }).click()
    await page.getByRole('button', { name: 'Insert Below', exact: true }).click()
    expect(await bridge.request('view.flush', {})).toEqual({ ready: true })
    await page.getByRole('button', { name: 'Insert Right', exact: true }).click()
    expect(await bridge.request('view.flush', {})).toEqual({ ready: true })
    const table = session.content().content!.find((node) => node.type === 'docTable')!
    expect(table.content).toHaveLength(3)
    expect(table.content![0]!.content).toHaveLength(3)
    // Insert → Pictures reads a local file in the view and saves a new media part.
    await editor.evaluate((element) =>
      (
        element as unknown as { editor: { commands: { focus(at: string): void } } }
      ).editor.commands.focus('end'),
    )
    await page.locator('.ribbon-tab', { hasText: /^Insert$/ }).click()
    const chooser = page.waitForEvent('filechooser')
    await page.locator('[data-tip="Insert a picture"]').first().click()
    await (
      await chooser
    ).setFiles({
      name: 'dot.png',
      mimeType: 'image/png',
      buffer: Buffer.from(
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==',
        'base64',
      ),
    })
    await page.waitForFunction(() => document.querySelector('.tiptap [data-office-protected]'))
    expect(await bridge.request('view.flush', {})).toEqual({ ready: true })
    expect(session.content().content!.some((node) => node.attrs?.blockType === 'image')).toBe(true)
    // The editor context menu and Ctrl+F open the shared desktop dialog and panel.
    await page.locator('.tiptap p').first().click({ button: 'right' })
    await page.locator('.ctx-menu .ctx-item', { hasText: /^Font/ }).first().click()
    await page.getByRole('dialog', { name: 'Font', exact: true }).waitFor()
    await page.keyboard.press('Escape')
    await page.locator('.tiptap p').first().click()
    await page.keyboard.press('Control+KeyF')
    const findPanel = page.locator('.find-panel')
    await findPanel.waitFor()
    await findPanel.locator('input').first().press('Escape')
    await findPanel.waitFor({ state: 'hidden' })
    // Home → Bullets allocates a list definition in the view; the host takes it with the steps.
    await page.locator('.tiptap p').first().click()
    await page.locator('.ribbon-tab', { hasText: /^Home$/ }).click()
    await page.getByRole('button', { name: 'Bullets', exact: true }).first().click()
    expect(await bridge.request('view.flush', {})).toEqual({ ready: true })
    const nodes: { type: string; attrs?: Record<string, unknown>; content?: unknown[] }[] = []
    const collect = (node: { type: string; content?: unknown[] }) => {
      nodes.push(node)
      for (const child of node.content ?? []) collect(child as { type: string })
    }
    collect(session.content())
    const listItem = nodes.find((node) => node.type === 'docListItem')!
    expect(listItem.attrs!.kind).toBe('bullet')
    expect(new Map(session.presentation().numbering).has(String(listItem.attrs!.numId))).toBe(true)
    expect(await page.locator('.office-docx-error').count()).toBe(0)
    // Layout → Margins → Narrow: a settings entry the page geometry and the save follow.
    await page
      .locator('.ribbon-tab', { hasText: /^Layout$/ })
      .first()
      .click()
    await page.locator('[data-tip="Margins"]').first().click()
    await page.locator('[data-rb-panel] button', { hasText: 'Narrow' }).first().click()
    await expect.poll(() => session.presentation().settings!.section!.marginLeft).toBe(720)
    expect(
      await page
        .locator('.office-docx-page')
        .evaluate((element) => getComputedStyle(element).paddingLeft),
    ).toBe('48px')
    expect(await bridge.request('view.flush', {})).toEqual({ ready: true })
    session.history('undo')
    expect(session.presentation().settings!.section!.marginLeft).toBe(1440)
    session.history('redo')
    await bridge.request('view.resync', projection())
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
    await page.evaluate(() => {
      document.body.classList.add('vscode-dark')
      for (const [name, value] of Object.entries({
        '--vscode-foreground': 'rgb(220, 220, 220)',
        '--vscode-descriptionForeground': 'rgb(170, 170, 170)',
        '--vscode-panel-border': 'rgb(68, 68, 68)',
        '--vscode-input-background': 'rgb(48, 48, 48)',
        '--vscode-sideBar-background': 'rgb(37, 37, 38)',
        '--vscode-button-background': 'rgb(14, 99, 156)',
        '--vscode-button-foreground': 'rgb(255, 255, 255)',
        '--vscode-button-hoverBackground': 'rgb(17, 119, 187)',
        '--vscode-toolbar-hoverBackground': 'rgb(51, 51, 51)',
      }))
        document.documentElement.style.setProperty(name, value)
    })
    await page.keyboard.press('Control+KeyD')
    expect(
      await page
        .getByRole('dialog', { name: 'Font', exact: true })
        .evaluate((element) => getComputedStyle(element).backgroundColor),
    ).toBe('rgb(20, 20, 20)')
    expect(
      await page
        .locator('.font-preview')
        .evaluate((element) => getComputedStyle(element).backgroundColor),
    ).toBe('rgb(255, 255, 255)')
    await page.screenshot({
      path: resolve('../../test-results/desktop-menus/docx-font-vscode-dark.png'),
    })
    await page.keyboard.press('Escape')
    expect(await page.getByRole('dialog').count()).toBe(0)
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
