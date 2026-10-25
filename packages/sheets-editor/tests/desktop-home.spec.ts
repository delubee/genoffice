import { test, expect } from '@playwright/test'
import { createServer, type ViteDevServer } from 'vite'
import { fileURLToPath } from 'node:url'
let server: ViteDevServer
let base: string

test.beforeAll(async () => {
  server = await createServer({
    configFile: false,
    root: fileURLToPath(new URL('../../..', import.meta.url)),
    server: { host: '127.0.0.1', port: 0 },
  })
  await server.listen()
  base = server.resolvedUrls!.local[0]!
})
test.afterAll(async () => {
  await server?.close()
})

test('desktop Home controls and Format Cells commit one validated dialog batch', async ({
  page,
}) => {
  const errors: string[] = []
  page.on('pageerror', (error) => errors.push(String(error)))
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(true)
    ;(window as any).toolbar.select('B2')
  })
  const home = page.locator('.workbook-desktop-header')
  await expect(home.locator('.ribbon-group')).toHaveCount(8)
  await home.locator('[data-tip="Bold"]').click()
  await expect
    .poll(() =>
      page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells.B2.style?.bold),
    )
    .toBe(true)
  await expect(home.locator('[data-tip="Bold"]')).toHaveClass(/is-active/)
  await home.getByRole('textbox', { name: 'Font family', exact: true }).fill('Georgia')
  await page.keyboard.press('Enter')
  await expect
    .poll(() =>
      page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells.B2.style?.fontFamily),
    )
    .toBe('Georgia')
  const beforeDialog = await page.evaluate(() => (window as any).toolbar.revision())
  await home.locator('[data-tip="Format Cells"]').click()
  const dialog = page.getByRole('dialog', { name: 'Format Cells', exact: true })
  await expect(dialog).toBeVisible()
  await dialog.getByRole('button', { name: 'Font', exact: true }).click()
  await dialog.getByRole('button', { name: 'Italic', exact: true }).click()
  await page.getByRole('option', { name: 'On', exact: true }).click()
  await dialog.getByRole('button', { name: 'Alignment', exact: true }).click()
  await dialog.getByRole('button', { name: 'Wrap text', exact: true }).click()
  await page.getByRole('option', { name: 'On', exact: true }).click()
  await dialog.getByRole('button', { name: 'OK', exact: true }).click()
  await expect(dialog).toHaveCount(0)
  expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(beforeDialog + 1)
  await expect
    .poll(() =>
      page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells.B2.style?.wrapText),
    )
    .toBe(true)
  await home.locator('[data-tip="Format Cells"]').click()
  await expect(dialog).toBeVisible()
  await dialog.getByRole('button', { name: 'Cancel', exact: true }).click()
  await home.locator('.ribbon').evaluate((element) => {
    element.scrollLeft = 0
  })
  await page.screenshot({
    path: fileURLToPath(
      new URL('../../../test-results/desktop-menus/sheets-home.png', import.meta.url),
    ),
  })
  expect(errors).toEqual([])
  await expect(page.getByRole('alert')).toHaveCount(0)
  await home.locator('[data-tip="Format Cells"]').click()
  await dialog.getByRole('button', { name: 'Alignment', exact: true }).click()
  await dialog.getByRole('button', { name: 'Wrap text', exact: true }).click()
  await page.getByRole('option', { name: 'Off', exact: true }).click()
  await page.evaluate(() => (window as any).toolbar.editNote('Concurrent update'))
  const changedRevision = await page.evaluate(() => (window as any).toolbar.revision())
  await dialog.getByRole('button', { name: 'OK', exact: true }).click()
  await expect(page.getByRole('alert')).toContainText('document changed')
  await expect(dialog).toBeVisible()
  expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(changedRevision)
  expect(
    await page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells.B2.style.wrapText),
  ).toBe(true)
  await dialog.getByRole('button', { name: 'Cancel', exact: true }).click()
})
