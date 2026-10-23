import { test, expect } from '@playwright/test'
import { createServer, type ViteDevServer } from 'vite'
import { fileURLToPath } from 'node:url'
let server: ViteDevServer
let base: string
test.beforeAll(async () => {
  server = await createServer({
    configFile: false,
    root: fileURLToPath(new URL('../../..', import.meta.url)),
    esbuild: { jsx: 'automatic' },
    server: { host: '127.0.0.1', port: 0 },
  })
  await server.listen()
  base = server.resolvedUrls!.local[0]!
})
test.afterAll(async () => {
  await server?.close()
})
test('manual commands become semantic intents and unsupported mutations are blocked', async ({
  page,
}) => {
  await page.goto(base)
  const result = await page.evaluate(async () => {
    const { manualValueIntents } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/src/manual-intents.ts')
    )
    return manualValueIntents('logical-sheet', {
      range: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 },
      value: [[{ v: '00123' }, { f: '=A2*2' }]],
    })
  })
  expect(result).toEqual([
    { op: 'set_cell', sheetId: 'logical-sheet', address: 'A1', value: '00123' },
    { op: 'set_formula', sheetId: 'logical-sheet', address: 'B1', formula: '=A2*2' },
  ])
})

test('visible Univer values wait for host commit and IME blocks flush', async ({ page }) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualRuntime } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-runtime.ts')
    )
    ;(window as unknown as Record<string, any>).manual = await mountManualRuntime()
  })
  try {
    await expect(page.evaluate(() => (window as any).manual.mutateDirectly())).rejects.toThrow(
      'unsupported_manual_command',
    )
    expect(
      await page.evaluate(() => (window as unknown as Record<string, any>).manual.value()),
    ).toBe(123)
    await page.evaluate(async () => {
      ;(window as unknown as Record<string, any>).manual.type(900)
      await Promise.resolve()
    })
    expect(
      await page.evaluate(
        () => (window as unknown as Record<string, any>).manual.requested().operations[0],
      ),
    ).toMatchObject({ op: 'set_cell', sheetId: 'manual-sheet', address: 'A1', value: 900 })
    expect(
      await page.evaluate(() => (window as unknown as Record<string, any>).manual.value()),
    ).toBe(123)
    await page.evaluate(async () => {
      await (window as unknown as Record<string, any>).manual.commit()
    })
    expect(
      await page.evaluate(() => (window as unknown as Record<string, any>).manual.value()),
    ).toBe(900)
    expect(
      await page.evaluate(async () => {
        const manual = (window as unknown as Record<string, any>).manual
        manual.composing(true)
        return manual.flush()
      }),
    ).toBe(false)
    expect(
      await page.evaluate(async () => {
        const manual = (window as unknown as Record<string, any>).manual
        manual.composing(false)
        return manual.flush()
      }),
    ).toBe(true)
  } finally {
    await page.evaluate(() => (window as unknown as Record<string, any>).manual.dispose())
  }
})

test('native formatting waits for host commit and preserves document colors and borders', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualRuntime } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-runtime.ts')
    )
    ;(window as any).manual = await mountManualRuntime()
  })
  try {
    await page.evaluate(() => (window as any).manual.format())
    expect(
      await page.evaluate(() => (window as any).manual.requested().operations[0]),
    ).toMatchObject({ op: 'format_range', sheetId: 'manual-sheet', format: { bold: true } })
    expect(await page.evaluate(() => (window as any).manual.style()?.bl ?? 0)).toBe(0)
    await page.evaluate(() => (window as any).manual.commitFormat())
    expect(await page.evaluate(() => (window as any).manual.style().bl)).toBe(1)
    const style = await page.evaluate(async () => {
      const { gridCell } = await import(
        /* @vite-ignore */ String('/packages/sheets-editor/src/grid-cell.ts')
      )
      return gridCell({
        kind: 'string',
        rawValue: 'literal',
        style: {
          fontColor: '#123456',
          fillColor: '#FEDCBA',
          wrapText: true,
          horizontalAlignment: 'right',
          verticalAlignment: 'center',
          borderTop: { style: 'double', color: '#654321' },
        },
      }).s
    })
    expect(style).toMatchObject({
      cl: { rgb: '#123456' },
      bg: { rgb: '#FEDCBA' },
      tb: 3,
      ht: 3,
      vt: 2,
      bd: { t: { s: 7, cl: { rgb: '#654321' } } },
    })
  } finally {
    await page.evaluate(() => (window as any).manual.dispose())
  }
})

test('toolbar creates, edits and deletes a chart whose cache follows current cells', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar()
    ;(window as any).toolbar.select('A1:B3')
  })
  try {
    await page.getByRole('button', { name: 'Add chart', exact: true }).click()
    await expect(page.locator('.workbook-charts svg')).toHaveCount(1)
    await page.getByLabel('Chart title', { exact: true }).fill('Current sales')
    const form = page.locator('.workbook-charts form')
    await form.getByRole('combobox').selectOption('line')
    await form.getByRole('button', { name: 'Apply', exact: true }).click()
    await expect(page.getByRole('img', { name: 'Current sales', exact: true })).toBeVisible()
    await page.evaluate(() => (window as any).toolbar.edit())
    expect(
      await page.evaluate(
        () =>
          (window as any).toolbar.snapshot().sheets[0].features.charts[0].chart.series[0].values,
      ),
    ).toEqual([200, -30])
    const beforeTheme = await page.evaluate(() => (window as any).toolbar.snapshot())
    await page.evaluate(() => document.body.classList.add('vscode-dark'))
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.darkTheme())).toBe(true)
    await page.evaluate(() => {
      document.body.classList.remove('vscode-dark')
      document.body.classList.add('vscode-light')
    })
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.darkTheme())).toBe(false)
    expect(await page.evaluate(() => (window as any).toolbar.snapshot())).toEqual(beforeTheme)
    await page.getByRole('button', { name: 'Delete chart', exact: true }).click()
    await expect(page.locator('.workbook-charts')).toHaveCount(0)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('toolbar copy, fill, formats and sheet/row structure share the canonical state', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar()
    ;(window as any).toolbar.select('C2')
  })
  const snapshot = () => page.evaluate(() => (window as any).toolbar.snapshot())
  try {
    await page.getByLabel('Source range (current sheet)', { exact: true }).fill('A2')
    await page
      .getByRole('button', { name: 'Copy values and formulas to selection', exact: true })
      .click()
    await expect.poll(async () => (await snapshot()).sheets[0].cells.C2?.rawValue).toBe('North')
    await page.evaluate(() => (window as any).toolbar.select('D2:D10'))
    await page.getByRole('button', { name: 'Fill selection with source', exact: true }).click()
    await expect.poll(async () => (await snapshot()).sheets[0].cells.D10?.rawValue).toBe('North')
    await page.locator('summary').filter({ hasText: 'More formatting' }).click()
    await page.getByLabel('More formatting', { exact: true }).selectOption('fillColor')
    await page.getByLabel('Fill color', { exact: true }).fill('#123456')
    await page.locator('details').getByRole('button', { name: 'Apply', exact: true }).click()
    await expect
      .poll(async () => (await snapshot()).sheets[0].cells.D10?.style?.fillColor)
      .toBe('#123456')
    await page.getByLabel('Sheet name', { exact: true }).fill('Extra')
    await page.getByRole('button', { name: 'Add sheet', exact: true }).click()
    await expect.poll(async () => (await snapshot()).sheets.length).toBe(2)
    const logicalId = (await snapshot()).sheets[1].id
    await page.evaluate(() => (window as any).toolbar.selectSheet('Extra'))
    await page.getByLabel('Sheet name', { exact: true }).fill('Renamed')
    await page.getByRole('button', { name: 'Rename sheet', exact: true }).click()
    await expect.poll(async () => (await snapshot()).sheets[1].name).toBe('Renamed')
    expect((await snapshot()).sheets[1].id).toBe(logicalId)
    await page.getByRole('button', { name: 'Delete sheet', exact: true }).click()
    await expect.poll(async () => (await snapshot()).sheets.length).toBe(1)
    await page.evaluate(() => (window as any).toolbar.select('A2'))
    await page.getByRole('button', { name: 'Insert row', exact: true }).click()
    await expect.poll(async () => (await snapshot()).sheets[0].cells.A3?.rawValue).toBe('North')
    await page.getByRole('button', { name: 'Delete row', exact: true }).click()
    await expect.poll(async () => (await snapshot()).sheets[0].cells.A2?.rawValue).toBe('North')
    await page.evaluate(() => (window as any).toolbar.select('B2'))
    await page.getByRole('button', { name: 'Insert column', exact: true }).click()
    await expect.poll(async () => (await snapshot()).sheets[0].cells.C2?.rawValue).toBe(120)
    await page.getByRole('button', { name: 'Delete column', exact: true }).click()
    await expect.poll(async () => (await snapshot()).sheets[0].cells.B2?.rawValue).toBe(120)
    await expect(page.getByRole('alert')).toHaveCount(0)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('basic formulas compute visibly after committed edits and update with their input', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar()
    ;(window as any).toolbar.formula()
  })
  try {
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.value('C2'))).toBe(240)
    expect(
      await page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells.C2.formula),
    ).toBe('=B2*2')
    await page.evaluate(() => (window as any).toolbar.edit())
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.value('C2'))).toBe(400)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})
