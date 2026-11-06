import { test, expect, type Page } from '@playwright/test'
import { createServer, type ViteDevServer } from 'vite'
import { fileURLToPath } from 'node:url'
async function pickOutlineCommand(page: Page, name: string) {
  const menu = /^(Ungroup|Show Detail|Clear Outline)/.test(name) ? 'Ungroup' : 'Group'
  await page.getByRole('button', { name: menu, exact: true }).click()
  await page.getByRole('option', { name, exact: true }).click()
}
test.setTimeout(60_000)
let server: ViteDevServer
let base: string
test.beforeAll(async () => {
  server = await createServer({
    configFile: false,
    cacheDir: fileURLToPath(new URL('../../../node_modules/.vite-manual-toolbar', import.meta.url)),
    root: fileURLToPath(new URL('../../..', import.meta.url)),
    esbuild: { jsx: 'automatic' },
    optimizeDeps: {
      include: [
        'react',
        'react-dom/client',
        'react-dom',
        '@univerjs/find-replace',
        'rxjs',
        'react/jsx-runtime',
        'react/jsx-dev-runtime',
        '@univerjs/core',
        '@univerjs/sheets',
        '@univerjs/engine-render',
        '@univerjs/engine-formula',
        '@univerjs/preset-sheets-core',
        '@univerjs/sheets-numfmt',
        '@univerjs/sheets-formula',
        '@univerjs/docs-ui',
        '@univerjs/sheets-ui',
      ],
    },
    server: { host: '127.0.0.1', port: 0, watch: { ignored: ['**/dist/**'] } },
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
    await page.getByRole('tab', { name: 'Insert', exact: true }).click()
    await page.locator('#ribbon-panel-insert details > summary').click()
    await page.getByRole('button', { name: 'Add chart', exact: true }).click()
    await expect(page.locator('.xlsx-print-visual svg')).toHaveCount(1)
    await page.locator('.chart-edit-button').click()
    const panel = page.locator('.chart-editor')
    await panel.locator('.chart-editor-title').fill('Current sales')
    await panel.locator('.chart-editor-type').click()
    await page.locator('button[role=option][data-value=line]').click()
    await panel.getByRole('button', { name: 'Apply', exact: true }).click()
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
    await page.locator('.shape-editable').focus()
    await page.keyboard.press('Delete')
    await expect(page.locator('.xlsx-print-visual')).toHaveCount(0)
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
    await page.locator('#ribbon-panel-home .workbook-additional-controls > summary').click()
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
    await page.locator('input[aria-label="Fill color"]').fill('#123456')
    await page
      .locator('details:has(> summary:text-is("More formatting"))')
      .getByRole('button', { name: 'Apply', exact: true })
      .click()
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
    await page
      .locator('.workbook-additional-controls')
      .getByRole('button', { name: 'Insert row', exact: true })
      .click()
    await expect.poll(async () => (await snapshot()).sheets[0].cells.A3?.rawValue).toBe('North')
    await page
      .locator('.workbook-additional-controls')
      .getByRole('button', { name: 'Delete row', exact: true })
      .click()
    await expect.poll(async () => (await snapshot()).sheets[0].cells.A2?.rawValue).toBe('North')
    await page.evaluate(() => (window as any).toolbar.select('B2'))
    await page
      .locator('.workbook-additional-controls')
      .getByRole('button', { name: 'Insert column', exact: true })
      .click()
    await expect.poll(async () => (await snapshot()).sheets[0].cells.C2?.rawValue).toBe(120)
    await page
      .locator('.workbook-additional-controls')
      .getByRole('button', { name: 'Delete column', exact: true })
      .click()
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

test('Excel style stays light and switching chrome preserves the current workbook', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    document.body.classList.add('vscode-dark')
    document.documentElement.style.setProperty('--vscode-editor-background', '#191919')
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(true)
  })
  try {
    const style = page.getByRole('combobox', { name: 'Interface style', exact: true })
    await expect(style).toHaveValue('excel')
    await page.waitForTimeout(350)
    await page.screenshot({ path: test.info().outputPath('excel-style.png') })
    expect(await page.evaluate(() => (window as any).toolbar.darkTheme())).toBe(false)
    await expect(page.locator('body')).toHaveCSS('background-color', 'rgb(255, 255, 255)')
    await page.evaluate(() => (window as any).toolbar.edit())
    const before = await page.evaluate(() => (window as any).toolbar.snapshot())
    await style.selectOption('vscode')
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.darkTheme())).toBe(true)
    await expect(page.locator('body')).toHaveCSS('background-color', 'rgb(25, 25, 25)')
    await page.waitForTimeout(350)
    await page.screenshot({ path: test.info().outputPath('vscode-style.png') })
    expect(await page.evaluate(() => (window as any).toolbar.snapshot())).toEqual(before)
    await page.evaluate(() => document.body.classList.replace('vscode-dark', 'vscode-light'))
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.darkTheme())).toBe(false)
    await page.evaluate(() => document.body.classList.replace('vscode-light', 'vscode-dark'))
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.darkTheme())).toBe(true)
    await style.selectOption('excel')
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.darkTheme())).toBe(false)
    expect(await page.evaluate(() => (window as any).toolbar.snapshot())).toEqual(before)
    await expect(page.getByRole('alert')).toHaveCount(0)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('horizontal and vertical merges survive manual editing and projection replacement', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualRuntime } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-runtime.ts')
    )
    ;(window as any).manual = await mountManualRuntime(true)
  })
  try {
    const merges = await page.evaluate(() => (window as any).manual.merges())
    expect(merges).toHaveLength(2)
    expect(merges[0]).toMatchObject({ startRow: 0, endRow: 0, startColumn: 0, endColumn: 2 })
    expect(merges[1]).toMatchObject({ startRow: 2, endRow: 4, startColumn: 0, endColumn: 0 })
    await page.evaluate(() => (window as any).manual.editMerged())
    expect(await page.evaluate(() => (window as any).manual.requested().operations)).toEqual([
      { op: 'set_cell', sheetId: 'manual-sheet', address: 'A1', value: 900 },
    ])
    await page.evaluate(() => (window as any).manual.commit())
    expect(await page.evaluate(() => (window as any).manual.value())).toBe(900)
    expect(await page.evaluate(() => (window as any).manual.merges())).toEqual(merges)
  } finally {
    await page.evaluate(() => (window as any).manual.dispose())
  }
})

test('typing into a merged selection forwards only the anchor', async ({ page }) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(false, true)
  })
  try {
    await expect
      .poll(() =>
        page.evaluate(() =>
          [...document.querySelectorAll('canvas')].some(
            (e) => e.getBoundingClientRect().height > 100,
          ),
        ),
      )
      .toBe(true)
    expect(await page.evaluate(() => (window as any).toolbar.beginMergedEdit())).toBe(true)
    await page.keyboard.press('ControlOrMeta+A')
    await page.keyboard.type('Typed heading')
    await page.keyboard.press('Enter')
    await expect
      .poll(() =>
        page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells.A1.rawValue),
      )
      .toBe('Typed heading')
    expect(
      await page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells.B1.rawValue),
    ).toBe('Sales')
    await page.screenshot({ path: test.info().outputPath('merged-heading.png') })
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('numeric sizes and native header drags commit persistent dimensions', async ({ page }) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar()
    ;(window as any).toolbar.select('B2')
  })
  try {
    await page.locator('#ribbon-panel-home .workbook-additional-controls > summary').click()
    await page.getByRole('spinbutton', { name: 'Column width (px)' }).fill('140')
    await page.getByRole('button', { name: 'Set column width', exact: true }).click()
    await expect
      .poll(() => page.evaluate(() => (window as any).toolbar.sizes().columns[1].pixels))
      .toBe(140)
    await page.getByRole('spinbutton', { name: 'Row height (pt)' }).fill('30')
    await page.getByRole('button', { name: 'Set row height', exact: true }).click()
    await expect
      .poll(() => page.evaluate(() => (window as any).toolbar.sizes().rows[1].pixels))
      .toBe(40)
    await expect
      .poll(() =>
        page.evaluate(() =>
          [...document.querySelectorAll('canvas')].some(
            (e) => e.getBoundingClientRect().height > 100,
          ),
        ),
      )
      .toBe(true)
    // The column header's right edge comes from the current rendered geometry.
    const position = await page.evaluate(() => {
      const canvas = [...document.querySelectorAll('canvas')].find(
        (e) => e.getBoundingClientRect().height > 100,
      )!
      const rect = canvas.getBoundingClientRect()
      const sizes = (window as any).toolbar.sizes()
      return { x: rect.x + 46 + sizes.columns[0].pixels + sizes.columns[1].pixels, y: rect.y + 10 }
    })
    await page.mouse.move(position.x, position.y)
    await page.mouse.down()
    await page.mouse.move(position.x + 40, position.y, { steps: 8 })
    await page.mouse.up()
    await expect
      .poll(() => page.evaluate(() => (window as any).toolbar.sizes().columns[1].pixels))
      .toBeGreaterThan(175)
    const draggedWidth = await page.evaluate(
      () => (window as any).toolbar.sizes().columns[1].pixels,
    )
    expect(draggedWidth).toBeLessThan(183)
    expect(
      await page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].features.colWidths.B),
    ).toBe(draggedWidth)
    await page.evaluate(() => (window as any).toolbar.flush())
    const rowPosition = await page.evaluate(() => {
      const canvas = [...document.querySelectorAll('canvas')].find(
        (e) => e.getBoundingClientRect().height > 100,
      )!
      const rect = canvas.getBoundingClientRect()
      return { x: rect.x + 23, y: rect.y + (window as any).toolbar.cellRect('B2').bottom }
    })
    await page.mouse.move(rowPosition.x, rowPosition.y)
    await page.mouse.down()
    await page.mouse.move(rowPosition.x, rowPosition.y + 16, { steps: 8 })
    await page.mouse.up()
    await expect
      .poll(() => page.evaluate(() => (window as any).toolbar.sizes().rows[1].pixels))
      .toBeGreaterThan(50)
    const draggedHeight = await page.evaluate(() => (window as any).toolbar.sizes().rows[1].pixels)
    expect(draggedHeight).toBeLessThan(60)
    expect(
      await page.evaluate(
        () => ((window as any).toolbar.snapshot().sheets[0].features.rowHeights['2'] * 4) / 3,
      ),
    ).toBeCloseTo(draggedHeight, 2)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('ribbon merges and unmerges the projected grid and refuses to hide existing data', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(true)
    ;(window as any).toolbar.select('C1:D2')
  })
  try {
    await page.getByRole('button', { name: 'Merge cells', exact: true }).click()
    await page.getByRole('option', { name: /^Merge cells$/i }).click()
    await expect
      .poll(() => page.evaluate(() => (window as any).toolbar.merges()))
      .toEqual([{ startRow: 0, endRow: 1, startColumn: 2, endColumn: 3 }])
    await page.evaluate(() => (window as any).toolbar.select('D2'))
    await page.getByRole('button', { name: 'Merge cells', exact: true }).click()
    await page.getByRole('option', { name: /^Unmerge cells$/i }).click()
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.merges())).toEqual([])
    await page.evaluate(() => (window as any).toolbar.select('A1:B1'))
    await page.getByRole('button', { name: 'Merge cells', exact: true }).click()
    await page.getByRole('option', { name: /^Merge cells$/i }).click()
    await expect(page.getByRole('alert')).toContainText('Other cells contain data')
    expect(await page.evaluate(() => (window as any).toolbar.value('B1'))).toBe('Sales')
    expect(await page.evaluate(() => (window as any).toolbar.merges())).toEqual([])
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('seven ribbon tabs execute page, formula, data, review and view operations', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(true)
    ;(window as any).toolbar.select('A1:B3')
  })
  const snapshot = () => page.evaluate(() => (window as any).toolbar.snapshot())
  const tab = (name: string) =>
    page
      .getByRole('tablist', { name: 'Ribbon', exact: true })
      .getByRole('tab', { name, exact: true })
      .click()
  try {
    for (const name of ['Home', 'Insert', 'Page Layout', 'Formulas', 'Data', 'Review', 'View'])
      await expect(
        page
          .getByRole('tablist', { name: 'Ribbon', exact: true })
          .getByRole('tab', { name, exact: true }),
      ).toBeVisible()
    await page.bringToFront()
    const homeTab = page
      .getByRole('tablist', { name: 'Ribbon', exact: true })
      .getByRole('tab', { name: 'Home', exact: true })
    await homeTab.focus()
    await expect(homeTab).toBeFocused()
    const beforeTabKey = await page.evaluate(() => (window as any).toolbar.selection())
    await page.keyboard.press('ArrowRight')
    await expect(
      page
        .getByRole('tablist', { name: 'Ribbon', exact: true })
        .getByRole('tab', { name: 'Insert', exact: true }),
    )
      .toBeFocused()
      .catch(async (error) => {
        console.log(
          'Tab focus diagnostic',
          await page.evaluate(() => ({
            html: document.activeElement?.outerHTML.slice(0, 500),
            active: document.hasFocus(),
            selection: (window as any).toolbar.selection(),
          })),
        )
        throw error
      })
    expect(await page.evaluate(() => (window as any).toolbar.selection())).toEqual(beforeTabKey)
    await tab('Page Layout')
    await page
      .locator('.workbook-desktop-header')
      .getByRole('button', { name: 'Orientation', exact: true })
      .click()
    await page.getByRole('option', { name: 'Landscape', exact: true }).click()
    await expect
      .poll(async () => (await snapshot()).sheets[0].features.pageSetupEdits.orientation)
      .toBe('landscape')
    await page
      .locator('.workbook-desktop-header')
      .getByRole('button', { name: 'Size', exact: true })
      .click()
    await page.getByRole('option', { name: 'A3', exact: true }).click()
    await expect
      .poll(async () => (await snapshot()).sheets[0].features.pageSetupEdits.paperSize)
      .toBe(8)
    await page
      .locator('.workbook-desktop-header')
      .getByRole('button', { name: 'Print Area', exact: true })
      .click()
    await page.getByRole('option', { name: 'Set Print Area', exact: true }).click()
    await expect
      .poll(async () => (await snapshot()).sheets[0].features.pageSetupEdits.printArea)
      .toBe('A1:B3')
    await tab('Formulas')
    await page.evaluate(() => (window as any).toolbar.select('B2:B3'))
    await page.getByRole('button', { name: 'AutoSum', exact: true }).click()
    await page.getByRole('option', { name: 'Sum', exact: true }).click()
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.value('B4'))).toBe(90)
    await tab('Data')
    await page.evaluate(() => (window as any).toolbar.select('A1:B3'))
    await page.evaluate(() => (window as any).toolbar.select('A2:B3'))
    await page
      .locator('.workbook-shared-sort-filter')
      .getByRole('button', { name: 'Sort', exact: true })
      .click()
    await page.getByRole('option', { name: 'Sort Z → A', exact: true }).click()
    await expect.poll(async () => (await snapshot()).sheets[0].cells.A2.rawValue).toBe('South')
    expect((await snapshot()).sheets[0].cells.B2.rawValue).toBe(-30)
    await page.evaluate(() => (window as any).toolbar.select('A1:B3'))
    await page
      .locator('.workbook-shared-sort-filter')
      .getByRole('button', { name: /^Filter\b/ })
      .click()
    await expect
      .poll(async () => (await snapshot()).sheets[0].features.filterEdits.filter.range.endRow)
      .toBe(2)
    const headerPosition = await page.evaluate(() => {
      const canvas = [...document.querySelectorAll('canvas')].find(
        (element) => element.getBoundingClientRect().height > 100,
      )!
      const box = canvas.getBoundingClientRect(),
        cell = (window as any).toolbar.cellRect('A1')
      return { x: box.x + cell.right - 9, y: box.y + cell.bottom - 9 }
    })
    await page.mouse.click(headerPosition.x, headerPosition.y)
    const filterPanel = page
      .locator('[data-u-comp="sheets-filter-panel"]')
      .filter({ has: page.locator('[data-u-comp="sheets-filter-panel-footer"]') })
    await expect(filterPanel).toBeVisible()
    await filterPanel
      .getByText('North', { exact: true })
      .locator('xpath=ancestor::div[label[@data-u-comp="checkbox"]][1]')
      .locator('label[data-u-comp="checkbox"]')
      .click()
    await filterPanel.getByRole('button', { name: 'Confirm', exact: true }).click()
    await expect(filterPanel).toBeHidden()
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.rowVisible(2))).toBe(false)
    await page
      .locator('.workbook-shared-sort-filter')
      .getByRole('button', { name: /^Filter\b/ })
      .click()
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.rowVisible(2))).toBe(true)
    await tab('Review')
    await expect(page.locator('#ribbon-panel-review')).toHaveClass(/workbook-desktop-panel/)
    await expect(
      page.locator('.workbook-desktop-header').getByRole('button', { name: /Protect Sheet/ }),
    ).toBeVisible()
    await page.locator('#ribbon-panel-review details > summary').click()
    await page.evaluate(() => (window as any).toolbar.select('B2'))
    await page.getByRole('textbox', { name: 'Note text', exact: true }).fill('Review this amount')
    await page.getByRole('button', { name: 'Save note', exact: true }).click()
    await expect
      .poll(async () => (await snapshot()).sheets[0].features.noteEdits.B2)
      .toBe('Review this amount')
    await page.evaluate(() => (window as any).toolbar.flush())
    await page.evaluate(() => (window as any).toolbar.editNote('Updated by another editor'))
    await expect(page.getByRole('textbox', { name: 'Note text', exact: true })).toHaveValue(
      'Updated by another editor',
    )
    await page.evaluate(() => (window as any).toolbar.select('A2:A3'))
    await page.getByRole('textbox', { name: 'Find', exact: true }).fill('South')
    await page.getByRole('textbox', { name: 'Replace with', exact: true }).fill('East')
    await page.getByRole('button', { name: 'Replace all in selection', exact: true }).click()
    await expect.poll(async () => (await snapshot()).sheets[0].cells.A2.rawValue).toBe('East')
    await tab('View')
    await expect(page.locator('#ribbon-panel-view')).toHaveClass(/workbook-desktop-panel/)
    await pageSettingChoice(page, 'Freeze Panes', 'Freeze Top Row')
    await expect
      .poll(() => page.evaluate(() => (window as any).toolbar.view().freeze.ySplit))
      .toBe(1)
    await pageSettingChoice(page, 'Zoom', '125%')
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.view().zoom)).toBe(1.25)
    const show = page.locator('.workbook-desktop-header .check-column-view')
    await show.getByRole('button', { name: /Gridlines/ }).click()
    await expect
      .poll(() => page.evaluate(() => (window as any).toolbar.view().gridlines))
      .toBe(false)
    await expect(show.getByRole('button', { name: /Gridlines/ }).locator('.check-box')).toHaveText(
      '',
    )
    // Right-to-left is saved sheet-view state; formula bar and cross highlight are view preferences.
    await show.getByRole('button', { name: /Right-to-Left/ }).click()
    await expect
      .poll(() =>
        page.evaluate(
          () => (window as any).toolbar.snapshot().sheets[0].features.pageSetupEdits.rightToLeft,
        ),
      )
      .toBe(true)
    await expect(
      show.getByRole('button', { name: /Right-to-Left/ }).locator('.check-box'),
    ).toHaveText('✓')
    await show.getByRole('button', { name: /Formula Bar/ }).click()
    await expect(page.locator('[data-u-comp="formula-bar"]').first()).toBeHidden()
    await show.getByRole('button', { name: /Formula Bar/ }).click()
    await expect(page.locator('[data-u-comp="formula-bar"]').first()).toBeVisible()
    await show.getByRole('button', { name: /Right-to-Left/ }).click()
    await expect
      .poll(() =>
        page.evaluate(
          () => (window as any).toolbar.snapshot().sheets[0].features.pageSetupEdits.rightToLeft,
        ),
      )
      .toBe(false)
    await expect(page.getByRole('alert')).toHaveCount(0)
    await tab('Home')
    await page.screenshot({ path: test.info().outputPath('excel-ribbon.png') })
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('ribbon tabs retain focus during engine selection sync and release it for a real grid click', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(true)
  })
  const home = page.getByRole('tab', { name: 'Home', exact: true })
  await home.focus()
  await page.evaluate(() => (window as any).toolbar.select('B2'))
  await expect(home).toBeFocused()
  const before = await page.evaluate(() => (window as any).toolbar.selection())
  await page.keyboard.press('ArrowRight')
  await expect(page.getByRole('tab', { name: 'Insert', exact: true })).toBeFocused()
  expect(await page.evaluate(() => (window as any).toolbar.selection())).toEqual(before)
  await expect
    .poll(() =>
      page.evaluate(() =>
        [...document.querySelectorAll('canvas')].some(
          (element) => element.getBoundingClientRect().height > 100,
        ),
      ),
    )
    .toBe(true)
  const position = await page.evaluate(() => {
    const canvas = [...document.querySelectorAll('canvas')].find(
      (element) => element.getBoundingClientRect().height > 100,
    )!
    const rect = canvas.getBoundingClientRect()
    const cell = (window as any).toolbar.cellRect('B2')
    return { x: rect.x + (cell.left + cell.right) / 2, y: rect.y + (cell.top + cell.bottom) / 2 }
  })
  await page.mouse.click(position.x, position.y)
  await page.keyboard.type('77')
  await page.keyboard.press('Enter')
  await expect
    .poll(() => page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells.B2.rawValue))
    .toBe(77)
})

test('shared inline chart editor retains failed drafts and flushes one pending commit', async ({
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
    await page.getByRole('tab', { name: 'Insert', exact: true }).click()
    await page.locator('#ribbon-panel-insert details > summary').click()
    await page.getByRole('button', { name: 'Add chart', exact: true }).click()
    await page.locator('.chart-edit-button').click()
    const panel = page.locator('.chart-editor')
    await panel.locator('.chart-editor-title').fill('Retry this draft')
    await page.evaluate(() => (window as any).toolbar.rejectNextSubmission('Chart write refused'))
    await panel.getByRole('button', { name: 'Apply', exact: true }).click()
    await expect(panel.locator('.chart-editor-error')).toHaveText('Chart write refused')
    await expect(panel.locator('.chart-editor-title')).toHaveValue('Retry this draft')
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(1)
    await page.evaluate(() => (window as any).toolbar.holdSubmissions())
    const calls = await page.evaluate(() => (window as any).toolbar.submissions())
    await panel.locator('.chart-editor-apply').evaluate((el: HTMLButtonElement) => {
      el.click()
      el.click()
    })
    await expect(panel).toHaveAttribute('aria-busy', 'true')
    await expect(panel.locator('.chart-editor-title')).toBeDisabled()
    await expect(panel.locator('.chart-editor-cancel')).toBeDisabled()
    await page.keyboard.press('Escape')
    await expect(panel).toHaveCount(1)
    await page.evaluate(() => {
      ;(window as any).chartFlush = 'waiting'
      void (window as any).toolbar.flush().then((ok: boolean) => {
        ;(window as any).chartFlush = ok ? 'ready' : 'failed'
      })
    })
    expect(await page.evaluate(() => (window as any).chartFlush)).toBe('waiting')
    expect(await page.evaluate(() => (window as any).toolbar.submissions())).toBe(calls + 1)
    await page.evaluate(() => (window as any).toolbar.releaseSubmissions())
    await expect(panel).toHaveCount(0)
    await expect.poll(() => page.evaluate(() => (window as any).chartFlush)).toBe('ready')
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(2)
    await expect(page.getByRole('img', { name: 'Retry this draft', exact: true })).toBeVisible()
    await page.locator('.chart-edit-button').click()
    await panel.locator('.chart-editor-title').fill('Discard this draft')
    await page.keyboard.press('Escape')
    await expect(panel).toHaveCount(0)
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(2)
    await expect(page.locator('.workbook-chart-tools')).toHaveCount(0)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('shared pie editor commits slice colors and rejects an overtaken draft', async ({ page }) => {
  page.setDefaultTimeout(8000)
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar()
    ;(window as any).toolbar.select('A1:B3')
  })
  try {
    await page.getByRole('tab', { name: 'Insert', exact: true }).click()
    await page.locator('#ribbon-panel-insert details > summary').click()
    await page.getByLabel('Chart', { exact: true }).selectOption('pie')
    await page.getByRole('button', { name: 'Add chart', exact: true }).click()
    await page.locator('.chart-edit-button').click()
    const panel = page.locator('.chart-editor')
    await panel.getByRole('button', { name: 'North', exact: true }).click()
    await page
      .getByRole('grid', { name: 'Standard Colors', exact: true })
      .getByRole('gridcell', { name: 'Red', exact: true })
      .click()
    await panel.getByRole('button', { name: 'Apply', exact: true }).click()
    await expect(panel).toHaveCount(0)
    const accepted = await page.evaluate(() => (window as any).toolbar.snapshot())
    expect(accepted.sheets[0].features.charts[0].chart.series[0].pointColors).toEqual([
      { index: 0, color: '#ff0000' },
    ])
    await page.locator('.chart-edit-button').click()
    await panel.locator('.chart-editor-title').fill('Overtaken draft')
    await page.evaluate(() => {
      void (window as any).toolbar.editNote('Concurrent change').catch(() => {})
    })
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.revision())).toBe(3)
    await expect(panel.locator('.chart-editor-title')).toHaveValue('Overtaken draft')
    await panel.getByRole('button', { name: 'Apply', exact: true }).click()
    await expect(panel).toHaveCount(0)
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(3)
    expect(
      await page.evaluate(
        () => (window as any).toolbar.snapshot().sheets[0].features.charts[0].chart,
      ),
    ).toEqual(accepted.sheets[0].features.charts[0].chart)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
  }
})

test('original chart uses the shared property editor and preserves immutable source data', async ({
  page,
}) => {
  page.setDefaultTimeout(8000)
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(
      false,
      false,
      false,
      false,
      false,
      false,
      true,
    )
  })
  try {
    const before = await page.evaluate(() => (window as any).toolbar.snapshot())
    await page.locator('.chart-edit-button').click()
    const panel = page.locator('.chart-editor')
    await panel.locator('.chart-editor-title').fill('Original sales')
    await panel.locator('.chart-editor-type').click()
    await page.locator('button[role=option][data-value=line]').click()
    await expect(panel.getByRole('listbox')).toHaveCount(0)
    await panel.getByRole('button', { name: 'Sales', exact: true }).click()
    await page
      .getByRole('grid', { name: 'Standard Colors', exact: true })
      .getByRole('gridcell', { name: 'Red', exact: true })
      .click()
    await panel.getByRole('button', { name: 'Apply', exact: true }).click()
    await expect(panel).toHaveCount(0)
    await expect(page.getByRole('img', { name: 'Original sales', exact: true })).toBeVisible()
    const after = await page.evaluate(() => (window as any).toolbar.snapshot())
    expect(after.sheets[0].features.chartEdits['xl/charts/chart1.xml']).toEqual({
      chartPath: 'xl/charts/chart1.xml',
      title: 'Original sales',
      chartType: 'line',
      seriesColors: { '0': '#ff0000' },
    })
    expect(after.features).toEqual(before.features)
    expect(after.sheets[0].cells).toEqual(before.sheets[0].cells)
    await page.locator('.chart-edit-button').click()
    await expect(panel.locator('.chart-editor-title')).toHaveValue('Original sales')
    await expect(panel.locator('.chart-editor-type')).toContainText('Line')
    await panel.getByRole('button', { name: 'Cancel', exact: true }).click()
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(1)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
  }
})

test('original chart range apply waits for host reads and save, validates input and retains references', async ({
  page,
}) => {
  page.setDefaultTimeout(8000)
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(
      false,
      false,
      false,
      false,
      false,
      true,
      true,
    )
  })
  try {
    await page.locator('.chart-edit-button').click()
    const panel = page.locator('.chart-editor')
    const fields = panel.locator('.chart-editor-range')
    await fields.nth(0).fill('A78:B79')
    await panel.locator('.chart-editor-apply').click()
    await expect(panel.locator('.chart-editor-error')).toHaveText('chart_vector_required')
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(0)
    await fields.nth(0).fill('D78')
    await panel.locator('.chart-editor-apply').click()
    await expect(panel.locator('.chart-editor-error')).toHaveText('chart_formula_pending')
    await fields.nth(0).fill('B79:B81')
    await fields.nth(1).fill('A79:A81')
    await page.evaluate(() => (window as any).toolbar.holdReads())
    await panel.locator('.chart-editor-apply').evaluate((el: HTMLButtonElement) => {
      el.click()
      el.click()
    })
    await expect(panel).toHaveAttribute('aria-busy', 'true')
    await page.evaluate(() => {
      ;(window as any).rangeFlush = 'waiting'
      void (window as any).toolbar.flush().then((ok: boolean) => {
        ;(window as any).rangeFlush = ok ? 'ready' : 'failed'
      })
    })
    expect(await page.evaluate(() => (window as any).rangeFlush)).toBe('waiting')
    await expect(panel.locator('.chart-editor-cancel')).toBeDisabled()
    await page.keyboard.press('Escape')
    await expect(panel).toHaveCount(1)
    await page.evaluate(() => (window as any).toolbar.releaseReads())
    await expect(panel).toHaveCount(0)
    await expect.poll(() => page.evaluate(() => (window as any).rangeFlush)).toBe('ready')
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(1)
    const series = await page.evaluate(
      () =>
        (window as any).toolbar.snapshot().sheets[0].features.chartEdits['xl/charts/chart1.xml']
          .series[0],
    )
    expect(series).toEqual({
      index: 0,
      values: [12, 8, 18],
      valuesRef: "'Data'!$B$79:$B$81",
      categories: ['2020', '2021', '2022'],
      categoriesRef: "'Data'!$A$79:$A$81",
    })
    await page.evaluate(() => (window as any).toolbar.editChartSource(42))
    expect(
      await page.evaluate(
        () =>
          (window as any).toolbar.snapshot().sheets[0].features.chartEdits['xl/charts/chart1.xml']
            .series[0].values,
      ),
    ).toEqual([42, 8, 18])
    await page.locator('.chart-edit-button').click()
    await expect(panel.locator('.chart-editor-range').nth(0)).toHaveAttribute(
      'placeholder',
      "'Data'!$B$79:$B$81",
    )
    await panel.locator('.chart-editor-cancel').click()
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
  }
})

test('chart range read overtaken by another transaction cannot publish its stale series', async ({
  page,
}) => {
  page.setDefaultTimeout(8000)
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(
      false,
      false,
      false,
      false,
      false,
      true,
      true,
    )
  })
  try {
    await page.locator('.chart-edit-button').click()
    const panel = page.locator('.chart-editor')
    await panel.locator('.chart-editor-range').nth(0).fill('B79:B81')
    await page.evaluate(() => (window as any).toolbar.holdReads())
    await panel.locator('.chart-editor-apply').click()
    await expect(panel).toHaveAttribute('aria-busy', 'true')
    await page.evaluate(() => {
      void (window as any).toolbar.editNote('Concurrent range read')
    })
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.revision())).toBe(1)
    await page.evaluate(() => (window as any).toolbar.releaseReads())
    await expect(panel).toHaveCount(0)
    const after = await page.evaluate(() => (window as any).toolbar.snapshot())
    expect(after.sheets[0].features.chartEdits).toBeUndefined()
    expect(after.sheets[0].features.noteEdits.B2).toBe('Concurrent range read')
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(1)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
  }
})

test('uncached original chart hydrates its values, categories and series name without authoring metadata', async ({
  page,
}) => {
  page.setDefaultTimeout(8000)
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(
      false,
      false,
      false,
      false,
      false,
      false,
      'uncached',
    )
  })
  try {
    await expect(page.locator('.xlsx-print-visual .chart-svg')).toHaveCount(1)
    await expect(page.locator('.xlsx-print-visual .chart-svg')).toContainText('North')
    await expect(page.locator('.xlsx-print-visual .chart-svg')).toContainText('South')
    await page.locator('.chart-edit-button').click()
    const panel = page.locator('.chart-editor')
    await expect(panel.getByRole('button', { name: 'Sales', exact: true })).toBeVisible()
    await panel.locator('.chart-editor-cancel').click()
    const state = await page.evaluate(() => (window as any).toolbar.snapshot())
    expect(state.features.originalMetadata.visuals[0].chart.series[0].values).toEqual([])
    expect(state.features.originalMetadata.visuals[0].chart.series[0].categories).toEqual([])
    expect(state.features.originalMetadata.visuals[0].chart.series[0].name).toBe('Series1')
    expect(state.sheets[0].features.chartEdits).toBeUndefined()
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(0)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
  }
})

test('desktop chart context menu opens the persistent format pane and commits its controls', async ({
  page,
}) => {
  page.setDefaultTimeout(8000)
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(
      false,
      false,
      false,
      false,
      false,
      false,
      true,
    )
  })
  try {
    const visual = page.locator('.xlsx-print-visual .xlsx-chart')
    await visual.click({ button: 'right', position: { x: 50, y: 15 } })
    const menu = page.locator('.chart-menu').first()
    await expect(menu.getByRole('button', { name: 'Select Data…', exact: true })).toBeEnabled()
    await menu.getByRole('button', { name: 'Add Data Labels', exact: true }).click()
    await expect
      .poll(() =>
        page.evaluate(
          () =>
            (window as any).toolbar.snapshot().sheets[0].features.chartEdits['xl/charts/chart1.xml']
              .dataLabels,
        ),
      )
      .toBe('value')
    await visual.click({ button: 'right', position: { x: 50, y: 15 } })
    await menu.getByRole('button', { name: 'Format Selection…', exact: true }).click()
    const pane = page.locator('.chart-format-pane')
    await expect(pane).toBeVisible()
    await pane.getByRole('button', { name: 'Legend Position', exact: true }).click()
    await pane.getByRole('option', { name: 'Bottom', exact: true }).click()
    await expect(pane).toHaveAttribute('aria-busy', 'false')
    await expect(pane.getByRole('button', { name: 'Legend Position', exact: true })).toContainText(
      'Bottom',
    )
    await pane.getByLabel('Gridlines', { exact: true }).uncheck()
    await expect
      .poll(() =>
        page.evaluate(
          () =>
            (window as any).toolbar.snapshot().sheets[0].features.chartEdits['xl/charts/chart1.xml']
              .gridlines,
        ),
      )
      .toBe(false)
    const minimum = pane.getByLabel('Minimum', { exact: true })
    await minimum.fill('0')
    await minimum.press('Enter')
    await expect
      .poll(() =>
        page.evaluate(
          () =>
            (window as any).toolbar.snapshot().sheets[0].features.chartEdits['xl/charts/chart1.xml']
              .valueAxis?.min,
        ),
      )
      .toBe(0)
    await pane.getByLabel('Maximum', { exact: true }).fill('50')
    await pane.getByLabel('Maximum', { exact: true }).press('Enter')
    await expect
      .poll(() =>
        page.evaluate(
          () =>
            (window as any).toolbar.snapshot().sheets[0].features.chartEdits['xl/charts/chart1.xml']
              .valueAxis?.max,
        ),
      )
      .toBe(50)
    await minimum.fill('60')
    await minimum.press('Enter')
    await expect(pane.getByRole('alert')).toHaveText('Maximum must be greater than minimum.')
    await minimum.fill('5')
    await minimum.press('Enter')
    await expect(pane.getByRole('alert')).toHaveCount(0)
    await expect
      .poll(() =>
        page.evaluate(
          () =>
            (window as any).toolbar.snapshot().sheets[0].features.chartEdits['xl/charts/chart1.xml']
              .valueAxis?.min,
        ),
      )
      .toBe(5)
    // Saving must blur the active pane input and wait for its canonical edit.
    await pane.getByRole('textbox', { name: 'Title', exact: true }).fill('Saved from format pane')
    expect(await page.evaluate(() => (window as any).toolbar.flush())).toBe(true)
    await expect(
      page.getByRole('img', { name: 'Saved from format pane', exact: true }),
    ).toBeVisible()
    await expect(pane.getByRole('textbox', { name: 'Title', exact: true })).toHaveValue(
      'Saved from format pane',
    )
    await pane.getByRole('button', { name: 'Close', exact: true }).click()
    await expect(pane).toHaveCount(0)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
  }
})

test('chart format pane keeps rejected text drafts and flushes one pending property edit', async ({
  page,
}) => {
  page.setDefaultTimeout(8000)
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(
      false,
      false,
      false,
      false,
      false,
      false,
      true,
    )
  })
  try {
    await page
      .locator('.xlsx-print-visual .xlsx-chart')
      .click({ button: 'right', position: { x: 50, y: 15 } })
    await page
      .locator('.chart-menu')
      .first()
      .getByRole('button', { name: 'Format Selection…', exact: true })
      .click()
    const pane = page.locator('.chart-format-pane')
    const title = pane.getByRole('textbox', { name: 'Title', exact: true })
    await page.evaluate(() => (window as any).toolbar.rejectNextSubmission('Format write refused'))
    await title.fill('Retain my draft')
    await title.press('Enter')
    await expect(pane.getByRole('alert')).toHaveText('Format write refused')
    await expect(title).toHaveValue('Retain my draft')
    await title.focus()
    // A background backup must wait for accepted edits without forcing this draft's blur.
    expect(
      await page.evaluate(() => (window as any).toolbar.flush({ commitChartDraft: false })),
    ).toBe(true)
    await expect(title).toBeFocused()
    await expect(title).toHaveValue('Retain my draft')
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(0)
    await page.evaluate(() => (window as any).toolbar.holdSubmissions())
    await title.focus()
    await title.press('Enter')
    await expect(pane).toHaveAttribute('aria-busy', 'true')
    await expect(pane.getByRole('button', { name: 'Close', exact: true })).toBeDisabled()
    await expect(title).toBeDisabled()
    await page.evaluate(() => {
      ;(window as any).formatFlush = 'waiting'
      void (window as any).toolbar.flush().then((ok: boolean) => {
        ;(window as any).formatFlush = ok ? 'ready' : 'failed'
      })
    })
    expect(await page.evaluate(() => (window as any).formatFlush)).toBe('waiting')
    await page.evaluate(() => (window as any).toolbar.releaseSubmissions())
    await expect(pane).toHaveAttribute('aria-busy', 'false')
    await expect.poll(() => page.evaluate(() => (window as any).formatFlush)).toBe('ready')
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(1)
    await expect(page.getByRole('img', { name: 'Retain my draft', exact: true })).toBeVisible()
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
  }
})

test('desktop doughnut pane commits keyboard sliders once and scopes fill to an actual slice', async ({
  page,
}) => {
  page.setDefaultTimeout(8000)
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(true)
    ;(window as any).toolbar.select('A1:B3')
  })
  try {
    await page.getByRole('tab', { name: 'Insert', exact: true }).click()
    await page
      .locator('.workbook-desktop-header')
      .getByRole('button', { name: 'Doughnut chart — from selection', exact: true })
      .click()
    const visual = page.locator('.xlsx-print-visual .xlsx-chart')
    await visual.click({ button: 'right', position: { x: 50, y: 15 } })
    await page
      .locator('.chart-menu')
      .first()
      .getByRole('button', { name: 'Format Selection…', exact: true })
      .click()
    const pane = page.locator('.chart-format-pane')
    const hole = pane.getByLabel('Doughnut Hole Size', { exact: true })
    await hole.focus()
    await hole.press('ArrowRight')
    await hole.press('ArrowRight')
    await hole.press('Tab')
    await expect
      .poll(() =>
        page.evaluate(
          () => (window as any).toolbar.snapshot().sheets[0].features.charts[0].chart.holeSizePct,
        ),
      )
      .toBe(60)
    await expect(pane).toHaveAttribute('aria-busy', 'false')
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(2)
    const explosion = pane.getByLabel('Pie Explosion', { exact: true })
    await explosion.focus()
    await explosion.press('ArrowRight')
    await explosion.press('ArrowRight')
    await explosion.press('ArrowRight')
    await explosion.press('Tab')
    await expect
      .poll(() =>
        page.evaluate(
          () =>
            (window as any).toolbar.snapshot().sheets[0].features.charts[0].chart.series[0]
              .explosionPct,
        ),
      )
      .toBe(15)
    await expect(pane).toHaveAttribute('aria-busy', 'false')
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(3)
    const point = await page
      .locator('.pie-svg path')
      .first()
      .evaluate((element: SVGPathElement) => {
        const box = element.getBBox(),
          matrix = element.getScreenCTM()!
        for (const fx of [0.15, 0.3, 0.5, 0.7, 0.85])
          for (const fy of [0.15, 0.3, 0.5, 0.7, 0.85]) {
            const local = new DOMPoint(box.x + box.width * fx, box.y + box.height * fy)
            if (!element.isPointInFill(local)) continue
            const screen = local.matrixTransform(matrix)
            if (document.elementFromPoint(screen.x, screen.y) === element)
              return { x: screen.x, y: screen.y }
          }
        throw new Error('No visible slice point')
      })
    await page.mouse.click(point.x, point.y)
    await expect(pane.getByRole('button', { name: 'Fill Color', exact: true })).toBeVisible()
    const perPoint = pane.getByLabel('Point Explosion', { exact: true })
    await perPoint.focus()
    await perPoint.press('ArrowRight')
    await perPoint.press('Tab')
    await expect
      .poll(() =>
        page.evaluate(
          () =>
            (window as any).toolbar.snapshot().sheets[0].features.charts[0].chart.series[0]
              .pointExplosions,
        ),
      )
      .toEqual([{ index: 0, pct: 20 }])
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(4)
    await pane.getByRole('button', { name: 'Fill Color', exact: true }).click()
    await page
      .getByRole('grid', { name: 'Standard Colors', exact: true })
      .getByRole('gridcell', { name: 'Red', exact: true })
      .click()
    await expect
      .poll(() =>
        page.evaluate(
          () =>
            (window as any).toolbar.snapshot().sheets[0].features.charts[0].chart.series[0]
              .pointColors,
        ),
      )
      .toEqual([{ index: 0, color: '#ff0000' }])
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
  }
})

for (const original of [true, false]) {
  test(`Select Data ${original ? 'original' : 'session'} edits retain rejected drafts, await save and transpose through host history`, async ({
    page,
  }) => {
    page.setDefaultTimeout(8000)
    await page.goto(base)
    await page.evaluate(async (original) => {
      const { mountManualToolbar } = await import(
        /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
      )
      ;(window as any).toolbar = original
        ? await mountManualToolbar(false, false, false, false, false, false, true)
        : await mountManualToolbar(true)
      if (!original) (window as any).toolbar.select('A1:B3')
    }, original)
    try {
      if (!original) {
        await page.getByRole('tab', { name: 'Insert', exact: true }).click()
        await page
          .locator('.workbook-desktop-header')
          .getByRole('button', { name: 'Column chart — from selection', exact: true })
          .click()
      }
      const open = async () => {
        await page
          .locator('.xlsx-print-visual .xlsx-chart')
          .click({ button: 'right', position: { x: 50, y: 15 } })
        await page
          .locator('.chart-menu')
          .first()
          .getByRole('button', { name: 'Select Data…', exact: true })
          .click()
      }
      await open()
      const dialog = page.locator('.select-data-dialog'),
        rows = dialog.locator('.select-data-rows .select-data-row')
      const before = await page.evaluate(() => (window as any).toolbar.revision())
      await rows.nth(0).locator('input').nth(0).fill('Renamed series')
      await dialog.getByRole('button', { name: '＋ Add Series', exact: true }).click()
      await rows.nth(1).locator('input').nth(0).fill('Added series')
      await rows.nth(1).locator('input').nth(1).fill('B2:B3')
      await rows.nth(1).locator('input').nth(2).fill('A2:A3')
      await rows.nth(1).getByRole('button', { name: 'Move Up', exact: true }).click()
      await page.evaluate(() =>
        (window as any).toolbar.rejectNextSubmission('Series write refused'),
      )
      await dialog.getByRole('button', { name: 'OK', exact: true }).click()
      await expect(dialog.locator('.chart-editor-error')).toHaveText('Series write refused')
      await expect(rows.nth(0).locator('input').nth(0)).toHaveValue('Added series')
      expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(before)
      await page.evaluate(() => (window as any).toolbar.holdSubmissions())
      await dialog.getByRole('button', { name: 'OK', exact: true }).click()
      await expect(dialog.getByRole('button', { name: 'Cancel', exact: true })).toBeDisabled()
      await expect(rows.nth(0).locator('input').nth(0)).toBeDisabled()
      await page.keyboard.press('Escape')
      await expect(dialog).toBeVisible()
      await page.evaluate(() => {
        ;(window as any).selectionFlush = 'waiting'
        void (window as any).toolbar.flush().then((ok: boolean) => {
          ;(window as any).selectionFlush = ok ? 'ready' : 'failed'
        })
      })
      expect(await page.evaluate(() => (window as any).selectionFlush)).toBe('waiting')
      await page.evaluate(() => (window as any).toolbar.releaseSubmissions())
      await expect(dialog).toHaveCount(0)
      await expect.poll(() => page.evaluate(() => (window as any).selectionFlush)).toBe('ready')
      expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(before + 1)
      const current = () =>
        page.evaluate((original) => {
          const s = (window as any).toolbar.snapshot().sheets[0]
          return original
            ? s.features.chartEdits['xl/charts/chart1.xml'].seriesSet
            : s.features.charts[0].chart.series
        }, original)
      expect((await current()).map((s: any) => s.name)).toEqual(['Added series', 'Renamed series'])
      await open()
      await dialog.getByRole('button', { name: 'Switch Row/Column', exact: true }).click()
      await expect(dialog).toHaveCount(0)
      expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(before + 2)
      expect((await current())[0].categories).toEqual(['Added series', 'Renamed series'])
      await open()
      await rows.nth(0).getByRole('button', { name: 'Remove Series', exact: true }).click()
      await dialog.getByRole('button', { name: 'Cancel', exact: true }).click()
      expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(before + 2)
      await open()
      await rows.nth(0).locator('input').nth(0).fill('Stale selection draft')
      await page.evaluate(() => (window as any).toolbar.seedClearTargets())
      const advanced = await page.evaluate(() => (window as any).toolbar.revision())
      await dialog.getByRole('button', { name: 'OK', exact: true }).click()
      await expect(dialog.locator('.chart-editor-error')).toContainText('The document changed')
      await expect(rows.nth(0).locator('input').nth(0)).toHaveValue('Stale selection draft')
      expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(advanced)
      await dialog.getByRole('button', { name: 'Cancel', exact: true }).click()
    } finally {
      await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
    }
  })
}

for (const original of [true, false]) {
  test(`shared Chart Design ribbon edits ${original ? 'source' : 'session'} charts with atomic presets and frozen text prompts`, async ({
    page,
  }) => {
    page.setDefaultTimeout(8000)
    await page.goto(base)
    await page.evaluate(async (original) => {
      const { mountManualToolbar } = await import(
        /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
      )
      ;(window as any).toolbar = original
        ? await mountManualToolbar(false, false, false, false, false, false, true)
        : await mountManualToolbar(true)
      if (!original) (window as any).toolbar.select('A1:B3')
    }, original)
    try {
      const chartTab = page.getByRole('tab', { name: 'Chart Design', exact: true })
      await expect(chartTab).toHaveCount(0)
      if (!original) {
        await page.getByRole('tab', { name: 'Insert', exact: true }).click()
        await page
          .locator('.workbook-desktop-header')
          .getByRole('button', { name: 'Column chart — from selection', exact: true })
          .click()
      }
      await page.locator('.xlsx-print-visual .xlsx-chart').click({ position: { x: 50, y: 15 } })
      await expect(chartTab).toBeVisible()
      await page.getByRole('tab', { name: 'Home', exact: true }).focus()
      await page.keyboard.press('End')
      await expect(chartTab).toBeFocused()
      await expect(chartTab).toHaveAttribute('aria-selected', 'true')
      const ribbon = page.locator('#ribbon-panel-chartDesign')
      const read = () =>
        page.evaluate((original) => {
          const s = (window as any).toolbar.snapshot().sheets[0]
          return original
            ? s.features.chartEdits?.['xl/charts/chart1.xml']
            : s.features.charts[0].chart
        }, original)
      const pick = async (label: string, name: string) => {
        await ribbon.getByRole('button', { name: label, exact: true }).click()
        await ribbon.getByRole('option', { name, exact: true }).click()
      }
      const before = await page.evaluate(() => (window as any).toolbar.revision())
      await pick('Quick Layout', 'Layout 2 — Legend top, value labels')
      await expect.poll(read).toMatchObject({ legend: 'top', dataLabels: 'value' })
      expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(before + 1)
      await pick('Add Chart Element', 'Legend: Bottom')
      await expect.poll(read).toMatchObject({ legend: 'bottom' })
      await ribbon.getByRole('button', { name: 'Add Chart Element', exact: true }).click()
      await expect(
        ribbon.getByRole('option', { name: 'Legend: Bottom ✓', exact: true }),
      ).toBeVisible()
      await page.keyboard.press('Escape')
      await pick('Stacked', 'Stacked')
      await expect.poll(read).toMatchObject({ grouping: 'stacked' })
      await pick('Change Colors', 'Monochromatic Blue')
      await expect
        .poll(async () => {
          const c = await read()
          return original ? c.seriesColors['0'] : c.series[0].color
        })
        .toBe('#1f4e79')
      await pick('Add Chart Element', 'Chart Title…')
      const title = page.getByRole('dialog', { name: 'Chart Title', exact: true })
      await title.getByRole('textbox').fill('Chart ribbon draft')
      await page.evaluate(() => (window as any).toolbar.rejectNextSubmission('Chart title refused'))
      await title.getByRole('button', { name: 'OK', exact: true }).click()
      await expect(title.getByRole('alert')).toHaveText('Chart title refused')
      await expect(title.getByRole('textbox')).toHaveValue('Chart ribbon draft')
      await page.evaluate(() => (window as any).toolbar.holdSubmissions())
      await title.getByRole('button', { name: 'OK', exact: true }).click()
      await expect(title.getByRole('button', { name: 'Cancel', exact: true })).toBeDisabled()
      await page.keyboard.press('Escape')
      await expect(title).toBeVisible()
      await page.evaluate(() => {
        ;(window as any).chartTitleFlush = 'waiting'
        void (window as any).toolbar.flush().then((ok: boolean) => {
          ;(window as any).chartTitleFlush = ok ? 'ready' : 'failed'
        })
      })
      expect(await page.evaluate(() => (window as any).chartTitleFlush)).toBe('waiting')
      await page.evaluate(() => (window as any).toolbar.releaseSubmissions())
      await expect(title).toHaveCount(0)
      await expect.poll(() => page.evaluate(() => (window as any).chartTitleFlush)).toBe('ready')
      await expect.poll(read).toMatchObject({ title: 'Chart ribbon draft' })
      await pick('Add Chart Element', 'Axis Title: Category…')
      const axis = page.getByRole('dialog', { name: 'Category Axis Title', exact: true })
      await axis.getByRole('textbox').fill('Regions')
      await axis.getByRole('button', { name: 'OK', exact: true }).click()
      await expect(axis).toHaveCount(0)
      await expect.poll(read).toMatchObject({ axisTitles: { category: 'Regions' } })
      await pick('Add Chart Element', 'Axis Title: Category… ✓')
      await expect(axis.getByRole('textbox')).toHaveValue('Regions')
      await page.evaluate(() => (window as any).toolbar.seedClearTargets())
      await axis.getByRole('button', { name: 'OK', exact: true }).click()
      await expect(axis.getByRole('alert')).toContainText('The document changed')
      await axis.getByRole('button', { name: 'Cancel', exact: true }).click()
      await ribbon.getByRole('button', { name: 'Line', exact: true }).click()
      await expect
        .poll(read)
        .toMatchObject(original ? { chartType: 'line' } : { chartTypes: ['lineChart'] })
      await ribbon.getByRole('button', { name: 'Select Data', exact: true }).click()
      const select = page.locator('.select-data-dialog')
      await expect(select).toBeVisible()
      await select.getByRole('button', { name: 'Cancel', exact: true }).click()
      await ribbon.getByRole('button', { name: 'Format Pane', exact: true }).click()
      await expect(page.locator('.chart-format-pane')).toBeVisible()
      await page
        .locator('.chart-format-pane')
        .getByRole('button', { name: 'Close', exact: true })
        .click()
      await ribbon.getByRole('button', { name: 'Delete Chart', exact: true }).click()
      await expect(chartTab).toHaveCount(0)
      await expect(page.getByRole('tab', { name: 'Home', exact: true })).toHaveAttribute(
        'aria-selected',
        'true',
      )
    } finally {
      await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
    }
  })
}

for (const [family, label] of [
  ['pie', 'Pie'],
  ['scatter', 'Scatter'],
  ['combo', 'Combo (column + line)'],
] as const) {
  test(`Chart Design applies ${family} family capabilities and palette semantics`, async ({
    page,
  }) => {
    page.setDefaultTimeout(8000)
    await page.goto(base)
    await page.evaluate(async () => {
      const { mountManualToolbar } = await import(
        /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
      )
      ;(window as any).toolbar = await mountManualToolbar(true)
      ;(window as any).toolbar.select('A1:B3')
    })
    try {
      await page.getByRole('tab', { name: 'Insert', exact: true }).click()
      await page
        .locator('.workbook-desktop-header')
        .getByRole('button', { name: `${label} chart — from selection`, exact: true })
        .click()
      await page.locator('.xlsx-print-visual .xlsx-chart').click({ position: { x: 50, y: 15 } })
      await page.getByRole('tab', { name: 'Chart Design', exact: true }).click()
      const ribbon = page.locator('#ribbon-panel-chartDesign')
      await expect(ribbon.getByRole('button', { name: 'Column', exact: true }))[
        family === 'pie' ? 'toBeEnabled' : 'toBeDisabled'
      ]()
      await ribbon.getByRole('button', { name: 'Add Chart Element', exact: true }).click()
      await expect(ribbon.getByRole('option', { name: /Axis Title:/ })).toHaveCount(0)
      await expect(ribbon.getByRole('option', { name: /Data Labels:/ })).toHaveCount(
        family === 'pie' ? 3 : family === 'combo' ? 2 : 0,
      )
      await page.keyboard.press('Escape')
      await ribbon.getByRole('button', { name: 'Quick Layout', exact: true }).click()
      await ribbon
        .getByRole('option', {
          name:
            family === 'pie'
              ? 'Layout 4 — No legend, name + % labels'
              : 'Layout 4 — No legend, value labels',
          exact: true,
        })
        .click()
      await expect
        .poll(() =>
          page.evaluate(
            () => (window as any).toolbar.snapshot().sheets[0].features.charts[0].chart.legend,
          ),
        )
        .toBe('none')
      const chart = await page.evaluate(
        () => (window as any).toolbar.snapshot().sheets[0].features.charts[0].chart,
      )
      expect(chart.dataLabels).toBe(
        family === 'pie' ? 'category-percent' : family === 'combo' ? 'value' : undefined,
      )
      await ribbon.getByRole('button', { name: 'Change Colors', exact: true }).click()
      await ribbon.getByRole('option', { name: 'Warm', exact: true }).click()
      await expect.poll(() => page.evaluate(() => (window as any).toolbar.revision())).toBe(3)
      const colored = await page.evaluate(
        () => (window as any).toolbar.snapshot().sheets[0].features.charts[0].chart.series[0],
      )
      if (family === 'pie')
        expect(colored.pointColors).toEqual([
          { index: 0, color: '#c00000' },
          { index: 1, color: '#ed7d31' },
        ])
      else expect(colored.color).toBe('#c00000')
    } finally {
      await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
    }
  })
}

for (const original of [true, false]) {
  test(`shared visual dialogs retain ${original ? 'source' : 'session'} alt text and placement through pending, failed and stale commits`, async ({
    page,
  }) => {
    page.setDefaultTimeout(8000)
    await page.goto(base)
    await page.evaluate(async (original) => {
      const { mountManualToolbar } = await import(
        /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
      )
      ;(window as any).toolbar = original
        ? await mountManualToolbar(false, false, false, false, false, false, true)
        : await mountManualToolbar(true)
      if (!original) (window as any).toolbar.select('A1:B3')
    }, original)
    try {
      if (!original) {
        await page.getByRole('tab', { name: 'Insert', exact: true }).click()
        await page
          .locator('.workbook-desktop-header')
          .getByRole('button', { name: 'Column chart — from selection', exact: true })
          .click()
      }
      const open = async (name: string) => {
        await page
          .locator('.xlsx-print-visual .xlsx-chart')
          .click({ button: 'right', position: { x: 50, y: 15 } })
        await page.locator('.chart-menu').first().getByRole('button', { name, exact: true }).click()
      }
      const current = () =>
        page.evaluate((original) => {
          const s = (window as any).toolbar.snapshot().sheets[0]
          return original ? s.features.visualEdits?.['original-chart'] : s.features.charts[0]
        }, original)
      await open('Edit Alt Text…')
      const alt = page.getByRole('dialog', { name: 'Alt Text', exact: true })
      await alt.getByRole('textbox').fill('Sales & regional comparison')
      await page.evaluate(() =>
        (window as any).toolbar.rejectNextSubmission('Visual write refused'),
      )
      await alt.getByRole('button', { name: 'OK', exact: true }).click()
      await expect(alt.getByRole('alert')).toHaveText('Visual write refused')
      await expect(alt.getByRole('textbox')).toHaveValue('Sales & regional comparison')
      await page.evaluate(() => (window as any).toolbar.holdSubmissions())
      await alt.getByRole('button', { name: 'OK', exact: true }).click()
      await expect(alt.getByRole('textbox')).toBeDisabled()
      await expect(alt.getByRole('button', { name: 'Cancel', exact: true })).toBeDisabled()
      await page.keyboard.press('Escape')
      await expect(alt).toBeVisible()
      await page.evaluate(() => {
        ;(window as any).visualFlush = 'waiting'
        void (window as any).toolbar.flush().then((ok: boolean) => {
          ;(window as any).visualFlush = ok ? 'ready' : 'failed'
        })
      })
      expect(await page.evaluate(() => (window as any).visualFlush)).toBe('waiting')
      await page.evaluate(() => (window as any).toolbar.releaseSubmissions())
      await expect(alt).toHaveCount(0)
      await expect.poll(() => page.evaluate(() => (window as any).visualFlush)).toBe('ready')
      await expect.poll(current).toMatchObject({ altText: 'Sales & regional comparison' })
      await open('Edit Alt Text…')
      await expect(alt.getByRole('textbox')).toHaveValue('Sales & regional comparison')
      await alt.getByRole('button', { name: 'Cancel', exact: true }).click()
      await open('Size and Properties…')
      const size = page.getByRole('dialog', { name: 'Size and Properties', exact: true })
      await size.getByRole('radio', { name: "Don't move or size with cells", exact: true }).check()
      await size.getByRole('button', { name: 'OK', exact: true }).click()
      await expect(size).toHaveCount(0)
      await expect
        .poll(current)
        .toMatchObject({ altText: 'Sales & regional comparison', editAs: 'absolute' })
      await open('Size and Properties…')
      await expect(
        size.getByRole('radio', { name: "Don't move or size with cells", exact: true }),
      ).toBeChecked()
      await size.getByRole('radio', { name: "Move but don't size with cells", exact: true }).check()
      await page.evaluate(() => (window as any).toolbar.seedClearTargets())
      await size.getByRole('button', { name: 'OK', exact: true }).click()
      await expect(size.getByRole('alert')).toContainText('The document changed')
      await expect(
        size.getByRole('radio', { name: "Move but don't size with cells", exact: true }),
      ).toBeChecked()
      await size.getByRole('button', { name: 'Cancel', exact: true }).click()
      await expect.poll(current).toMatchObject({ editAs: 'absolute' })
    } finally {
      await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
    }
  })
  test(`visual layer menu reorders ${original ? 'source' : 'session'} charts and preserves the other drawing properties`, async ({
    page,
  }) => {
    page.setDefaultTimeout(8000)
    await page.goto(base)
    await page.evaluate(async (original) => {
      const { mountManualToolbar } = await import(
        /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
      )
      ;(window as any).toolbar = original
        ? await mountManualToolbar(false, false, false, false, false, false, 'multiple')
        : await mountManualToolbar()
      if (!original) await (window as any).toolbar.seedVisualLayers()
    }, original)
    try {
      const order = () =>
        page
          .locator('.xlsx-print-visual .xlsx-chart svg.chart-svg')
          .evaluateAll((nodes) => nodes.map((n) => n.getAttribute('aria-label')))
      await expect.poll(order).toEqual(['Layer 0', 'Layer 1', 'Layer 2'])
      const layer = page.getByRole('img', { name: 'Layer 0', exact: true })
      const before = await page.evaluate(() => (window as any).toolbar.revision())
      await layer.click({ button: 'right', position: { x: 45, y: 10 } })
      let menu = page.locator('.chart-menu').first()
      await menu.getByRole('button', { name: 'Bring to Front', exact: true }).hover()
      await page
        .locator('.chart-menu-child')
        .getByRole('button', { name: 'Bring to Front', exact: true })
        .click()
      await expect.poll(order).toEqual(['Layer 1', 'Layer 2', 'Layer 0'])
      expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(before + 1)
      await layer.click({ button: 'right', position: { x: 45, y: 10 } })
      menu = page.locator('.chart-menu').first()
      await menu.getByRole('button', { name: 'Send to Back', exact: true }).hover()
      await page
        .locator('.chart-menu-child')
        .getByRole('button', { name: 'Send Backward', exact: true })
        .click()
      await expect.poll(order).toEqual(['Layer 1', 'Layer 0', 'Layer 2'])
      expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(before + 2)
      await page.evaluate(() => (window as any).toolbar.overlapVisualLayers())
      const front = page.getByRole('img', { name: 'Layer 2', exact: true })
      const box = (await front.boundingBox())!
      const point = { x: box.x + 10, y: box.y + 10 }
      const top = () =>
        page.evaluate((point) => {
          for (const node of document.elementsFromPoint(point.x, point.y)) {
            const chart = node.closest('.xlsx-chart')
            if (chart) return chart.querySelector('svg.chart-svg')?.getAttribute('aria-label')
          }
          return null
        }, point)
      await expect.poll(top).toBe('Layer 2')
      await page.mouse.click(point.x, point.y, { button: 'right' })
      menu = page.locator('.chart-menu').first()
      await menu.getByRole('button', { name: 'Send to Back', exact: true }).hover()
      await page
        .locator('.chart-menu-child')
        .getByRole('button', { name: 'Send to Back', exact: true })
        .click()
      await expect.poll(top).toBe('Layer 0')
      await expect.poll(order).toEqual(['Layer 2', 'Layer 1', 'Layer 0'])
    } finally {
      await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
    }
  })
}

test('source shape and picture contextual ribbons use shared paint, transform and size controls', async ({
  page,
}) => {
  page.setDefaultTimeout(8000)
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(
      false,
      false,
      false,
      false,
      false,
      false,
      false,
      true,
    )
  })
  try {
    const frame = (id: string) => page.locator(`[data-print-visual="${id}"] .shape-editable`)
    const picture = page.locator('[data-print-visual="picture-0"] img')
    await expect(picture).toBeVisible()
    await expect
      .poll(() => picture.evaluate((img: HTMLImageElement) => img.complete && img.naturalWidth > 0))
      .toBe(true)
    expect(await page.evaluate(() => (window as any).toolbar.mediaReads())).toContain('picture-0')
    await expect(page.getByRole('tab', { name: 'Shape Format', exact: true })).toHaveCount(0)
    await frame('shape-0').click({ position: { x: 15, y: 15 } })
    await page.getByRole('tab', { name: 'Shape Format', exact: true }).click()
    const ribbon = page.locator('#ribbon-panel-shapeFormat')
    const read = () =>
      page.evaluate(
        () => (window as any).toolbar.snapshot().sheets[0].features.visualEdits?.['shape-0'],
      )
    await ribbon.getByRole('button', { name: 'Shape fill options', exact: true }).click()
    await page
      .getByRole('grid', { name: 'Standard Colors', exact: true })
      .getByRole('gridcell', { name: 'Red', exact: true })
      .click()
    await expect.poll(read).toMatchObject({ fillColor: '#ff0000' })
    await ribbon.getByRole('button', { name: 'Shape outline options', exact: true }).click()
    await page.getByRole('button', { name: 'No Outline', exact: true }).click()
    await expect.poll(read).toMatchObject({ lineColor: 'none' })
    const before = await page.evaluate(() => (window as any).toolbar.revision())
    await ribbon.getByRole('button', { name: 'Rotate', exact: true }).click()
    await ribbon.getByRole('option', { name: 'Rotate Right 90°', exact: true }).click()
    await expect.poll(read).toMatchObject({ rotation: 90 })
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(before + 1)
    await ribbon.getByRole('button', { name: 'Rotate', exact: true }).click()
    await ribbon.getByRole('option', { name: 'Flip Horizontal', exact: true }).click()
    await expect.poll(read).toMatchObject({ flipH: true })
    await ribbon.getByRole('button', { name: 'Size and Properties…', exact: true }).click()
    const dialog = page.getByRole('dialog', { name: 'Size and Properties', exact: true })
    await dialog.getByRole('spinbutton', { name: /Width/ }).fill('3.2')
    await dialog.getByRole('spinbutton', { name: /Height/ }).fill('2.1')
    await dialog.getByRole('spinbutton', { name: /Rotation/ }).fill('35')
    await page.evaluate(() => (window as any).toolbar.rejectNextSubmission('Size write refused'))
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog.getByRole('alert')).toHaveText('Size write refused')
    await expect(dialog.getByRole('spinbutton', { name: /Width/ })).toHaveValue('3.2')
    const sizeBefore = await page.evaluate(() => (window as any).toolbar.revision())
    await page.evaluate(() => (window as any).toolbar.holdSubmissions())
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog.getByRole('spinbutton', { name: /Width/ })).toBeDisabled()
    await expect(dialog.getByRole('button', { name: 'Cancel', exact: true })).toBeDisabled()
    await page.keyboard.press('Escape')
    await expect(dialog).toBeVisible()
    await page.evaluate(() => {
      ;(window as any).sourceSizeFlush = 'waiting'
      void (window as any).toolbar.flush().then((ok: boolean) => {
        ;(window as any).sourceSizeFlush = ok ? 'ready' : 'failed'
      })
    })
    expect(await page.evaluate(() => (window as any).sourceSizeFlush)).toBe('waiting')
    await page.evaluate(() => (window as any).toolbar.releaseSubmissions())
    await expect(dialog).toHaveCount(0)
    await expect.poll(() => page.evaluate(() => (window as any).sourceSizeFlush)).toBe('ready')
    await expect.poll(read).toMatchObject({
      fillColor: '#ff0000',
      lineColor: 'none',
      rotation: 35,
      flipH: true,
      frameSize: { width: 1152000, height: 756000 },
    })
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(sizeBefore + 1)
    await frame('picture-0').click({ position: { x: 15, y: 15 } })
    await expect(page.getByRole('tab', { name: 'Shape Format', exact: true })).toHaveCount(0)
    await page.getByRole('tab', { name: 'Picture Format', exact: true }).click()
    const pictureRibbon = page.locator('#ribbon-panel-pictureFormat')
    await expect(
      pictureRibbon.getByRole('button', { name: 'Shape fill', exact: true }),
    ).toHaveCount(0)
    await pictureRibbon.getByRole('button', { name: 'Rotate', exact: true }).click()
    await pictureRibbon.getByRole('option', { name: 'Flip Vertical', exact: true }).click()
    await expect
      .poll(() =>
        page.evaluate(
          () =>
            (window as any).toolbar.snapshot().sheets[0].features.visualEdits['picture-0'].flipV,
        ),
      )
      .toBe(true)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
  }
})

test('source shape multi-selection aligns and distributes as atomic edits using actual desktop geometry', async ({
  page,
}) => {
  page.setDefaultTimeout(8000)
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(
      false,
      false,
      false,
      false,
      false,
      false,
      false,
      true,
    )
  })
  try {
    const frame = (i: number) => page.locator(`[data-print-visual="shape-${i}"] .shape-editable`)
    await frame(0).click({ position: { x: 15, y: 15 } })
    await frame(1).click({ position: { x: 15, y: 15 }, modifiers: ['Shift'] })
    await page.getByRole('tab', { name: 'Shape Format', exact: true }).click()
    const ribbon = page.locator('#ribbon-panel-shapeFormat')
    const before = await page.evaluate(() => (window as any).toolbar.revision())
    await ribbon.getByRole('button', { name: 'Align', exact: true }).click()
    await ribbon.getByRole('option', { name: 'Align Top', exact: true }).click()
    await expect
      .poll(() => page.evaluate(() => (window as any).toolbar.revision()))
      .toBe(before + 1)
    expect(
      Math.abs((await frame(0).boundingBox())!.y - (await frame(1).boundingBox())!.y),
    ).toBeLessThan(1)
    await frame(2).click({ position: { x: 15, y: 15 }, modifiers: ['Shift'] })
    await ribbon.getByRole('button', { name: 'Align', exact: true }).click()
    await ribbon.getByRole('option', { name: 'Distribute Horizontally', exact: true }).click()
    await expect
      .poll(() => page.evaluate(() => (window as any).toolbar.revision()))
      .toBe(before + 2)
    const boxes = await Promise.all([0, 1, 2].map((i) => frame(i).boundingBox()))
    const sorted = boxes.map((box) => box!).sort((a, b) => a.x - b.x)
    expect(
      Math.abs(
        sorted[1]!.x -
          sorted[0]!.x -
          sorted[0]!.width -
          (sorted[2]!.x - sorted[1]!.x - sorted[1]!.width),
      ),
    ).toBeLessThan(1)
    await ribbon.getByRole('button', { name: 'Align', exact: true }).click()
    await ribbon.getByRole('option', { name: 'Align Left', exact: true }).click()
    await expect
      .poll(() => page.evaluate(() => (window as any).toolbar.revision()))
      .toBe(before + 3)
    await ribbon.getByRole('button', { name: 'Align', exact: true }).click()
    await ribbon.getByRole('option', { name: 'Distribute Vertically', exact: true }).click()
    await expect
      .poll(() => page.evaluate(() => (window as any).toolbar.revision()))
      .toBe(before + 4)
    const vertical = (await Promise.all([0, 1, 2].map((i) => frame(i).boundingBox())))
      .map((box) => box!)
      .sort((a, b) => a.y - b.y)
    expect(
      Math.abs(
        vertical[1]!.y -
          vertical[0]!.y -
          vertical[0]!.height -
          (vertical[2]!.y - vertical[1]!.y - vertical[1]!.height),
      ),
    ).toBeLessThan(1)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
  }
})

test('shared shape gallery draws a session shape and commits its focused text through save flushing', async ({
  page,
}) => {
  page.setDefaultTimeout(8000)
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar()
  })
  try {
    await page.getByRole('tab', { name: 'Insert', exact: true }).click()
    const header = page.locator('.workbook-desktop-header')
    await header.getByRole('button', { name: 'Shapes', exact: true }).click()
    await page
      .locator('.rb-shape-gallery')
      .getByRole('button', { name: 'Rectangle', exact: true })
      .click()
    const grid = page.locator('.workbook-grid')
    await expect(grid).toHaveClass(/sheet-shape-drawing/)
    const surface = await grid.locator('canvas').evaluateAll(
      (nodes) =>
        nodes
          .map((node) => {
            const r = node.getBoundingClientRect()
            return { x: r.x, y: r.y, width: r.width, height: r.height }
          })
          .sort((a, b) => b.width * b.height - a.width * a.height)[0]!,
    )
    const start = { x: surface.x + 230, y: surface.y + 150 }
    await page.mouse.move(start.x, start.y)
    await page.mouse.down()
    await page.mouse.move(start.x + 150, start.y + 90, { steps: 6 })
    await page.mouse.up()
    const frame = page.locator('[data-print-visual^="added-shape-"] .shape-editable')
    await expect(frame).toBeVisible()
    const box = (await frame.boundingBox())!
    expect(Math.abs(box.width - 150)).toBeLessThan(2)
    expect(Math.abs(box.height - 90)).toBeLessThan(2)
    const shape = await page.evaluate(
      () => (window as any).toolbar.snapshot().sheets[0].features.visuals[0],
    )
    expect(shape).toMatchObject({ kind: 'shape', shapeType: 'rect', fillColor: '#DDEBF7' })
    await frame.dblclick({ position: { x: 30, y: 25 } })
    const editor = frame.locator('.shape-text-editor')
    await editor.fill('Keep rejected text')
    await page.evaluate(() => (window as any).toolbar.rejectNextSubmission('Shape text refused'))
    await editor.press('Enter')
    await expect(frame.getByRole('alert')).toHaveText('Shape text refused')
    await expect(editor).toHaveText('Keep rejected text')
    await editor.fill('Final shape text')
    await page.evaluate(() => (window as any).toolbar.holdSubmissions())
    await editor.press('Enter')
    await expect(editor).toHaveAttribute('contenteditable', 'false')
    await page.evaluate(() => {
      ;(window as any).shapeFlush = 'waiting'
      void (window as any).toolbar.flush().then((ok: boolean) => {
        ;(window as any).shapeFlush = ok ? 'ready' : 'failed'
      })
    })
    expect(await page.evaluate(() => (window as any).shapeFlush)).toBe('waiting')
    await page.evaluate(() => (window as any).toolbar.releaseSubmissions())
    await expect.poll(() => page.evaluate(() => (window as any).shapeFlush)).toBe('ready')
    await expect(editor).toHaveCount(0)
    await expect(frame).toContainText('Final shape text')
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(2)
    await header.getByRole('button', { name: 'Shapes', exact: true }).click()
    await page
      .locator('.rb-shape-gallery')
      .getByRole('button', { name: 'Oval', exact: true })
      .click()
    await expect(grid).toHaveClass(/sheet-shape-drawing/)
    await page.keyboard.press('Escape')
    await expect(grid).not.toHaveClass(/sheet-shape-drawing/)
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(2)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
  }
})

test('shared shape draw uses the default click size and Shift square while text boxes use desktop cell defaults', async ({
  page,
}) => {
  page.setDefaultTimeout(8000)
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar()
  })
  try {
    await page.getByRole('tab', { name: 'Insert', exact: true }).click()
    const header = page.locator('.workbook-desktop-header'),
      grid = page.locator('.workbook-grid')
    const arm = async () => {
      await header.getByRole('button', { name: 'Shapes', exact: true }).click()
      await page
        .locator('.rb-shape-gallery')
        .getByRole('button', { name: 'Oval', exact: true })
        .click()
      await expect(grid).toHaveClass(/sheet-shape-drawing/)
    }
    const surface = await grid.locator('canvas').evaluateAll(
      (nodes) =>
        nodes
          .map((n) => {
            const r = n.getBoundingClientRect()
            return { x: r.x, y: r.y, w: r.width, h: r.height }
          })
          .sort((a, b) => b.w * b.h - a.w * a.h)[0]!,
    )
    await arm()
    await page.mouse.click(surface.x + 200, surface.y + 150)
    const frames = page.locator('[data-print-visual^="added-shape-"] .shape-editable')
    await expect(frames).toHaveCount(1)
    const box = (await frames.first().boundingBox())!
    expect(Math.abs(box.width - 96)).toBeLessThan(1)
    expect(Math.abs(box.height - 96)).toBeLessThan(1)
    const firstAnchor = await page.evaluate(
      () => (window as any).toolbar.snapshot().sheets[0].features.visuals[0].anchor,
    )
    await arm()
    await page.keyboard.down('Shift')
    await page.mouse.move(surface.x + 220, surface.y + 170)
    await page.mouse.down()
    await page.mouse.move(surface.x + 350, surface.y + 210, { steps: 5 })
    await page.mouse.up()
    await page.keyboard.up('Shift')
    await expect(frames).toHaveCount(2)
    expect(
      await page.evaluate(
        () => (window as any).toolbar.snapshot().sheets[0].features.visuals[0].anchor,
      ),
    ).toEqual(firstAnchor)
    const square = (await frames.nth(1).boundingBox())!
    expect(Math.abs(square.width - square.height)).toBeLessThan(1)
    expect(Math.abs(square.width - 130)).toBeLessThan(2)
    await header.getByRole('button', { name: 'Text Box', exact: true }).click()
    await expect(frames).toHaveCount(3)
    expect(
      await page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].features.visuals[2]),
    ).toMatchObject({ kind: 'shape', isTextBox: true, text: 'Text', fillColor: '#FFFFFF' })
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(3)
    await arm()
    await page.evaluate(() => window.dispatchEvent(new Event('blur')))
    await expect(grid).not.toHaveClass(/sheet-shape-drawing/)
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(3)
    await arm()
    await page.evaluate(() => (window as any).toolbar.seedClearTargets())
    const advanced = await page.evaluate(() => (window as any).toolbar.revision())
    await page.mouse.click(surface.x + 650, surface.y + 170)
    await expect(page.locator('.workbook-error')).toContainText('The document changed')
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(advanced)
    expect(
      await page.evaluate(
        () => (window as any).toolbar.snapshot().sheets[0].features.visuals.length,
      ),
    ).toBe(3)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
  }
})

test('shared picture picker validates files, awaits commit and preserves the frozen target', async ({
  page,
}) => {
  page.setDefaultTimeout(8000)
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar()
  })
  const pixel = Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
    'base64',
  )
  try {
    await page.getByRole('tab', { name: 'Insert', exact: true }).click()
    const choose = async (name: string, mimeType: string, buffer: Buffer) => {
      const picker = page.waitForEvent('filechooser')
      await page
        .locator('.workbook-desktop-header')
        .getByRole('button', { name: 'Pictures', exact: true })
        .click()
      const chooser = await picker.catch(async (error) => {
        const diagnostics = await page.evaluate(() => (window as any).toolbar.diagnostics())
        throw new Error(`${String(error)}; ${JSON.stringify(diagnostics)}`, { cause: error })
      })
      expect(await chooser.element().getAttribute('accept')).toBe('image/png,image/jpeg,image/gif')
      await chooser.setFiles({ name, mimeType, buffer })
    }
    await choose('wrong.svg', 'image/svg+xml', Buffer.from('<svg/>'))
    await expect(page.locator('.workbook-error')).toContainText(/PNG|JPEG|GIF/)
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(0)
    const cancelled = page.waitForEvent('filechooser')
    await page
      .locator('.workbook-desktop-header')
      .getByRole('button', { name: 'Pictures', exact: true })
      .click()
    await (await cancelled).setFiles([])
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(0)
    await page.evaluate(() => (window as any).toolbar.holdSubmissions())
    await choose('pixel.png', 'image/png', pixel)
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.submissions())).toBe(1)
    await expect(page.locator('[data-print-visual^="added-image-"]')).toHaveCount(0)
    await page.evaluate(() => {
      ;(window as any).pictureFlush = 'waiting'
      void (window as any).toolbar.flush().then((ok: boolean) => {
        ;(window as any).pictureFlush = ok ? 'ready' : 'failed'
      })
    })
    expect(await page.evaluate(() => (window as any).pictureFlush)).toBe('waiting')
    await page.evaluate(() => (window as any).toolbar.releaseSubmissions())
    await expect.poll(() => page.evaluate(() => (window as any).pictureFlush)).toBe('ready')
    const image = page.locator('[data-print-visual^="added-image-"] img')
    await expect(image).toBeVisible()
    await expect
      .poll(() => image.evaluate((el: HTMLImageElement) => el.complete && el.naturalWidth === 1))
      .toBe(true)
    const read = () =>
      page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].features.visuals[0])
    const added = await read()
    expect(added.anchor.toRow - added.anchor.fromRow).toBe(2)
    expect(added.anchor.toColumn - added.anchor.fromColumn).toBe(2)
    expect(added.mediaDataUrl).toBe(`data:image/png;base64,${pixel.toString('base64')}`)
    await page
      .locator('[data-print-visual^="added-image-"] .shape-editable')
      .click({ position: { x: 15, y: 15 } })
    await page.getByRole('tab', { name: 'Picture Format', exact: true }).click()
    const ribbon = page.locator('#ribbon-panel-pictureFormat')
    await ribbon.getByRole('button', { name: 'Rotate', exact: true }).click()
    await ribbon.getByRole('option', { name: 'Flip Horizontal', exact: true }).click()
    await expect.poll(read).toMatchObject({ flipH: true })
    await page.getByRole('tab', { name: 'Insert', exact: true }).click()
    const stalePicker = page.waitForEvent('filechooser')
    await page
      .locator('.workbook-desktop-header')
      .getByRole('button', { name: 'Pictures', exact: true })
      .click()
    const chooser = await stalePicker
    await page.evaluate(() => (window as any).toolbar.seedClearTargets())
    const revision = await page.evaluate(() => (window as any).toolbar.revision())
    await chooser.setFiles({ name: 'stale.png', mimeType: 'image/png', buffer: pixel })
    await expect(page.locator('.workbook-error')).toContainText('The document changed')
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(revision)
    expect(
      await page.evaluate(
        () => (window as any).toolbar.snapshot().sheets[0].features.visuals.length,
      ),
    ).toBe(1)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
  }
})

test('shared picture transfers paste at selection and drop at the actual grid cell through Save flushing', async ({
  page,
}) => {
  page.setDefaultTimeout(8000)
  await page.context().grantPermissions(['clipboard-read', 'clipboard-write'])
  await page.goto(base)
  const pixel =
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII='
  await page.evaluate(async (base64) => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar()
    ;(window as any).toolbar.select('B2')
    const bytes = Uint8Array.from(atob(base64), (c) => c.charCodeAt(0))
    await navigator.clipboard.write([
      new ClipboardItem({ 'image/png': new Blob([bytes], { type: 'image/png' }) }),
    ])
  }, pixel)
  try {
    const grid = page.locator('.workbook-grid')
    const point = await grid.locator('canvas').evaluateAll((nodes) => {
      const canvas = (nodes as HTMLCanvasElement[]).sort(
        (a, b) => b.width * b.height - a.width * a.height,
      )[0]!
      const r = canvas.getBoundingClientRect(),
        cell = (window as any).toolbar.cellRect('B2')
      return { x: r.x + (cell.left + cell.right) / 2, y: r.y + (cell.top + cell.bottom) / 2 }
    })
    await page.mouse.click(point.x, point.y)
    await page.evaluate(() => (window as any).toolbar.holdSubmissions())
    await page.keyboard.press('Control+V')
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.submissions())).toBe(1)
    await page.evaluate(() => {
      ;(window as any).transferFlush = 'waiting'
      void (window as any).toolbar.flush().then((ok: boolean) => {
        ;(window as any).transferFlush = ok ? 'ready' : 'failed'
      })
    })
    expect(await page.evaluate(() => (window as any).transferFlush)).toBe('waiting')
    await page.evaluate(() => (window as any).toolbar.releaseSubmissions())
    await expect.poll(() => page.evaluate(() => (window as any).transferFlush)).toBe('ready')
    const read = () =>
      page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].features.visuals)
    await expect.poll(read).toHaveLength(1)
    expect((await read())[0].anchor).toMatchObject({ fromRow: 1, fromColumn: 1 })
    const image = page.locator('[data-print-visual^="added-image-"] img')
    await expect
      .poll(() =>
        image.first().evaluate((el: HTMLImageElement) => el.complete && el.naturalWidth === 1),
      )
      .toBe(true)
    await page.evaluate((base64) => {
      const grid = document.querySelector('.workbook-grid')!
      const canvas = [...grid.querySelectorAll('canvas')].sort(
        (a, b) => b.width * b.height - a.width * a.height,
      )[0]!
      const bounds = canvas.getBoundingClientRect(),
        cell = (window as any).toolbar.cellRect('D5')
      const data = new DataTransfer()
      data.items.add(
        new File([Uint8Array.from(atob(base64), (c) => c.charCodeAt(0))], 'dropped.png', {
          type: 'image/png',
        }),
      )
      canvas.dispatchEvent(
        new DragEvent('drop', {
          bubbles: true,
          cancelable: true,
          dataTransfer: data,
          clientX: bounds.x + (cell.left + cell.right) / 2,
          clientY: bounds.y + (cell.top + cell.bottom) / 2,
        }),
      )
    }, pixel)
    await expect.poll(read).toHaveLength(2)
    const visuals = await read()
    expect(visuals[1].anchor).toMatchObject({ fromRow: 4, fromColumn: 3 })
    expect(visuals[1].mediaDataUrl).toBe(`data:image/png;base64,${pixel}`)
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(2)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
  }
})

test('shared picture transfers leave table and text pastes to their editors and reject overtaken drops', async ({
  page,
}) => {
  page.setDefaultTimeout(8000)
  await page.context().grantPermissions(['clipboard-read', 'clipboard-write'])
  await page.goto(base)
  const pixel =
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII='
  await page.evaluate(async (base64) => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar()
    ;(window as any).toolbar.select('B2')
    const bytes = Uint8Array.from(atob(base64), (c) => c.charCodeAt(0))
    await navigator.clipboard.write([
      new ClipboardItem({
        'image/png': new Blob([bytes], { type: 'image/png' }),
        'text/html': new Blob(['<table><tr><td>77</td></tr></table>'], { type: 'text/html' }),
        'text/plain': new Blob(['77'], { type: 'text/plain' }),
      }),
    ])
  }, pixel)
  try {
    const point = await page.locator('.workbook-grid canvas').evaluateAll((nodes) => {
      const canvas = (nodes as HTMLCanvasElement[]).sort(
        (a, b) => b.width * b.height - a.width * a.height,
      )[0]!
      const r = canvas.getBoundingClientRect(),
        cell = (window as any).toolbar.cellRect('B2')
      return { x: r.x + (cell.left + cell.right) / 2, y: r.y + (cell.top + cell.bottom) / 2 }
    })
    await page.mouse.click(point.x, point.y)
    await page.keyboard.press('Control+V')
    await expect
      .poll(() =>
        page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells.B2.rawValue),
      )
      .toBe(77)
    expect(
      await page.evaluate(
        () => (window as any).toolbar.snapshot().sheets[0].features.visuals ?? [],
      ),
    ).toEqual([])
    const untouched = await page.evaluate(async (base64) => {
      const makeData = () => {
        const data = new DataTransfer()
        data.items.add(
          new File([Uint8Array.from(atob(base64), (c) => c.charCodeAt(0))], 'paste.png', {
            type: 'image/png',
          }),
        )
        return data
      }
      const input = document.createElement('input')
      document.querySelector('.workbook-grid')!.append(input)
      input.focus()
      const event = new ClipboardEvent('paste', {
        bubbles: true,
        cancelable: true,
        clipboardData: makeData(),
      })
      input.dispatchEvent(event)
      input.remove()
      await (window as any).toolbar.flush()
      return {
        prevented: event.defaultPrevented,
        visuals: (window as any).toolbar.snapshot().sheets[0].features.visuals ?? [],
      }
    }, pixel)
    expect(untouched).toEqual({ prevented: false, visuals: [] })
    const revision = await page.evaluate(async (base64) => {
      const grid = document.querySelector('.workbook-grid')!,
        canvas = [...grid.querySelectorAll('canvas')].sort(
          (a, b) => b.width * b.height - a.width * a.height,
        )[0]!
      const r = canvas.getBoundingClientRect(),
        data = new DataTransfer()
      data.items.add(
        new File([Uint8Array.from(atob(base64), (c) => c.charCodeAt(0))], 'stale.png', {
          type: 'image/png',
        }),
      )
      canvas.dispatchEvent(
        new DragEvent('drop', {
          bubbles: true,
          cancelable: true,
          dataTransfer: data,
          clientX: r.x + 350,
          clientY: r.y + 140,
        }),
      )
      await (window as any).toolbar.seedClearTargets()
      return (window as any).toolbar.revision()
    }, pixel)
    await expect(page.locator('.workbook-error')).toContainText('The document changed')
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(revision)
    expect(
      await page.evaluate(
        () => (window as any).toolbar.snapshot().sheets[0].features.visuals ?? [],
      ),
    ).toEqual([])
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
  }
})

test('desktop parsed external cell paste commits values, styles and merges in one authoritative revision', async ({
  page,
}) => {
  page.setDefaultTimeout(8000)
  await page.context().grantPermissions(['clipboard-read', 'clipboard-write'])
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar()
    ;(window as any).toolbar.select('D5')
    await navigator.clipboard.write([
      new ClipboardItem({
        'text/html': new Blob(
          [
            '<table><tr><td rowspan="2" style="font-weight:bold;background-color:#ffeecc;color:#112233">Merged</td><td>12.5</td></tr><tr><td style="font-style:italic">Second</td></tr></table>',
          ],
          { type: 'text/html' },
        ),
        'text/plain': new Blob(['Merged\t12.5\n\tSecond'], { type: 'text/plain' }),
      }),
    ])
  })
  try {
    const cellPoint = () =>
      page.locator('.workbook-grid canvas').evaluateAll((nodes) => {
        const canvas = (nodes as HTMLCanvasElement[]).sort(
            (a, b) => b.width * b.height - a.width * a.height,
          )[0]!,
          r = canvas.getBoundingClientRect(),
          cell = (window as any).toolbar.cellRect('D5')
        return { x: r.x + (cell.left + cell.right) / 2, y: r.y + (cell.top + cell.bottom) / 2 }
      })
    const point = await cellPoint()
    await page.mouse.click(point.x, point.y)
    await page.evaluate(() =>
      (window as any).toolbar.rejectNextSubmission('Clipboard write refused'),
    )
    await page.keyboard.press('Control+V')
    await expect(page.locator('.workbook-error')).toContainText('Clipboard write refused')
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(0)
    expect(await page.evaluate(() => (window as any).toolbar.value('D5'))).toBeNull()
    const retryPoint = await cellPoint()
    await page.mouse.click(retryPoint.x, retryPoint.y)
    await page.evaluate(() => (window as any).toolbar.holdSubmissions())
    await page.keyboard.press('Control+V')
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.submissions())).toBe(2)
    expect(
      await page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells.D5),
    ).toBeUndefined()
    expect(await page.evaluate(() => (window as any).toolbar.value('D5'))).toBeNull()
    await page.evaluate(() => {
      ;(window as any).cellPasteFlush = 'waiting'
      void (window as any).toolbar.flush().then((ok: boolean) => {
        ;(window as any).cellPasteFlush = ok ? 'ready' : 'failed'
      })
    })
    expect(await page.evaluate(() => (window as any).cellPasteFlush)).toBe('waiting')
    await page.evaluate(() => (window as any).toolbar.releaseSubmissions())
    await expect.poll(() => page.evaluate(() => (window as any).cellPasteFlush)).toBe('ready')
    const sheet = await page.evaluate(() => (window as any).toolbar.snapshot().sheets[0])
    expect(sheet.cells.D5).toMatchObject({
      kind: 'string',
      rawValue: 'Merged',
      style: { bold: true, fillColor: '#FFEECC', fontColor: '#112233' },
    })
    expect(sheet.cells.E5).toMatchObject({ kind: 'number', rawValue: 12.5 })
    expect(sheet.cells.E6).toMatchObject({
      kind: 'string',
      rawValue: 'Second',
      style: { italic: true },
    })
    expect(sheet.features.mergeEdits).toContainEqual({
      startRow: 4,
      endRow: 5,
      startColumn: 3,
      endColumn: 3,
    })
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(1)
    expect(await page.evaluate(() => (window as any).toolbar.merges())).toContainEqual(
      expect.objectContaining({ startRow: 4, endRow: 5, startColumn: 3, endColumn: 3 }),
    )
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
  }
})

test('desktop Paste button shares image and formatted table transactions and preserves native paste when clipboard reads are denied', async ({
  page,
}) => {
  page.setDefaultTimeout(8000)
  await page.context().grantPermissions(['clipboard-read', 'clipboard-write'])
  await page.goto(base)
  const pixel =
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII='
  await page.evaluate(async (base64) => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar()
    ;(window as any).toolbar.select('B2')
    await navigator.clipboard.write([
      new ClipboardItem({
        'image/png': new Blob([Uint8Array.from(atob(base64), (c) => c.charCodeAt(0))], {
          type: 'image/png',
        }),
      }),
    ])
    const stored = (await navigator.clipboard.read())[0]!
    const bytes = new Uint8Array(await (await stored.getType('image/png')).arrayBuffer())
    ;(window as any).expectedClipboardImage =
      `data:image/png;base64,${btoa(String.fromCharCode(...bytes))}`
  }, pixel)
  try {
    const paste = page
      .locator('.workbook-desktop-header')
      .getByRole('button', { name: 'Paste', exact: true })
    await page.evaluate(() => (window as any).toolbar.holdSubmissions())
    await paste.click()
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.submissions())).toBe(1)
    await page.evaluate(() => {
      ;(window as any).menuPasteFlush = 'waiting'
      void (window as any).toolbar.flush().then((ok: boolean) => {
        ;(window as any).menuPasteFlush = ok ? 'ready' : 'failed'
      })
    })
    expect(await page.evaluate(() => (window as any).menuPasteFlush)).toBe('waiting')
    await page.evaluate(() => (window as any).toolbar.releaseSubmissions())
    await expect.poll(() => page.evaluate(() => (window as any).menuPasteFlush)).toBe('ready')
    const image = await page.evaluate(
      () => (window as any).toolbar.snapshot().sheets[0].features.visuals[0],
    )
    expect(image).toMatchObject({ kind: 'image', anchor: { fromRow: 1, fromColumn: 1 } })
    expect(image.mediaDataUrl).toBe(
      await page.evaluate(() => (window as any).expectedClipboardImage),
    )
    await page.evaluate(async (base64) => {
      await navigator.clipboard.write([
        new ClipboardItem({
          'image/png': new Blob([Uint8Array.from(atob(base64), (c) => c.charCodeAt(0))], {
            type: 'image/png',
          }),
          'text/html': new Blob(
            [
              '<table><tr><td style="font-weight:bold;background-color:#ffeecc">Menu table</td></tr></table>',
            ],
            { type: 'text/html' },
          ),
          'text/plain': new Blob(['Menu table'], { type: 'text/plain' }),
        }),
      ])
    }, pixel)
    await paste.click()
    await expect
      .poll(() => page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells.B2))
      .toMatchObject({ rawValue: 'Menu table', style: { bold: true, fillColor: '#FFEECC' } })
    expect(
      await page.evaluate(
        () => (window as any).toolbar.snapshot().sheets[0].features.visuals.length,
      ),
    ).toBe(1)
    const before = await page.evaluate(() => (window as any).toolbar.revision())
    await page.evaluate(async () => {
      await navigator.clipboard.writeText('Recovery')
      Object.defineProperty(navigator, 'clipboard', {
        configurable: true,
        value: { read: () => Promise.reject(new DOMException('denied', 'NotAllowedError')) },
      })
    })
    await paste.click()
    await expect(page.locator('.workbook-error')).toContainText('Unable to read the clipboard')
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(before)
    await page.locator('.workbook-grid').evaluate((el) => {
      el.tabIndex = 0
      el.focus()
    })
    await page.keyboard.press('Control+V')
    await expect
      .poll(() =>
        page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells.B2.rawValue),
      )
      .toBe('Recovery')
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
  }
})

test('desktop Paste button freezes the target and makes Save await a delayed clipboard read', async ({
  page,
}) => {
  page.setDefaultTimeout(8000)
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar()
    ;(window as any).toolbar.select('B2')
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: {
        read: () =>
          new Promise((resolve) => {
            ;(window as any).releaseClipboardRead = () =>
              resolve([
                new ClipboardItem({
                  'text/plain': new Blob(['Stale clipboard'], { type: 'text/plain' }),
                }),
              ])
          }),
      },
    })
  })
  try {
    await page
      .locator('.workbook-desktop-header')
      .getByRole('button', { name: 'Paste', exact: true })
      .click()
    await expect
      .poll(() => page.evaluate(() => typeof (window as any).releaseClipboardRead))
      .toBe('function')
    await page.evaluate(() => {
      ;(window as any).clipboardReadFlush = 'waiting'
      void (window as any).toolbar.flush().then((ok: boolean) => {
        ;(window as any).clipboardReadFlush = ok ? 'ready' : 'failed'
      })
    })
    expect(await page.evaluate(() => (window as any).clipboardReadFlush)).toBe('waiting')
    await page.evaluate(() => (window as any).toolbar.seedClearTargets())
    const revision = await page.evaluate(() => (window as any).toolbar.revision())
    await page.evaluate(() => (window as any).releaseClipboardRead())
    await expect(page.locator('.workbook-error')).toContainText('The document changed')
    await expect.poll(() => page.evaluate(() => (window as any).clipboardReadFlush)).toBe('failed')
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(revision)
    expect(
      await page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells.B2.rawValue),
    ).toBe(120)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
  }
})

test('desktop Copy reads a complete typed source and keeps its clipboard snapshot across source edits', async ({
  page,
}) => {
  page.setDefaultTimeout(8000)
  await page.context().grantPermissions(['clipboard-read', 'clipboard-write'])
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar()
    await (window as any).toolbar.seedClipboardSource()
    ;(window as any).toolbar.select('A80:B81')
  })
  try {
    const header = page.locator('.workbook-desktop-header')
    const initial = await page.evaluate(() => (window as any).toolbar.revision())
    await header.getByRole('button', { name: /^Copy / }).click()
    await expect
      .poll(() => page.evaluate(async () => await navigator.clipboard.readText()))
      .toContain('=literal')
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(initial)
    expect(await page.evaluate(() => (window as any).toolbar.reads())).toContainEqual(
      expect.objectContaining({ startRow: 79, endRow: 80, startColumn: 0, endColumn: 1 }),
    )
    await page.evaluate(() => (window as any).toolbar.changeCopiedSource())
    await page.evaluate(() => (window as any).toolbar.select('D5'))
    await header.getByRole('button', { name: 'Paste', exact: true }).click()
    await expect
      .poll(() => page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells.D5))
      .toMatchObject({
        kind: 'string',
        rawValue: '=literal',
        style: {
          bold: true,
          fillColor: '#FFEECC',
          borderLeft: { style: 'mediumDashDotDot', color: '#123456' },
        },
      })
    const cells = await page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells)
    expect(cells.A80.rawValue).toBe('New source')
    expect(cells.D5.formula).toBeUndefined()
    expect(cells.E5).toMatchObject({ formula: '=D5&$A$1', calculationState: 'pending' })
    expect(cells.E5.cachedValue).toBeUndefined()
    expect(cells.D6).toMatchObject({
      kind: 'date',
      rawValue: 45000,
      dateSystem: '1900',
      numberFormat: 'yyyy-mm-dd',
    })
    expect(cells.E6).toMatchObject({ kind: 'error', rawValue: '#DIV/0!' })
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
  }
})

test('native grid Copy and Paste preserve merged cells and tile a copied block to a repeated target', async ({
  page,
}) => {
  page.setDefaultTimeout(8000)
  await page.context().grantPermissions(['clipboard-read', 'clipboard-write'])
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(false, true)
    ;(window as any).toolbar.select('A1:B1')
  })
  try {
    await page.locator('.workbook-grid').evaluate((el) => {
      el.tabIndex = 0
      el.focus()
    })
    await page.keyboard.press('Control+C')
    await expect
      .poll(() => page.evaluate(async () => await navigator.clipboard.readText()))
      .toContain('Region')
    await page.evaluate(() => (window as any).toolbar.select('D5:G6'))
    await page.locator('.workbook-grid').evaluate((el) => el.focus())
    await page.keyboard.press('Control+V')
    await expect
      .poll(() =>
        page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].features.mergeEdits),
      )
      .toEqual(
        expect.arrayContaining([
          { startRow: 4, endRow: 4, startColumn: 3, endColumn: 4 },
          { startRow: 4, endRow: 4, startColumn: 5, endColumn: 6 },
          { startRow: 5, endRow: 5, startColumn: 3, endColumn: 4 },
          { startRow: 5, endRow: 5, startColumn: 5, endColumn: 6 },
        ]),
      )
    const cells = await page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells)
    for (const address of ['D5', 'F5', 'D6', 'F6']) expect(cells[address].rawValue).toBe('Region')
    for (const address of ['E5', 'G5', 'E6', 'G6']) expect(cells[address].rawValue).toBeNull()
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(1)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
  }
})

test('desktop Cut moves an overlapping source as one transaction and updates distant formula dependents', async ({
  page,
}) => {
  page.setDefaultTimeout(8000)
  await page.context().grantPermissions(['clipboard-read', 'clipboard-write'])
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar()
    await (window as any).toolbar.seedCutSource()
    ;(window as any).toolbar.select('A80:B80')
  })
  try {
    const header = page.locator('.workbook-desktop-header')
    const before = await page.evaluate(() => (window as any).toolbar.revision())
    await header.getByRole('button', { name: /^Cut / }).click()
    await expect
      .poll(() => page.evaluate(async () => navigator.clipboard.readText()))
      .toContain('12')
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(before)
    await page.evaluate(() => (window as any).toolbar.select('B80'))
    await header.getByRole('button', { name: 'Paste', exact: true }).click()
    await expect
      .poll(() => page.evaluate(() => (window as any).toolbar.revision()))
      .toBe(before + 1)
    const cells = await page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells)
    expect(cells.A80).toMatchObject({ rawValue: null, styleReset: true })
    expect(cells.B80).toMatchObject({ rawValue: 12, style: { bold: true } })
    expect(cells.C80).toMatchObject({ formula: '=B80+$A$1', calculationState: 'pending' })
    expect(cells.J90.formula).toBe('=$B$80')
    await header.getByRole('button', { name: 'Paste', exact: true }).click()
    await expect(page.locator('.workbook-error')).toContainText('clipboard')
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(before + 1)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
  }
})

test('grid Cut rejects a source changed after cutting without clearing source or destination', async ({
  page,
}) => {
  page.setDefaultTimeout(8000)
  await page.context().grantPermissions(['clipboard-read', 'clipboard-write'])
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar()
    await (window as any).toolbar.seedCutSource()
    ;(window as any).toolbar.select('A80:B80')
  })
  try {
    await page.locator('.workbook-grid').evaluate((el) => {
      el.tabIndex = 0
      el.focus()
    })
    await page.keyboard.press('Control+X')
    await expect
      .poll(() => page.evaluate(async () => navigator.clipboard.readText()))
      .toContain('12')
    await page.evaluate(async () => {
      await (window as any).toolbar.changeCopiedSource()
      ;(window as any).toolbar.select('D5')
    })
    const before = await page.evaluate(() => (window as any).toolbar.snapshot())
    await page.locator('.workbook-grid').evaluate((el) => el.focus())
    await page.keyboard.press('Control+V')
    await expect(page.locator('.workbook-error')).toContainText('changed')
    expect(await page.evaluate(() => (window as any).toolbar.snapshot())).toEqual(before)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
  }
})

test('Escape cancels a pending grid Cut without clearing cells and a new Copy remains usable', async ({
  page,
}) => {
  page.setDefaultTimeout(8000)
  await page.context().grantPermissions(['clipboard-read', 'clipboard-write'])
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar()
    await (window as any).toolbar.seedCutSource()
    ;(window as any).toolbar.select('A80:B80')
  })
  try {
    const grid = page.locator('.workbook-grid'),
      header = page.locator('.workbook-desktop-header')
    const before = await page.evaluate(() => (window as any).toolbar.snapshot())
    await grid.evaluate((el) => {
      el.tabIndex = 0
      el.focus()
    })
    await page.keyboard.press('Control+X')
    await expect
      .poll(() => page.evaluate(async () => navigator.clipboard.readText()))
      .toContain('12')
    await page.keyboard.press('Escape')
    await page.evaluate(() => (window as any).toolbar.select('D5'))
    await header.getByRole('button', { name: 'Paste', exact: true }).click()
    await expect(page.locator('.workbook-error')).toContainText('clipboard')
    expect(await page.evaluate(() => (window as any).toolbar.snapshot())).toEqual(before)
    await page.evaluate(() => (window as any).toolbar.select('A80:B80'))
    await header.getByRole('button', { name: /^Copy / }).click()
    await page.evaluate(() => (window as any).toolbar.select('D5'))
    await header.getByRole('button', { name: 'Paste', exact: true }).click()
    await expect
      .poll(() =>
        page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells.D5?.rawValue),
      )
      .toBe(12)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
  }
})

test('unsupported cut metadata shows a localized no-change message and retains the cut for retry', async ({
  page,
}) => {
  page.setDefaultTimeout(8000)
  await page.context().grantPermissions(['clipboard-read', 'clipboard-write'])
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar()
    await (window as any).toolbar.seedCutSource()
    ;(window as any).toolbar.select('A80:B80')
  })
  try {
    const header = page.locator('.workbook-desktop-header')
    await header.getByRole('button', { name: /^Cut / }).click()
    await expect
      .poll(() => page.evaluate(async () => navigator.clipboard.readText()))
      .toContain('12')
    await page.evaluate(() => {
      ;(window as any).toolbar.select('D5')
      ;(window as any).toolbar.rejectNextSubmission('unsupported_move_feature: definedNames')
    })
    const before = await page.evaluate(() => (window as any).toolbar.snapshot())
    await header.getByRole('button', { name: 'Paste', exact: true }).click()
    await expect(page.locator('.workbook-error')).toContainText('No cells were changed.')
    await expect(page.locator('.workbook-error')).not.toContainText('unsupported_move_feature')
    expect(await page.evaluate(() => (window as any).toolbar.snapshot())).toEqual(before)
    await header.getByRole('button', { name: 'Paste', exact: true }).click()
    await expect
      .poll(() =>
        page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells.D5?.rawValue),
      )
      .toBe(12)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
  }
})

async function mountSpecialPaste(page: import('@playwright/test').Page) {
  page.setDefaultTimeout(8000)
  await page.context().grantPermissions(['clipboard-read', 'clipboard-write'])
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar()
    await (window as any).toolbar.seedSpecialSource()
    ;(window as any).toolbar.select('A80:B81')
  })
  await page
    .locator('.workbook-desktop-header')
    .getByRole('button', { name: /^Copy / })
    .click()
  await expect.poll(() => page.evaluate(async () => navigator.clipboard.readText())).toContain('2')
  await page.evaluate(() => (window as any).toolbar.select('D5'))
}
async function specialChoice(page: import('@playwright/test').Page, name: string) {
  await page
    .locator('.workbook-desktop-header')
    .getByRole('button', { name: 'Paste Special', exact: true })
    .click()
  await page.getByRole('option', { name, exact: true }).click()
}

test('shared Paste Special quick values, formats and except-borders preserve unrelated destination fields', async ({
  page,
}) => {
  await mountSpecialPaste(page)
  try {
    await specialChoice(page, 'Values Only')
    await expect
      .poll(() =>
        page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells.D5.rawValue),
      )
      .toBe(2)
    let cells = await page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells)
    expect(cells.D5.style).toMatchObject({
      italic: true,
      fillColor: '#CCEEFF',
      borderLeft: { style: 'double' },
    })
    await specialChoice(page, 'All Except Borders')
    await expect
      .poll(() =>
        page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells.D5.style?.bold),
      )
      .toBe(true)
    cells = await page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells)
    expect(cells.D5.style).toMatchObject({
      bold: true,
      fillColor: '#FFDD99',
      borderLeft: { style: 'double', color: '#123456' },
      borderRight: null,
    })
    await specialChoice(page, 'Formatting Only')
    await expect
      .poll(() =>
        page.evaluate(
          () => (window as any).toolbar.snapshot().sheets[0].cells.D5.style?.borderRight?.style,
        ),
      )
      .toBe('thin')
    expect(
      await page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells.D5.rawValue),
    ).toBe(2)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
  }
})

test('shared Paste Special dialog matches desktop defaults and commits arithmetic, transpose and skip blanks once', async ({
  page,
}) => {
  await mountSpecialPaste(page)
  try {
    const before = await page.evaluate(() => (window as any).toolbar.revision())
    await specialChoice(page, 'Paste Special…')
    const dialog = page.getByRole('dialog', { name: 'Paste Special', exact: true })
    await expect(dialog.getByRole('radio', { name: 'All', exact: true })).toBeChecked()
    await expect(dialog.getByRole('radio', { name: 'None', exact: true })).toBeChecked()
    await expect(
      dialog.getByRole('radio', { name: 'Comments and Notes', exact: true }),
    ).toBeDisabled()
    await dialog.getByRole('radio', { name: 'Values', exact: true }).check()
    await dialog.getByRole('radio', { name: 'Add', exact: true }).check()
    await dialog.getByRole('checkbox', { name: 'Transpose', exact: true }).check()
    await dialog.getByRole('checkbox', { name: 'Skip blanks', exact: true }).check()
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog).toBeHidden()
    const cells = await page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells)
    expect([cells.D5.rawValue, cells.E5.rawValue, cells.D6.rawValue, cells.E6.rawValue]).toEqual([
      12, 24, 30, 40,
    ])
    expect(cells.D5.style.fillColor).toBe('#CCEEFF')
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(before + 1)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
  }
})

test('grid Paste Special shortcut links original coordinates and column widths follow the current desktop source after an edit', async ({
  page,
}) => {
  await mountSpecialPaste(page)
  try {
    await page.locator('.workbook-grid').evaluate((el) => {
      el.tabIndex = 0
      el.focus()
    })
    await page.keyboard.press('Control+Alt+V')
    const dialog = page.getByRole('dialog', { name: 'Paste Special', exact: true })
    await dialog.getByRole('button', { name: 'Paste Link', exact: true }).click()
    await expect(dialog).toBeHidden()
    expect(
      await page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells.D5.formula),
    ).toBe('=A80')
    await page.evaluate(() => (window as any).toolbar.resizeSpecialSource())
    await specialChoice(page, 'Column Widths Only')
    await expect
      .poll(() =>
        page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].features.colWidths.D),
      )
      .toBe(240)
    expect(
      await page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].features.colWidths.E),
    ).toBe(120)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
  }
})

test('Paste Special keeps failed choices for retry, waits during Save and rejects stale dialog roots', async ({
  page,
}) => {
  await mountSpecialPaste(page)
  try {
    await specialChoice(page, 'Paste Special…')
    const dialog = page.getByRole('dialog', { name: 'Paste Special', exact: true })
    await dialog.getByRole('radio', { name: 'Values and number formats', exact: true }).check()
    await page.evaluate(() =>
      (window as any).toolbar.rejectNextSubmission('Rejected Paste Special'),
    )
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog.getByRole('alert')).toContainText('Rejected Paste Special')
    await expect(
      dialog.getByRole('radio', { name: 'Values and number formats', exact: true }),
    ).toBeChecked()
    await page.evaluate(() => (window as any).toolbar.holdSubmissions())
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog.getByRole('button', { name: 'OK', exact: true })).toBeDisabled()
    await page.evaluate(() => {
      ;(window as any).specialSave = 'waiting'
      void (window as any).toolbar.flush().then((ok: boolean) => {
        ;(window as any).specialSave = ok ? 'ready' : 'failed'
      })
    })
    expect(await page.evaluate(() => (window as any).specialSave)).toBe('waiting')
    await page.evaluate(() => (window as any).toolbar.releaseSubmissions())
    await expect(dialog).toBeHidden()
    await expect.poll(() => page.evaluate(() => (window as any).specialSave)).toBe('ready')
    expect(
      await page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells.D5.numberFormat),
    ).toBe('0.00')
    await specialChoice(page, 'Paste Special…')
    await expect(dialog).toBeVisible()
    await page.evaluate(() => (window as any).toolbar.resizeSpecialSource())
    const before = await page.evaluate(() => (window as any).toolbar.snapshot())
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog.getByRole('alert')).toBeVisible()
    expect(await page.evaluate(() => (window as any).toolbar.snapshot())).toEqual(before)
    await dialog.getByRole('button', { name: 'Cancel', exact: true }).click()
    await expect(dialog).toBeHidden()
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
  }
})

test('external Paste Special uses desktop HTML value hooks and restricts the shared dialog to text-supported choices', async ({
  page,
}) => {
  await mountSpecialPaste(page)
  try {
    await page.evaluate(async () =>
      navigator.clipboard.write([
        new ClipboardItem({
          'text/html': new Blob(
            [
              '<table><tr><td style="background-color:rgb(255,0,0);font-weight:bold">12</td><td>TRUE</td></tr></table>',
            ],
            { type: 'text/html' },
          ),
          'text/plain': new Blob(['12\tTRUE'], { type: 'text/plain' }),
        }),
      ]),
    )
    await specialChoice(page, 'Values Only')
    await expect
      .poll(() =>
        page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells.D5.rawValue),
      )
      .toBe(12)
    expect(
      await page.evaluate(
        () => (window as any).toolbar.snapshot().sheets[0].cells.D5.style.fillColor,
      ),
    ).toBe('#CCEEFF')
    expect(
      await page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells.E5.kind),
    ).toBe('string')
    await specialChoice(page, 'Paste Special…')
    const dialog = page.getByRole('dialog', { name: 'Paste Special', exact: true })
    await expect(dialog.getByRole('radio', { name: 'Formats', exact: true })).toBeDisabled()
    await expect(dialog.getByRole('button', { name: 'Paste Link', exact: true })).toBeDisabled()
    await dialog.getByRole('checkbox', { name: 'Transpose', exact: true }).check()
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog).toBeHidden()
    expect(
      await page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells.D6),
    ).toMatchObject({ kind: 'boolean', rawValue: true })
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
  }
})

test('Paste Special Formulas Only shifts references and retains target formatting for pending formulas', async ({
  page,
}) => {
  await mountSpecialPaste(page)
  try {
    await page.evaluate(async () => {
      await (window as any).toolbar.setSpecialFormula()
      ;(window as any).toolbar.select('A80:B81')
    })
    await page
      .locator('.workbook-desktop-header')
      .getByRole('button', { name: /^Copy / })
      .click()
    await page.evaluate(() => (window as any).toolbar.flush())
    await page.evaluate(() => (window as any).toolbar.select('D5'))
    await specialChoice(page, 'Formulas Only')
    await expect
      .poll(() =>
        page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells.E6.formula),
      )
      .toBe('=D5+$A$1')
    expect(
      await page.evaluate(
        () => (window as any).toolbar.snapshot().sheets[0].cells.D5.style.fillColor,
      ),
    ).toBe('#CCEEFF')
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
  }
})

test('Paste Special follows desktop live-source rules for advanced values while plain Values keeps the copied snapshot', async ({
  page,
}) => {
  await mountSpecialPaste(page)
  try {
    await page.evaluate(() => (window as any).toolbar.changeCopiedSource())
    await specialChoice(page, 'Values Only')
    await expect
      .poll(() =>
        page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells.D5.rawValue),
      )
      .toBe(2)
    await specialChoice(page, 'Paste Special…')
    const dialog = page.getByRole('dialog', { name: 'Paste Special', exact: true })
    await dialog.getByRole('radio', { name: 'Values and number formats', exact: true }).check()
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog).toBeHidden()
    expect(
      await page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells.D5),
    ).toMatchObject({ rawValue: 'New source', numberFormat: '0.00' })
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
  }
})

test('Paste Special All keeps cached values and uses current desktop source formatting', async ({
  page,
}) => {
  await mountSpecialPaste(page)
  try {
    await page.evaluate(async () => {
      await (window as any).toolbar.changeCopiedSource()
      await (window as any).toolbar.formatSpecialSource()
    })
    await specialChoice(page, 'Paste Special…')
    const dialog = page.getByRole('dialog', { name: 'Paste Special', exact: true })
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog).toBeHidden()
    expect(
      await page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells.D5),
    ).toMatchObject({ rawValue: 2, style: { bold: false, fontColor: '#445566' } })
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
  }
})

async function mountPageLayout(page: import('@playwright/test').Page) {
  page.setDefaultTimeout(8000)
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(
      false,
      false,
      false,
      false,
      false,
      false,
      false,
      false,
      true,
    )
    ;(window as any).toolbar.select('B4:C50')
  })
  await page.getByRole('tab', { name: 'Page Layout', exact: true }).click()
}
async function pageSettingChoice(
  page: import('@playwright/test').Page,
  name: string,
  option: string,
) {
  await page.locator('.workbook-desktop-header').getByRole('button', { name, exact: true }).click()
  await page.getByRole('option', { name: option, exact: true }).click()
}

test('shared Page Layout band edits saved page settings and keeps view and print flags separate', async ({
  page,
}) => {
  await mountPageLayout(page)
  try {
    await expect(page.locator('#ribbon-panel-pageLayout')).toHaveClass(/workbook-desktop-panel/)
    await expect(page.locator('#ribbon-panel-pageLayout')).not.toHaveClass(/workbook-toolbar/)
    const header = page.locator('.workbook-desktop-header')
    await expect(header.getByRole('button', { name: 'Themes', exact: true })).toBeDisabled()
    await expect(
      header.getByRole('button', { name: 'Orientation', exact: true }).locator('..').locator('..'),
    ).toHaveAttribute('data-tip', /Landscape/)
    await pageSettingChoice(page, 'Orientation', 'Portrait')
    await pageSettingChoice(page, 'Margins', 'Narrow')
    await pageSettingChoice(page, 'Size', 'A3')
    await pageSettingChoice(page, 'Print Area', 'Set Print Area')
    await pageSettingChoice(page, 'Print Titles', 'Repeat Selected Rows')
    await pageSettingChoice(page, 'Fit to width', '3 pages')
    const gridlines = header.locator('.check-column').filter({ hasText: 'Gridlines' })
    await gridlines.getByRole('button', { name: /Print/ }).click()
    await expect
      .poll(() =>
        page.evaluate(
          () => (window as any).toolbar.snapshot().sheets[0].features.pageSetupEdits.printGridlines,
        ),
      )
      .toBe(false)
    expect(
      await page.evaluate(
        () => (window as any).toolbar.snapshot().sheets[0].features.pageSetupEdits,
      ),
    ).toMatchObject({
      orientation: 'portrait',
      margins: 'narrow',
      paperSize: 8,
      printArea: 'B4:C50',
      printTitles: '4:24',
      fitToWidth: 3,
      fitToHeight: 0,
    })
    await gridlines.getByRole('button', { name: /View/ }).click()
    await expect
      .poll(() =>
        page.evaluate(
          () => (window as any).toolbar.snapshot().sheets[0].features.pageSetupEdits.showGridlines,
        ),
      )
      .toBe(false)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
  }
})

test('shared Page Layout break menus preserve file breaks and remove or reset them atomically', async ({
  page,
}) => {
  await mountPageLayout(page)
  try {
    await pageSettingChoice(page, 'Breaks', 'Insert Page Break')
    await expect
      .poll(() =>
        page.evaluate(
          () => (window as any).toolbar.snapshot().sheets[0].features.pageSetupEdits.rowBreaks,
        ),
      )
      .toEqual([2, 3, 10])
    expect(
      await page.evaluate(
        () => (window as any).toolbar.snapshot().sheets[0].features.pageSetupEdits.colBreaks,
      ),
    ).toEqual([1])
    await pageSettingChoice(page, 'Breaks', 'Remove Page Break')
    await expect
      .poll(() =>
        page.evaluate(
          () => (window as any).toolbar.snapshot().sheets[0].features.pageSetupEdits.rowBreaks,
        ),
      )
      .toEqual([2, 10])
    await pageSettingChoice(page, 'Breaks', 'Reset All Page Breaks')
    await expect
      .poll(() =>
        page.evaluate(
          () => (window as any).toolbar.snapshot().sheets[0].features.pageSetupEdits.colBreaks,
        ),
      )
      .toEqual([])
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
  }
})

test('shared Page Layout waits for Save and rejects settings read against an overtaken root', async ({
  page,
}) => {
  await mountPageLayout(page)
  try {
    await page.evaluate(() => (window as any).toolbar.holdSubmissions())
    await pageSettingChoice(page, 'Orientation', 'Portrait')
    await page.evaluate(() => {
      ;(window as any).pageSave = 'waiting'
      void (window as any).toolbar.flush().then((ok: boolean) => {
        ;(window as any).pageSave = ok ? 'ready' : 'failed'
      })
    })
    expect(await page.evaluate(() => (window as any).pageSave)).toBe('waiting')
    await page.evaluate(() => (window as any).toolbar.releaseSubmissions())
    await expect.poll(() => page.evaluate(() => (window as any).pageSave)).toBe('ready')
    await page.evaluate(() => (window as any).toolbar.holdReads())
    await pageSettingChoice(page, 'Size', 'A4')
    const revision = await page.evaluate(() => (window as any).toolbar.revision())
    await page.evaluate(() => {
      ;(window as any).pageConcurrentEdit = (window as any).toolbar.resizeSpecialSource()
    })
    await expect
      .poll(() => page.evaluate(() => (window as any).toolbar.revision()))
      .toBe(revision + 1)
    const before = await page.evaluate(() => (window as any).toolbar.snapshot())
    await page.evaluate(async () => {
      ;(window as any).toolbar.releaseReads()
      await (window as any).pageConcurrentEdit
    })
    await expect(page.locator('.workbook-error')).toContainText('changed')
    expect(await page.evaluate(() => (window as any).toolbar.snapshot())).toEqual(before)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
  }
})

test('Page Layout scale popovers close after selection in a narrow editor and leave print controls clickable', async ({
  page,
}) => {
  await page.setViewportSize({ width: 900, height: 720 })
  await mountPageLayout(page)
  try {
    await pageSettingChoice(page, 'Fit to width', '3 pages')
    const header = page.locator('.workbook-desktop-header')
    await expect(header.getByRole('listbox')).toHaveCount(0)
    await header
      .locator('.check-column')
      .filter({ hasText: 'Gridlines' })
      .getByRole('button', { name: /Print/ })
      .click()
    await expect
      .poll(() =>
        page.evaluate(
          () => (window as any).toolbar.snapshot().sheets[0].features.pageSetupEdits.printGridlines,
        ),
      )
      .toBe(false)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose()).catch(() => {})
  }
})

async function mountFunctionLibrary(page: import('@playwright/test').Page) {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar()
    ;(window as any).toolbar.select('C4')
  })
  await page.getByRole('tab', { name: 'Formulas', exact: true }).click()
}
async function pickFunction(page: import('@playwright/test').Page, name: string) {
  const catalog = page.getByRole('dialog', { name: 'Insert Function', exact: true })
  await catalog.getByPlaceholder('Search functions…').fill(name)
  await catalog.getByRole('option', { name: new RegExp(`^${name}\\b`) }).click()
  await catalog.getByRole('button', { name: 'OK', exact: true }).click()
  return page.getByRole('dialog', { name: 'Function Arguments', exact: true })
}

test('shared function library searches the live catalog, previews authoritative inputs and freezes its write target', async ({
  page,
}) => {
  await mountFunctionLibrary(page)
  try {
    await page
      .locator('.workbook-desktop-header')
      .getByRole('button', { name: 'Insert Function', exact: true })
      .click()
    const args = await pickFunction(page, 'SUM')
    await args.getByRole('textbox').first().fill('B2:B3')
    await expect(args.locator('.fn-args-result')).toContainText('90')
    // Programmatic selection changes do not count as a pointer range pick or retarget the dialog.
    await page.evaluate(() => (window as any).toolbar.select('D8'))
    await args.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(args).not.toBeVisible()
    await expect
      .poll(
        async () =>
          (await page.evaluate(() => (window as any).toolbar.snapshot())).sheets[0].cells.C4
            ?.formula,
      )
      .toBe('=SUM(B2:B3)')
    const snapshot = await page.evaluate(() => (window as any).toolbar.snapshot())
    expect(snapshot.sheets[0].cells.D8).toBeUndefined()
    expect(snapshot.sheets[0].cells.B2.rawValue).toBe(120)
    expect(snapshot.sheets[0].cells.B3.rawValue).toBe(-30)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('function arguments retain rejected drafts, disable pending actions and make Save await the canonical commit', async ({
  page,
}) => {
  await mountFunctionLibrary(page)
  try {
    await page
      .locator('.workbook-desktop-header')
      .getByRole('button', { name: 'Insert Function', exact: true })
      .click()
    const args = await pickFunction(page, 'SUM')
    await args.getByRole('textbox').first().fill('1,2')
    await page.evaluate(() => (window as any).toolbar.rejectNextSubmission('function-retry'))
    await args.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(args.getByRole('alert')).toContainText('function-retry')
    await expect(args.getByRole('textbox').first()).toHaveValue('1,2')
    const before = await page.evaluate(() => (window as any).toolbar.snapshot())
    await page.evaluate(() => (window as any).toolbar.holdSubmissions())
    await args.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(args.getByRole('button', { name: 'OK', exact: true })).toBeDisabled()
    await expect(args.getByRole('button', { name: 'Cancel', exact: true })).toBeDisabled()
    await page.evaluate(() => {
      ;(window as any).functionSaveFinished = false
      ;(window as any).functionSave = (window as any).toolbar.flush().then((result: boolean) => {
        ;(window as any).functionSaveFinished = true
        return result
      })
    })
    expect(await page.evaluate(() => (window as any).functionSaveFinished)).toBe(false)
    expect(await page.evaluate(() => (window as any).toolbar.snapshot())).toEqual(before)
    await page.evaluate(() => (window as any).toolbar.releaseSubmissions())
    await expect(args).not.toBeVisible()
    expect(await page.evaluate(() => (window as any).functionSave)).toBe(true)
    await expect
      .poll(
        async () =>
          (await page.evaluate(() => (window as any).toolbar.snapshot())).sheets[0].cells.C4
            ?.formula,
      )
      .toBe('=SUM(1,2)')
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('function arguments reject an overtaken document and cancel without a write', async ({
  page,
}) => {
  await mountFunctionLibrary(page)
  try {
    const header = page.locator('.workbook-desktop-header')
    await header.getByRole('button', { name: 'Insert Function', exact: true }).click()
    const args = await pickFunction(page, 'SUM')
    await args.getByRole('textbox').first().fill('1')
    await page.evaluate(() => (window as any).toolbar.resizeSpecialSource())
    const before = await page.evaluate(() => (window as any).toolbar.snapshot())
    await args.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(args.getByRole('alert')).toContainText('document changed')
    await args.getByRole('button', { name: 'Cancel', exact: true }).click()
    expect(await page.evaluate(() => (window as any).toolbar.snapshot())).toEqual(before)
    await header.getByRole('button', { name: 'Math & Trig', exact: true }).click()
    await expect(page.getByRole('dialog', { name: 'Insert Function', exact: true })).toBeVisible()
    await page.getByRole('dialog').getByRole('button', { name: 'Cancel', exact: true }).click()
    expect(await page.evaluate(() => (window as any).toolbar.snapshot())).toEqual(before)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('shared AutoSum applies each selected column as one transaction using desktop geometry', async ({
  page,
}) => {
  await mountFunctionLibrary(page)
  try {
    await page.evaluate(() => (window as any).toolbar.select('A2:B3'))
    const before = await page.evaluate(() => (window as any).toolbar.submissions())
    await page.getByRole('button', { name: 'AutoSum', exact: true }).click()
    await page.getByRole('option', { name: 'Sum', exact: true }).click()
    await expect
      .poll(
        async () =>
          (await page.evaluate(() => (window as any).toolbar.snapshot())).sheets[0].cells.B4
            ?.formula,
      )
      .toBe('=SUM(B2:B3)')
    const snapshot = await page.evaluate(() => (window as any).toolbar.snapshot())
    expect(snapshot.sheets[0].cells.A4.formula).toBe('=SUM(A2:A3)')
    expect(await page.evaluate(() => (window as any).toolbar.submissions())).toBe(before + 1)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('function library reads off-page original inputs and leaves unverified computed precedents blank', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(false, false, false, false, false, true)
    ;(window as any).toolbar.select('E4')
  })
  try {
    await page.getByRole('tab', { name: 'Formulas', exact: true }).click()
    await page
      .locator('.workbook-desktop-header')
      .getByRole('button', { name: 'Insert Function', exact: true })
      .click()
    const args = await pickFunction(page, 'SUM')
    await args.getByRole('textbox').first().fill('B79:B81')
    await expect(args.locator('.fn-args-result')).toContainText('38')
    await args.getByRole('textbox').first().fill('D78')
    await expect(args.locator('.fn-args-result strong')).toHaveText('')
    await args.getByRole('button', { name: 'Cancel', exact: true }).click()
    const snapshot = await page.evaluate(() => (window as any).toolbar.snapshot())
    expect(snapshot.sheets[0].cells.E4).toBeUndefined()
    expect(snapshot.sheets[0].cells.B79).toBeUndefined()
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('grid Shift F3 opens the shared function library and ignores typing outside the grid', async ({
  page,
}) => {
  await mountFunctionLibrary(page)
  try {
    const grid = page.locator('.workbook-grid')
    await grid.evaluate((el) => {
      el.tabIndex = 0
      el.focus()
    })
    await page.keyboard.press('Shift+F3')
    const catalog = page.getByRole('dialog', { name: 'Insert Function', exact: true })
    await expect(catalog).toBeVisible()
    await catalog.getByRole('button', { name: 'Cancel', exact: true }).click()
    await page
      .locator('.workbook-additional-controls summary')
      .filter({ hasText: 'Insert formula' })
      .click()
    await page.getByRole('textbox', { name: 'Formula text', exact: true }).focus()
    await page.keyboard.press('Shift+F3')
    await expect(catalog).not.toBeVisible()
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('shared function argument range picker observes pointer selections inside the embedded grid', async ({
  page,
}) => {
  await mountFunctionLibrary(page)
  try {
    await page
      .locator('.workbook-desktop-header')
      .getByRole('button', { name: 'Insert Function', exact: true })
      .click()
    const args = await pickFunction(page, 'SUM')
    await args
      .getByRole('button', { name: 'Select a range in the sheet', exact: true })
      .first()
      .click()
    const point = await page.locator('.workbook-grid canvas').evaluateAll((nodes) => {
      const canvas = (nodes as HTMLCanvasElement[]).sort(
        (a, b) => b.width * b.height - a.width * a.height,
      )[0]!
      const box = canvas.getBoundingClientRect(),
        cell = (window as any).toolbar.cellRect('B2')
      return { x: box.x + (cell.left + cell.right) / 2, y: box.y + (cell.top + cell.bottom) / 2 }
    })
    await page.mouse.click(point.x, point.y)
    await expect(args.getByRole('textbox').first()).toHaveValue('B2')
    await args.getByRole('button', { name: 'Expand dialog', exact: true }).click()
    await args.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(args).not.toBeVisible()
    await expect
      .poll(
        async () =>
          (await page.evaluate(() => (window as any).toolbar.snapshot())).sheets[0].cells.C4
            ?.formula,
      )
      .toBe('=SUM(B2)')
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('native formula bar fx inserts a valid canonical formula even when the editor reports caret zero', async ({
  page,
}) => {
  await mountFunctionLibrary(page)
  try {
    const before = await page.evaluate(() => (window as any).toolbar.snapshot())
    await page.locator('.workbook-grid .univerjs-icon-fx-icon').locator('..').click()
    const args = await pickFunction(page, 'SUM')
    expect(await page.evaluate(() => (window as any).toolbar.snapshot())).toEqual(before)
    await args.getByRole('textbox').first().fill('B2:B3')
    await args.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(args).not.toBeVisible()
    await expect
      .poll(
        async () =>
          (await page.evaluate(() => (window as any).toolbar.snapshot())).sheets[0].cells.C4
            ?.formula,
      )
      .toBe('=SUM(B2:B3)')
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('function opening is single-instance and Save awaits the authoritative target read', async ({
  page,
}) => {
  await mountFunctionLibrary(page)
  try {
    const before = await page.evaluate(() => (window as any).toolbar.snapshot())
    await page.evaluate(() => (window as any).toolbar.holdReads())
    await page
      .locator('.workbook-desktop-header')
      .getByRole('button', { name: 'Insert Function', exact: true })
      .dblclick()
    await page.evaluate(() => {
      ;(window as any).functionOpeningFinished = false
      ;(window as any).functionOpeningSave = (window as any).toolbar
        .flush()
        .then((value: boolean) => {
          ;(window as any).functionOpeningFinished = true
          return value
        })
    })
    expect(await page.evaluate(() => (window as any).functionOpeningFinished)).toBe(false)
    await page.evaluate(() => (window as any).toolbar.releaseReads())
    const catalog = page.getByRole('dialog', { name: 'Insert Function', exact: true })
    await expect(catalog).toHaveCount(1)
    expect(await page.evaluate(() => (window as any).functionOpeningSave)).toBe(true)
    await catalog.getByRole('button', { name: 'Cancel', exact: true }).click()
    expect(await page.evaluate(() => (window as any).toolbar.snapshot())).toEqual(before)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('function confirmation preserves an unfinished grid draft and rejects its overtaken target', async ({
  page,
}) => {
  await mountFunctionLibrary(page)
  try {
    await page
      .locator('.workbook-desktop-header')
      .getByRole('button', { name: 'Insert Function', exact: true })
      .click()
    const args = await pickFunction(page, 'SUM')
    await args.getByRole('textbox').first().fill('1')
    expect(await page.evaluate(() => (window as any).toolbar.beginMergedEdit())).toBe(true)
    await page.keyboard.press('ControlOrMeta+A')
    await page.keyboard.type('Keep this draft')
    await args.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(args.getByRole('alert')).toContainText('document changed')
    const snapshot = await page.evaluate(() => (window as any).toolbar.snapshot())
    expect(snapshot.sheets[0].cells.A1.rawValue).toBe('Keep this draft')
    expect(snapshot.sheets[0].cells.C4).toBeUndefined()
    await args.getByRole('button', { name: 'Cancel', exact: true }).click()
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('canonical name edits update Go To and the shared engine projection without cell history writes', async ({
  page,
}) => {
  await mountFunctionLibrary(page)
  try {
    const before = await page.evaluate(() => (window as any).toolbar.snapshot())
    await page.evaluate(() => (window as any).toolbar.setName('NewRange', 'Data!$B$2'))
    await page.getByRole('tab', { name: 'Home', exact: true }).click()
    await page.locator('.workbook-desktop-header').getByRole('button', { name: /Go To/ }).click()
    const go = page.getByRole('dialog', { name: 'Go To', exact: true })
    await expect(go.getByRole('button', { name: /NewRange/ })).toBeVisible()
    await go.getByRole('textbox', { name: 'Reference', exact: true }).fill('NewRange')
    await go.getByRole('button', { name: 'Go', exact: true }).click()
    await expect(go).not.toBeVisible()
    expect(await page.evaluate(() => (window as any).toolbar.selection())).toMatchObject({
      selection: { address: 'B2:B2' },
    })
    const after = await page.evaluate(() => (window as any).toolbar.snapshot())
    expect(after.sheets[0].cells).toEqual(before.sheets[0].cells)
    await page.evaluate(() => (window as any).toolbar.deleteName('NewRange'))
    await page.locator('.workbook-desktop-header').getByRole('button', { name: /Go To/ }).click()
    await expect(go.getByRole('button', { name: /NewRange/ })).toHaveCount(0)
    await go.getByRole('button', { name: 'Cancel', exact: true }).click()
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

async function openSharedNames(page: import('@playwright/test').Page) {
  await page.getByRole('tab', { name: 'Formulas', exact: true }).click()
  await page
    .locator('.workbook-desktop-header')
    .getByRole('button', { name: 'Name Manager', exact: true })
    .click()
  return page.getByRole('dialog', { name: 'Name Manager', exact: true })
}
async function nameInput(dialog: import('@playwright/test').Locator, name: string, ref: string) {
  await dialog.getByRole('textbox', { name: 'Name', exact: true }).fill(name)
  await dialog.getByRole('textbox', { name: 'Refers to', exact: true }).fill(ref)
}

test('shared Name Manager renames and removes dependent formulas and retains source cells', async ({
  page,
}) => {
  await mountFunctionLibrary(page)
  try {
    await page.evaluate(async () => {
      await (window as any).toolbar.setName('Revenue', 'Data!$B$2')
      await (window as any).toolbar.putFormula('C5', '=Revenue+1')
    })
    const dialog = await openSharedNames(page)
    await dialog.getByRole('cell', { name: 'Revenue', exact: true }).click()
    await nameInput(dialog, 'Sales', 'Data!$B$2')
    await dialog.getByRole('button', { name: 'Update', exact: true }).click()
    await expect(dialog.getByRole('cell', { name: 'Sales', exact: true })).toBeVisible()
    await expect
      .poll(
        async () =>
          (await page.evaluate(() => (window as any).toolbar.snapshot())).sheets[0].cells.C5
            .formula,
      )
      .toBe('=Sales+1')
    await dialog
      .getByRole('cell', { name: 'Sales', exact: true })
      .locator('..')
      .getByRole('button', { name: 'Delete name', exact: true })
      .click()
    await expect(dialog.getByRole('cell', { name: 'Sales', exact: true })).toHaveCount(0)
    await expect
      .poll(
        async () =>
          (await page.evaluate(() => (window as any).toolbar.snapshot())).sheets[0].cells.C5
            .formula,
      )
      .toBe('=#REF!+1')
    expect(
      (await page.evaluate(() => (window as any).toolbar.snapshot())).sheets[0].cells.B2.rawValue,
    ).toBe(120)
    await dialog.getByRole('button', { name: 'Close', exact: true }).click()
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('shared Name Manager keeps rejected drafts, blocks pending actions and makes Save await one definition commit', async ({
  page,
}) => {
  await mountFunctionLibrary(page)
  try {
    const dialog = await openSharedNames(page)
    await nameInput(dialog, 'Revenue', 'Data!$B$2')
    await page.evaluate(() => (window as any).toolbar.rejectNextSubmission('name-retry'))
    await dialog.getByRole('button', { name: 'Add', exact: true }).click()
    await expect(dialog.getByRole('alert')).toContainText('name-retry')
    await expect(dialog.getByRole('textbox', { name: 'Name', exact: true })).toHaveValue('Revenue')
    await page.evaluate(() => (window as any).toolbar.holdSubmissions())
    await dialog.getByRole('button', { name: 'Add', exact: true }).click()
    await expect(dialog.getByRole('button', { name: 'Close', exact: true })).toBeDisabled()
    await expect(dialog.getByRole('textbox', { name: 'Name', exact: true })).toBeDisabled()
    await page.evaluate(() => {
      ;(window as any).nameSaveFinished = false
      ;(window as any).nameSave = (window as any).toolbar.flush().then((value: boolean) => {
        ;(window as any).nameSaveFinished = true
        return value
      })
    })
    expect(await page.evaluate(() => (window as any).nameSaveFinished)).toBe(false)
    await page.evaluate(() => (window as any).toolbar.releaseSubmissions())
    await expect(dialog.getByRole('cell', { name: 'Revenue', exact: true })).toBeVisible()
    expect(await page.evaluate(() => (window as any).nameSave)).toBe(true)
    await expect(dialog.getByRole('textbox', { name: 'Name', exact: true })).toHaveValue('')
    await nameInput(dialog, 'Revenue', '1')
    await dialog.getByRole('button', { name: 'Add', exact: true }).click()
    await expect(dialog.getByRole('alert')).toContainText('already used')
    await dialog.getByRole('button', { name: 'Close', exact: true }).click()
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('shared Name Manager distinguishes workbook and sheet scopes and rejects stale edits', async ({
  page,
}) => {
  await mountFunctionLibrary(page)
  try {
    const dialog = await openSharedNames(page)
    await nameInput(dialog, 'Scoped', '1')
    await dialog.getByRole('button', { name: 'Add', exact: true }).click()
    await expect(dialog.getByRole('cell', { name: 'Scoped', exact: true })).toHaveCount(1)
    await nameInput(dialog, 'Scoped', '2')
    await dialog.getByRole('button', { name: 'Scope', exact: true }).click()
    await page.getByRole('option', { name: 'Data', exact: true }).click()
    await dialog.getByRole('button', { name: 'Add', exact: true }).click()
    await expect(dialog.getByRole('cell', { name: 'Scoped', exact: true })).toHaveCount(2)
    const root = await page.evaluate(() => (window as any).toolbar.snapshot())
    expect(root.features.definedNamesEdits.names).toEqual(
      expect.arrayContaining([
        { name: 'Scoped', formula: '1' },
        { name: 'Scoped', formula: '2', scopeSheetId: 'toolbar-sheet' },
      ]),
    )
    await page.evaluate(() => (window as any).toolbar.resizeSpecialSource())
    await nameInput(dialog, 'AfterEdit', '3')
    await dialog.getByRole('button', { name: 'Add', exact: true }).click()
    await expect(dialog.getByRole('alert')).toContainText('document changed')
    await dialog.getByRole('button', { name: 'Close', exact: true }).click()
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('shared Defined Names creates names from labels once and inserts a name only into an empty cell', async ({
  page,
}) => {
  await mountFunctionLibrary(page)
  try {
    await page.evaluate(() => (window as any).toolbar.select('A1:B10'))
    const header = page.locator('.workbook-desktop-header')
    const create = async () => {
      await header.getByRole('button', { name: 'Create from Selection', exact: true }).click()
      await page.getByRole('option', { name: 'From top row', exact: true }).click()
    }
    const before = await page.evaluate(() => (window as any).toolbar.submissions())
    await create()
    await expect(page.getByRole('status').filter({ hasText: 'Created 2 names.' })).toBeVisible()
    let root = await page.evaluate(() => (window as any).toolbar.snapshot())
    expect(root.features.definedNamesEdits.names).toEqual([
      { name: 'Region', formula: "'Data'!$A$2:$A$10" },
      { name: 'Sales', formula: "'Data'!$B$2:$B$10" },
    ])
    expect(await page.evaluate(() => (window as any).toolbar.submissions())).toBe(before + 1)
    await create()
    await expect(
      page.getByRole('status').filter({ hasText: 'Created 0 names, skipped 2.' }),
    ).toBeVisible()
    expect(await page.evaluate(() => (window as any).toolbar.submissions())).toBe(before + 1)
    await page.evaluate(() => (window as any).toolbar.select('D5'))
    await header.getByRole('button', { name: 'Use in Formula', exact: true }).click()
    await page.getByRole('option', { name: 'Sales', exact: true }).click()
    await expect
      .poll(
        async () =>
          (await page.evaluate(() => (window as any).toolbar.snapshot())).sheets[0].cells.D5
            ?.formula,
      )
      .toBe('=Sales')
    await page.evaluate(() => (window as any).toolbar.select('B2'))
    await header.getByRole('button', { name: 'Use in Formula', exact: true }).click()
    await page.getByRole('option', { name: 'Sales', exact: true }).click()
    await expect(page.getByRole('alert')).toContainText('empty cell')
    root = await page.evaluate(() => (window as any).toolbar.snapshot())
    expect(root.sheets[0].cells.B2.rawValue).toBe(120)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('shared Create from Selection formats date labels and creates over 100 definitions in one transaction', async ({
  page,
}) => {
  await mountFunctionLibrary(page)
  try {
    await page.evaluate(async () => {
      await (window as any).toolbar.dateLabel()
      ;(window as any).toolbar.select('A1:A2')
    })
    const header = page.locator('.workbook-desktop-header')
    await header.getByRole('button', { name: 'Create from Selection', exact: true }).click()
    await page.getByRole('option', { name: 'From top row', exact: true }).click()
    await expect
      .poll(
        async () =>
          (await page.evaluate(() => (window as any).toolbar.snapshot())).features.definedNamesEdits
            ?.names[0]?.name,
      )
      .toBe('Jan_2023')
    await page.evaluate(async () => {
      await (window as any).toolbar.putHeaders(Array.from({ length: 101 }, (_, i) => `Header_${i}`))
      ;(window as any).toolbar.select('A1:CW2')
    })
    const before = await page.evaluate(() => (window as any).toolbar.submissions())
    await header.getByRole('button', { name: 'Create from Selection', exact: true }).click()
    await page.getByRole('option', { name: 'From top row', exact: true }).click()
    await expect
      .poll(
        async () =>
          (await page.evaluate(() => (window as any).toolbar.snapshot())).features.definedNamesEdits
            .names.length,
      )
      .toBe(102)
    expect(await page.evaluate(() => (window as any).toolbar.submissions())).toBe(before + 1)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('shared Create from Selection supports row labels and rejects unverified formula labels without a name write', async ({
  page,
}) => {
  await mountFunctionLibrary(page)
  try {
    await page.evaluate(() => (window as any).toolbar.select('A1:C3'))
    const header = page.locator('.workbook-desktop-header')
    await header.getByRole('button', { name: 'Create from Selection', exact: true }).click()
    await page.getByRole('option', { name: 'From left column', exact: true }).click()
    await expect
      .poll(
        async () =>
          (await page.evaluate(() => (window as any).toolbar.snapshot())).features.definedNamesEdits
            ?.names.length,
      )
      .toBe(3)
    const root = await page.evaluate(() => (window as any).toolbar.snapshot())
    expect(root.features.definedNamesEdits.names).toEqual([
      { name: 'Region', formula: "'Data'!$B$1:$C$1" },
      { name: 'North', formula: "'Data'!$B$2:$C$2" },
      { name: 'South', formula: "'Data'!$B$3:$C$3" },
    ])
    await page.evaluate(async () => {
      await (window as any).toolbar.putFormula('A1', '=SUM(B2:B3)')
      ;(window as any).toolbar.select('A1:B3')
    })
    const before = await page.evaluate(() => (window as any).toolbar.snapshot())
    await header.getByRole('button', { name: 'Create from Selection', exact: true }).click()
    await page.getByRole('option', { name: 'From top row', exact: true }).click()
    await expect(page.getByRole('alert')).toContainText('Formula labels are not verified')
    expect(await page.evaluate(() => (window as any).toolbar.snapshot())).toEqual(before)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('shared Create from Selection waits for Save and rejects an original-label read overtaken by another edit', async ({
  page,
}) => {
  await mountFunctionLibrary(page)
  try {
    await page.evaluate(() => {
      ;(window as any).toolbar.select('A1:B3')
      ;(window as any).toolbar.holdReads()
    })
    const header = page.locator('.workbook-desktop-header')
    await header.getByRole('button', { name: 'Create from Selection', exact: true }).click()
    await page.getByRole('option', { name: 'From top row', exact: true }).click()
    await page.evaluate(() => {
      ;(window as any).namesReadSaveDone = false
      ;(window as any).namesReadSave = (window as any).toolbar.flush().then((value: boolean) => {
        ;(window as any).namesReadSaveDone = true
        return value
      })
      ;(window as any).namesConcurrentEdit = (window as any).toolbar.resizeSpecialSource()
    })
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.revision())).toBe(1)
    expect(await page.evaluate(() => (window as any).namesReadSaveDone)).toBe(false)
    const before = await page.evaluate(() => (window as any).toolbar.snapshot())
    await page.evaluate(async () => {
      ;(window as any).toolbar.releaseReads()
      await (window as any).namesConcurrentEdit
    })
    await expect(page.getByRole('alert')).toContainText('document changed')
    expect(await page.evaluate(() => (window as any).namesReadSave)).toBe(false)
    expect(await page.evaluate(() => (window as any).toolbar.snapshot())).toEqual(before)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('shared Formula Auditing draws actual precedent and full-source dependent arrows without authored edits', async ({
  page,
}) => {
  await mountFunctionLibrary(page)
  try {
    await page.evaluate(async () => {
      await (window as any).toolbar.seedAuditSource()
      await (window as any).toolbar.putFormula('C5', '=B2+Other!A1')
      await (window as any).toolbar.select('C5:C5')
    })
    await page.getByRole('tab', { name: 'Formulas', exact: true }).click()
    const before = await page.evaluate(() => (window as any).toolbar.snapshot())
    const ribbon = page.locator('.workbook-desktop-header')
    await ribbon.getByRole('button', { name: 'Trace Precedents', exact: true }).click()
    await expect(page.locator('.trace-precedent-box')).toHaveCount(1)
    await expect(page.getByRole('status').filter({ hasText: 'Traced 1 precedent' })).toContainText(
      '1 reference(s) on other sheets',
    )
    await page.evaluate(() => (window as any).toolbar.select('B2:B2'))
    await ribbon.getByRole('button', { name: 'Trace Dependents', exact: true }).click()
    await expect(page.getByRole('status').filter({ hasText: 'Traced 2 dependent' })).toContainText(
      '1 dependent(s) on other sheets',
    )
    await expect(page.locator('.trace-arrow-layer')).toHaveCount(3)
    await ribbon.getByRole('button', { name: 'Remove Arrows', exact: true }).click()
    await expect(page.locator('.trace-arrow-layer')).toHaveCount(0)
    await expect(page.locator('.trace-precedent-box')).toHaveCount(0)
    expect(await page.evaluate(() => (window as any).toolbar.snapshot())).toEqual(before)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('shared Show Formulas displays actual formula text, persists one canonical view setting and respects per-sheet import', async ({
  page,
}) => {
  await mountFunctionLibrary(page)
  try {
    await page.evaluate(async () => {
      await (window as any).toolbar.seedAuditSource(true)
      await (window as any).toolbar.putFormula('C5', '=B2+1')
    })
    await page.getByRole('tab', { name: 'Formulas', exact: true }).click()
    await expect
      .poll(() => page.evaluate(() => (window as any).toolbar.displayedCell('C5').v))
      .toBe('=B2+1')
    const before = await page.evaluate(() => (window as any).toolbar.snapshot())
    await page
      .locator('.workbook-desktop-header')
      .getByRole('button', { name: 'Show Formulas', exact: true })
      .click()
    await expect
      .poll(
        async () =>
          (await page.evaluate(() => (window as any).toolbar.snapshot())).sheets[0].features
            .pageSetupEdits?.showFormulas,
      )
      .toBe(false)
    await expect
      .poll(() => page.evaluate(() => (window as any).toolbar.displayedCell('C5').v))
      .toBe(121)
    expect(
      (await page.evaluate(() => (window as any).toolbar.snapshot())).sheets.map(
        (sheet: any) => sheet.cells,
      ),
    ).toEqual(before.sheets.map((sheet: any) => sheet.cells))
    await page.evaluate(() => (window as any).toolbar.select('B2:B2'))
    await page.locator('.workbook-grid canvas').first().focus()
    await page.locator('.workbook-grid canvas').first().press('Control+Backquote')
    await expect
      .poll(
        async () =>
          (await page.evaluate(() => (window as any).toolbar.snapshot())).sheets[0].features
            .pageSetupEdits?.showFormulas,
      )
      .toBe(true)
    await page.evaluate(() => (window as any).toolbar.reveal("'Other'!A1"))
    await expect
      .poll(() => page.evaluate(() => (window as any).toolbar.displayedCell('A1').v))
      .toBe(120)
    await page.evaluate(() => (window as any).toolbar.reveal("'Data'!C5"))
    await expect
      .poll(() => page.evaluate(() => (window as any).toolbar.displayedCell('C5').v))
      .toBe('=B2+1')
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('shared Error Checking cycles off-page and cross-sheet source errors without edits', async ({
  page,
}) => {
  await mountFunctionLibrary(page)
  try {
    await page.evaluate(async () => {
      await (window as any).toolbar.seedAuditSource()
      await (window as any).toolbar.select('A1:A1')
    })
    await page.getByRole('tab', { name: 'Formulas', exact: true }).click()
    const before = await page.evaluate(() => (window as any).toolbar.snapshot())
    const check = page
      .locator('.workbook-desktop-header')
      .getByRole('button', { name: 'Error Checking', exact: true })
    await check.click()
    await expect
      .poll(() => page.evaluate(() => (window as any).toolbar.selection()))
      .toMatchObject({ selection: { address: 'D100001:D100001' } })
    await check.click()
    await expect
      .poll(() => page.evaluate(() => (window as any).toolbar.selection()))
      .toMatchObject({ selection: { sheetId: 'audit-other', address: 'B2:B2' } })
    await check.click()
    await expect
      .poll(() => page.evaluate(() => (window as any).toolbar.selection()))
      .toMatchObject({ selection: { address: 'D100001:D100001' } })
    expect(await page.evaluate(() => (window as any).toolbar.snapshot())).toEqual(before)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('shared Watch Window reads pinned cells, refreshes after edits and survives tab changes without history writes', async ({
  page,
}) => {
  await mountFunctionLibrary(page)
  try {
    await page.evaluate(() => (window as any).toolbar.select('B2:B2'))
    await page.getByRole('tab', { name: 'Formulas', exact: true }).click()
    await page
      .locator('.workbook-desktop-header')
      .getByRole('button', { name: 'Watch Window', exact: true })
      .click()
    const watch = page.getByRole('dialog', { name: 'Watch Window', exact: true })
    const before = await page.evaluate(() => (window as any).toolbar.snapshot())
    await watch.getByRole('button', { name: 'Add watch', exact: true }).click()
    await expect(watch.getByRole('cell', { name: '120', exact: true })).toBeVisible()
    expect(await page.evaluate(() => (window as any).toolbar.snapshot())).toEqual(before)
    await page.evaluate(() => (window as any).toolbar.putFormula('B2', '=1+2'))
    await expect(watch.locator('.watch-formula')).toHaveText('=1+2')
    await expect(watch.getByRole('cell', { name: '…', exact: true })).toBeVisible()
    await page.locator('#ribbon-tab-data').click()
    await expect(watch).toBeVisible()
    await watch.locator('.watch-remove').click()
    await expect(watch.locator('tbody tr')).toHaveCount(0)
    await watch.locator('.watch-close').click()
    await expect(watch).toHaveCount(0)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('shared Formula Auditing waits for an in-flight scan during Save and rejects an overtaken source query', async ({
  page,
}) => {
  await mountFunctionLibrary(page)
  try {
    await page.evaluate(async () => {
      await (window as any).toolbar.seedAuditSource()
      await (window as any).toolbar.select('A1:A1')
      ;(window as any).toolbar.holdReads()
    })
    const check = page
      .locator('.workbook-desktop-header')
      .getByRole('button', { name: 'Error Checking', exact: true })
    await check.click()
    await check.click()
    await page.evaluate(() => {
      ;(window as any).auditSave = 'waiting'
      void (window as any).toolbar.flush().then((ok: boolean) => {
        ;(window as any).auditSave = ok ? 'ready' : 'failed'
      })
    })
    expect(await page.evaluate(() => (window as any).auditSave)).toBe('waiting')
    await page.evaluate(() => (window as any).toolbar.releaseReads())
    await expect.poll(() => page.evaluate(() => (window as any).auditSave)).toBe('ready')
    await expect
      .poll(() => page.evaluate(() => (window as any).toolbar.selection()))
      .toMatchObject({ selection: { address: 'D100001:D100001' } })
    await page.evaluate(() => {
      ;(window as any).toolbar.select('A1:A1')
      ;(window as any).toolbar.holdReads()
    })
    await check.click()
    const revision = await page.evaluate(() => (window as any).toolbar.revision())
    await page.evaluate(() => {
      void (window as any).toolbar.resizeSpecialSource()
    })
    await expect
      .poll(() => page.evaluate(() => (window as any).toolbar.revision()))
      .toBe(revision + 1)
    await page.evaluate(() => (window as any).toolbar.releaseReads())
    await expect(page.getByRole('alert')).toContainText('ready')
    await expect
      .poll(() => page.evaluate(() => (window as any).toolbar.selection()))
      .toMatchObject({ selection: { address: 'A1:A1' } })
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('shared Show Formulas retains the current view on rejected writes and Save awaits the accepted retry', async ({
  page,
}) => {
  await mountFunctionLibrary(page)
  try {
    const before = await page.evaluate(() => (window as any).toolbar.snapshot())
    await page.evaluate(() => (window as any).toolbar.rejectNextSubmission('view-retry'))
    const toggle = page
      .locator('.workbook-desktop-header')
      .getByRole('button', { name: 'Show Formulas', exact: true })
    await toggle.click()
    await expect(page.getByRole('alert')).toContainText('view-retry')
    expect(await page.evaluate(() => (window as any).toolbar.snapshot())).toEqual(before)
    await page.evaluate(() => (window as any).toolbar.holdSubmissions())
    await toggle.click()
    await page.evaluate(() => {
      ;(window as any).auditViewSave = 'waiting'
      void (window as any).toolbar.flush().then((ok: boolean) => {
        ;(window as any).auditViewSave = ok ? 'ready' : 'failed'
      })
    })
    expect(await page.evaluate(() => (window as any).auditViewSave)).toBe('waiting')
    await page.evaluate(() => (window as any).toolbar.releaseSubmissions())
    await expect.poll(() => page.evaluate(() => (window as any).auditViewSave)).toBe('ready')
    await expect
      .poll(
        async () =>
          (await page.evaluate(() => (window as any).toolbar.snapshot())).sheets[0].features
            .pageSetupEdits.showFormulas,
      )
      .toBe(true)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('shared Show Formulas works on protected worksheets without granting cell-write permission', async ({
  page,
}) => {
  await mountFunctionLibrary(page)
  try {
    await page.evaluate(() => (window as any).toolbar.protectViewSheet())
    const before = await page.evaluate(() => (window as any).toolbar.snapshot())
    const toggle = page
      .locator('.workbook-desktop-header')
      .getByRole('button', { name: 'Show Formulas', exact: true })
    await toggle.click()
    await expect
      .poll(
        async () =>
          (await page.evaluate(() => (window as any).toolbar.snapshot())).sheets[0].features
            .pageSetupEdits?.showFormulas,
      )
      .toBe(true)
    await toggle.click()
    await expect
      .poll(
        async () =>
          (await page.evaluate(() => (window as any).toolbar.snapshot())).sheets[0].features
            .pageSetupEdits?.showFormulas,
      )
      .toBe(false)
    const after = await page.evaluate(() => (window as any).toolbar.snapshot())
    expect(after.sheets[0].protected).toBe(true)
    expect(after.sheets[0].cells).toEqual(before.sheets[0].cells)
    const error = await page.evaluate(async () => {
      try {
        await (window as any).toolbar.putFormula('B2', '=99')
        return ''
      } catch (error) {
        return String(error)
      }
    })
    expect(error).toContain('protected_target')
    expect(await page.evaluate(() => (window as any).toolbar.snapshot())).toEqual(after)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('Consolidate keeps rejected drafts, commits label formulas once and rejects stale dialogs', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar()
    ;(window as any).toolbar.select('D1')
  })
  try {
    await page.locator('#ribbon-tab-data').click()
    await page.getByRole('button', { name: 'Consolidate', exact: true }).click()
    const dialog = page.getByRole('dialog', { name: 'Consolidate', exact: true })
    await dialog.locator('input:not([type=checkbox])').fill('Data!A2:B3')
    await dialog.getByRole('checkbox').check()
    await page.evaluate(() => (window as any).toolbar.rejectNextSubmission('retry consolidation'))
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog.getByRole('alert')).toContainText('retry consolidation')
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(0)
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog).toHaveCount(0)
    const result = await page.evaluate(() => ({
      revision: (window as any).toolbar.revision(),
      cells: (window as any).toolbar.snapshot().sheets[0].cells,
    }))
    expect(result.revision).toBe(1)
    expect(result.cells.D1.rawValue).toBe('North')
    expect(result.cells.D2.rawValue).toBe('South')
    expect(result.cells.E1.formula).toBe('=SUMIF(Data!$A$2:$A$3,$D1,Data!B2:B3)')
    await page.evaluate(() => (window as any).toolbar.select('G1'))
    await page.getByRole('button', { name: 'Consolidate', exact: true }).click()
    await dialog.locator('input:not([type=checkbox])').fill('Data!B2:B3')
    await page.evaluate(() => (window as any).toolbar.edit())
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog.getByRole('alert')).toContainText('document changed')
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(2)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('shared Custom Sort retains a rejected draft and commits selected levels once', async ({
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
    await page.locator('#ribbon-tab-data').click()
    await page
      .locator('.workbook-shared-sort-filter')
      .getByRole('button', { name: 'Sort', exact: true })
      .click()
    await page.getByRole('option', { name: 'Custom Sort…', exact: true }).click()
    const dialog = page.locator('.sort-dialog')
    await dialog.getByRole('button', { name: 'Sort by', exact: true }).click()
    await page.locator('[role=option][data-value="1"]').click()
    await dialog.locator('.sort-level').nth(1).getByRole('button').first().click()
    await page.locator('[role=option][data-value="0"]').click()
    await page.evaluate(() => (window as any).toolbar.rejectNextSubmission('retry custom sort'))
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog.getByRole('alert')).toContainText('retry custom sort')
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(0)
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog).toHaveCount(0)
    const result = await page.evaluate(() => ({
      revision: (window as any).toolbar.revision(),
      cells: (window as any).toolbar.snapshot().sheets[0].cells,
    }))
    expect(result.revision).toBe(1)
    expect(result.cells.A1.rawValue).toBe('Region')
    expect(result.cells.A2.rawValue).toBe('South')
    expect(result.cells.B2.rawValue).toBe(-30)
    expect(result.cells.A3.rawValue).toBe('North')
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('outline detail controls project nested visibility and clear both axes atomically', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar()
  })
  const command = async (selection: string, name: string, revision: number) => {
    await page.evaluate((selection) => (window as any).toolbar.select(selection), selection)
    await page.locator('#ribbon-tab-data').click()
    await pickOutlineCommand(page, name)
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.revision())).toBe(revision)
  }
  try {
    await command('A2:A5', 'Group Rows', 1)
    await command('A3:A4', 'Group Rows', 2)
    await command('A3', 'Hide Detail (Rows)', 3)
    await command('A2', 'Hide Detail (Rows)', 4)
    await command('A2', 'Show Detail (Rows)', 5)
    expect(
      await page.evaluate(() => (window as any).toolbar.axisVisibility('rows', [1, 2, 3, 4])),
    ).toEqual([true, false, false, true])
    await command('B1:D1', 'Group Columns', 6)
    await command('C1', 'Group Columns', 7)
    await command('C1', 'Hide Detail (Columns)', 8)
    await command('B1', 'Hide Detail (Columns)', 9)
    await command('B1', 'Show Detail (Columns)', 10)
    expect(
      await page.evaluate(() => (window as any).toolbar.axisVisibility('cols', [1, 2, 3])),
    ).toEqual([true, false, true])
    await command('A1', 'Clear Outline', 11)
    expect(
      await page.evaluate(() => (window as any).toolbar.axisVisibility('rows', [1, 2, 3, 4])),
    ).toEqual([true, true, true, true])
    expect(
      await page.evaluate(() => (window as any).toolbar.axisVisibility('cols', [1, 2, 3])),
    ).toEqual([true, true, true])
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('shared desktop outline gutters fold exact groups and apply depth buttons', async ({
  page,
}) => {
  await page.setViewportSize({ width: 591, height: 365 })
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar()
  })
  const revision = (expected: number) =>
    expect.poll(() => page.evaluate(() => (window as any).toolbar.revision())).toBe(expected)
  const command = async (selection: string, name: string, expected: number) => {
    await page.evaluate((selection) => (window as any).toolbar.select(selection), selection)
    await page.locator('#ribbon-tab-data').click()
    await pickOutlineCommand(page, name)
    await revision(expected)
  }
  const groupButton = async (axis: 'rows' | 'cols', summary: string, level: number) => {
    await page.evaluate((summary) => (window as any).toolbar.reveal(summary), summary)
    await expect(page.locator(`.outline-gutter-${axis}`)).toBeVisible()
    const point = await page.evaluate(
      ({ axis, summary, level }) => {
        const gutter = document.querySelector(`.outline-gutter-${axis}`)!.getBoundingClientRect()
        const index =
          axis === 'rows' ? Number(summary.replace(/^[A-Z]+/, '')) - 1 : summary.charCodeAt(0) - 65
        const position = (window as any).toolbar.outlinePoint(axis, index)
        const along = 3 + (level - 1) * 13 + 6.5
        return axis === 'rows'
          ? { x: gutter.left + along, y: position.y }
          : { x: position.x, y: gutter.top + along }
      },
      { axis, summary, level },
    )
    await page.mouse.click(point.x, point.y)
  }
  const depthButton = async (axis: 'rows' | 'cols', level: number) => {
    await expect(page.locator(`.outline-gutter-${axis}`)).toBeVisible()
    const point = await page.evaluate(
      ({ axis, level }) => {
        const gutter = document.querySelector(`.outline-gutter-${axis}`)!.getBoundingClientRect()
        const grid = [...document.querySelectorAll('.workbook-grid canvas')]
          .map((node) => node.getBoundingClientRect())
          .sort((a, b) => b.width * b.height - a.width * a.height)[0]!
        const along = 3 + (level - 1) * 13 + 6.5
        const columnBand =
          document.querySelector('.outline-gutter-cols')?.getBoundingClientRect().height ?? 0
        const rowDepthTop = Math.max(0, (20 + columnBand - 19) / 2)
        return axis === 'rows'
          ? { x: gutter.left + along, y: grid.top + rowDepthTop + 9.5 }
          : { x: gutter.left + 9.5, y: gutter.top + along }
      },
      { axis, level },
    )
    await page.mouse.click(point.x, point.y)
  }
  try {
    expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true)
    await command('A2:A5', 'Group Rows', 1)
    await command('A3:A4', 'Group Rows', 2)
    await groupButton('rows', 'A5', 2)
    await revision(3)
    expect(
      await page.evaluate(() => (window as any).toolbar.axisVisibility('rows', [1, 2, 3, 4])),
    ).toEqual([true, false, false, true])
    await depthButton('rows', 1)
    await revision(4)
    expect(
      await page.evaluate(() => (window as any).toolbar.axisVisibility('rows', [1, 2, 3, 4])),
    ).toEqual([false, false, false, false])
    await depthButton('rows', 3)
    await revision(5)
    expect(
      await page.evaluate(() => (window as any).toolbar.axisVisibility('rows', [1, 2, 3, 4])),
    ).toEqual([true, true, true, true])
    await command('B1:D1', 'Group Columns', 6)
    await command('C1', 'Group Columns', 7)
    await groupButton('cols', 'D1', 2)
    await revision(8)
    expect(
      await page.evaluate(() => (window as any).toolbar.axisVisibility('cols', [1, 2, 3])),
    ).toEqual([true, false, true])
    await depthButton('cols', 1)
    await revision(9)
    expect(
      await page.evaluate(() => (window as any).toolbar.axisVisibility('cols', [1, 2, 3])),
    ).toEqual([false, false, false])
    await depthButton('cols', 3)
    await revision(10)
    expect(
      await page.evaluate(() => (window as any).toolbar.axisVisibility('cols', [1, 2, 3])),
    ).toEqual([true, true, true])
    await depthButton('rows', 1)
    await revision(11)
    expect(
      await page.evaluate(() => (window as any).toolbar.axisVisibility('rows', [1, 2, 3, 4])),
    ).toEqual([false, false, false, false])
    await command('A1', 'Clear Outline', 12)
    await expect(page.locator('.outline-gutter-rows')).toHaveCount(0)
    await expect(page.locator('.outline-gutter-cols')).toHaveCount(0)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('shared Remove Duplicates retains options on retry and waits for pending atomic edits', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar()
    await (window as any).toolbar.seedDedupeSource()
    ;(window as any).toolbar.select('A1:B5')
  })
  try {
    await page.locator('#ribbon-tab-data').click()
    await page.getByRole('button', { name: 'Remove Duplicates', exact: true }).click()
    const dialog = page.getByRole('dialog', { name: 'Remove Duplicates', exact: true })
    const headers = dialog.getByRole('checkbox', { name: 'My data has headers', exact: true })
    await expect(headers).toBeChecked()
    await headers.uncheck()
    await page.evaluate(() =>
      (window as any).toolbar.rejectNextSubmission('retry duplicate removal'),
    )
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog.getByRole('alert')).toContainText('retry duplicate removal')
    await expect(headers).not.toBeChecked()
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(0)
    await page.evaluate(() => (window as any).toolbar.holdSubmissions())
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog.getByRole('button', { name: 'OK', exact: true })).toBeDisabled()
    await expect(dialog.getByRole('button', { name: 'Cancel', exact: true })).toBeDisabled()
    await page.keyboard.press('Escape')
    await expect(dialog).toBeVisible()
    await page.evaluate(() => {
      ;(window as any).dedupeFlush = 'waiting'
      void (window as any).toolbar.flush().then((ok: boolean) => {
        ;(window as any).dedupeFlush = ok ? 'ready' : 'failed'
      })
    })
    expect(await page.evaluate(() => (window as any).dedupeFlush)).toBe('waiting')
    await page.evaluate(() => (window as any).toolbar.releaseSubmissions())
    await expect(dialog).toHaveCount(0)
    await expect.poll(() => page.evaluate(() => (window as any).dedupeFlush)).toBe('ready')
    const result = await page.evaluate(() => ({
      revision: (window as any).toolbar.revision(),
      cells: (window as any).toolbar.snapshot().sheets[0].cells,
    }))
    expect(result.revision).toBe(1)
    expect(result.cells.B1.formula).toBe('=60*2')
    expect(result.cells.B2).toMatchObject({ kind: 'boolean', rawValue: false })
    expect(result.cells.B3).toMatchObject({ kind: 'string', rawValue: 'false' })
    expect(result.cells.A3.style.fillColor).toBe('#CCEEFF')
    expect(result.cells.A4.rawValue).toBeNull()
    expect(result.cells.B5.rawValue).toBeNull()
    await expect(page.getByRole('status')).toContainText('Removed 2 duplicate row(s).')
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('shared Remove Duplicates cancels without edits and rejects a dialog overtaken by an edit', async ({
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
    await page.locator('#ribbon-tab-data').click()
    const open = page.getByRole('button', { name: 'Remove Duplicates', exact: true })
    await open.click()
    const dialog = page.getByRole('dialog', { name: 'Remove Duplicates', exact: true })
    await dialog.getByRole('button', { name: 'Cancel', exact: true }).click()
    await expect(dialog).toHaveCount(0)
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(0)
    await open.click()
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog).toHaveCount(0)
    await expect(page.getByRole('status')).toContainText('No duplicate rows found')
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(0)
    await open.click()
    await page.evaluate(() => (window as any).toolbar.edit())
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog.getByRole('alert')).toContainText('document changed')
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(1)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('shared Remove Duplicates Save barrier includes reads and rejects an overtaken source', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar()
    await (window as any).toolbar.seedDedupeSource()
    ;(window as any).toolbar.select('A1:B5')
  })
  try {
    await page.locator('#ribbon-tab-data').click()
    await page.getByRole('button', { name: 'Remove Duplicates', exact: true }).click()
    const dialog = page.getByRole('dialog', { name: 'Remove Duplicates', exact: true })
    await page.evaluate(() => (window as any).toolbar.holdReads())
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog.getByRole('button', { name: 'OK', exact: true })).toBeDisabled()
    await page.evaluate(() => {
      ;(window as any).dedupeFlush = 'waiting'
      void (window as any).toolbar.flush().then((ok: boolean) => {
        ;(window as any).dedupeFlush = ok ? 'ready' : 'failed'
      })
    })
    expect(await page.evaluate(() => (window as any).dedupeFlush)).toBe('waiting')
    expect(await page.evaluate(() => (window as any).toolbar.submissions())).toBe(0)
    await page.evaluate(() => {
      void (window as any).toolbar.edit()
    })
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.revision())).toBe(1)
    await page.evaluate(() => (window as any).toolbar.releaseReads())
    await expect(dialog.getByRole('alert')).toContainText('document changed')
    await expect.poll(() => page.evaluate(() => (window as any).dedupeFlush)).toBe('failed')
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(1)
    expect(
      await page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells.A4.rawValue),
    ).toBe('south')
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('shared Remove Duplicates refuses merged selections before producing writes', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(false, true)
    ;(window as any).toolbar.select('A1:B3')
  })
  try {
    await page.locator('#ribbon-tab-data').click()
    await page.getByRole('button', { name: 'Remove Duplicates', exact: true }).click()
    const dialog = page.getByRole('dialog', { name: 'Remove Duplicates', exact: true })
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog.getByRole('alert')).toContainText('merged cells')
    expect(await page.evaluate(() => (window as any).toolbar.submissions())).toBe(0)
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(0)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('shared Advanced Filter validates, retains failed conditions and commits a pending AND filter once', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar()
    await (window as any).toolbar.createFilter()
  })
  try {
    await page.locator('#ribbon-tab-data').click()
    await page.getByRole('button', { name: 'Advanced', exact: true }).click()
    const dialog = page.getByRole('dialog', { name: 'Advanced Filter', exact: true })
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog.getByRole('alert')).toContainText('Enter a value')
    await dialog.getByRole('button', { name: 'Column', exact: true }).click()
    await page.locator('[role=option][data-value="1"]').click()
    await dialog.getByRole('button', { name: 'Condition 1', exact: true }).click()
    await page.locator('[role=option][data-value="greaterThanOrEqual"]').click()
    await dialog.getByRole('textbox', { name: 'Value', exact: true }).nth(0).fill('10')
    await dialog.getByRole('button', { name: 'Condition 2 (optional)', exact: true }).click()
    await page.locator('[role=option][data-value="lessThan"]').click()
    await dialog.getByRole('textbox', { name: 'Value', exact: true }).nth(1).fill('200')
    await page.evaluate(() => (window as any).toolbar.rejectNextSubmission('retry advanced filter'))
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog.getByRole('alert')).toContainText('retry advanced filter')
    await expect(dialog.getByRole('textbox', { name: 'Value', exact: true }).nth(0)).toHaveValue(
      '10',
    )
    await expect(dialog.getByRole('radio', { name: 'And', exact: true })).toBeChecked()
    await page.evaluate(() => (window as any).toolbar.holdSubmissions())
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog.getByRole('button', { name: 'OK', exact: true })).toBeDisabled()
    await expect(dialog.getByRole('button', { name: 'Cancel', exact: true })).toBeDisabled()
    await page.keyboard.press('Escape')
    await expect(dialog).toBeVisible()
    await page.evaluate(() => {
      ;(window as any).advancedFlush = 'waiting'
      void (window as any).toolbar.flush().then((ok: boolean) => {
        ;(window as any).advancedFlush = ok ? 'ready' : 'failed'
      })
    })
    expect(await page.evaluate(() => (window as any).advancedFlush)).toBe('waiting')
    await page.evaluate(() => (window as any).toolbar.releaseSubmissions())
    await expect(dialog).toHaveCount(0)
    await expect.poll(() => page.evaluate(() => (window as any).advancedFlush)).toBe('ready')
    const result = await page.evaluate(() => ({
      revision: (window as any).toolbar.revision(),
      state: (window as any).toolbar.snapshot(),
      visible: (window as any).toolbar.axisVisibility('rows', [1, 2]),
    }))
    expect(result.revision).toBe(2)
    expect(result.visible).toEqual([true, false])
    expect(result.state.sheets[0].features.filterEdits).toMatchObject({
      hiddenRows: [2],
      filter: {
        columns: [
          {
            colId: 1,
            customs: {
              and: true,
              filters: [
                { operator: 'greaterThanOrEqual', val: '10' },
                { operator: 'lessThan', val: '200' },
              ],
            },
          },
        ],
      },
    })
    expect(result.state.sheets[0].cells.B2.rawValue).toBe(120)
  } finally {
    if (!page.isClosed()) await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('shared Advanced Filter supports OR wildcards and refuses stale dialogs', async ({ page }) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar()
    await (window as any).toolbar.createFilter()
  })
  try {
    await page.locator('#ribbon-tab-data').click()
    const open = page.getByRole('button', { name: 'Advanced', exact: true })
    await open.click()
    const dialog = page.getByRole('dialog', { name: 'Advanced Filter', exact: true })
    await dialog.getByRole('textbox', { name: 'Value', exact: true }).nth(0).fill('N*')
    await dialog.getByRole('textbox', { name: 'Value', exact: true }).nth(1).fill('S*')
    await dialog.getByRole('radio', { name: 'Or', exact: true }).check()
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog).toHaveCount(0)
    expect(
      await page.evaluate(() => (window as any).toolbar.axisVisibility('rows', [1, 2])),
    ).toEqual([true, true])
    expect(
      await page.evaluate(
        () =>
          (window as any).toolbar.snapshot().sheets[0].features.filterEdits.filter.columns[0]
            .customs,
      ),
    ).toEqual({ and: false, filters: [{ val: 'N*' }, { val: 'S*' }] })
    await open.click()
    await dialog.getByRole('button', { name: 'Cancel', exact: true }).click()
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(2)
    await open.click()
    await dialog.getByRole('textbox', { name: 'Value', exact: true }).nth(0).fill('N*')
    await page.evaluate(() => (window as any).toolbar.edit())
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog.getByRole('alert')).toContainText('document changed')
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(3)
  } finally {
    if (!page.isClosed()) await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('shared desktop Sort & Filter group toggles smart source regions and clears criteria without removing the filter', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar()
    ;(window as any).toolbar.select('B2')
  })
  try {
    await page.locator('#ribbon-tab-data').click()
    const band = page.locator('.workbook-shared-sort-filter')
    await band.getByRole('button', { name: 'Sort', exact: true }).click()
    expect(await page.getByRole('option').allTextContents()).toEqual([
      'Sort A → Z',
      'Sort Z → A',
      'Custom Sort…',
    ])
    await page.keyboard.press('Escape')
    const toggle = band.getByRole('button', { name: /^Filter\b/ })
    await expect(band.getByRole('button', { name: 'Sort', exact: true })).toBeFocused()
    await page.keyboard.press('Control+Shift+l')
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.revision())).toBe(1)
    expect(
      await page.evaluate(
        () => (window as any).toolbar.snapshot().sheets[0].features.filterEdits.filter.range,
      ),
    ).toEqual({ startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 })
    await band.getByRole('button', { name: 'Advanced', exact: true }).click()
    const dialog = page.getByRole('dialog', { name: 'Advanced Filter', exact: true })
    await page.keyboard.press('Control+Shift+l')
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(1)
    await dialog.getByRole('textbox', { name: 'Value', exact: true }).nth(0).fill('N*')
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog).toHaveCount(0)
    expect(
      await page.evaluate(() => (window as any).toolbar.axisVisibility('rows', [1, 2])),
    ).toEqual([true, false])
    await band.getByRole('button', { name: 'Clear', exact: true }).click()
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.revision())).toBe(3)
    expect(
      await page.evaluate(
        () => (window as any).toolbar.snapshot().sheets[0].features.filterEdits.filter.columns,
      ),
    ).toEqual([])
    expect(
      await page.evaluate(() => (window as any).toolbar.axisVisibility('rows', [1, 2])),
    ).toEqual([true, true])
    await band.getByRole('button', { name: 'Clear', exact: true }).click()
    await page.evaluate(() => (window as any).toolbar.flush())
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(3)
    await toggle.click()
    await expect
      .poll(() =>
        page.evaluate(
          () => (window as any).toolbar.snapshot().sheets[0].features.filterEdits.filter,
        ),
      )
      .toBeNull()
    await toggle.click()
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.revision())).toBe(5)
  } finally {
    if (!page.isClosed()) await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('shared Reapply reads edited source values and rejects a read overtaken by a concurrent edit', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar()
    await (window as any).toolbar.createFilter()
  })
  try {
    await page.locator('#ribbon-tab-data').click()
    const band = page.locator('.workbook-shared-sort-filter')
    await band.getByRole('button', { name: 'Advanced', exact: true }).click()
    const dialog = page.getByRole('dialog', { name: 'Advanced Filter', exact: true })
    await dialog.getByRole('button', { name: 'Column', exact: true }).click()
    await page.locator('[role=option][data-value="1"]').click()
    await dialog.getByRole('button', { name: 'Condition 1', exact: true }).click()
    await page.locator('[role=option][data-value="greaterThan"]').click()
    await dialog.getByRole('textbox', { name: 'Value', exact: true }).nth(0).fill('0')
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog).toHaveCount(0)
    await page.evaluate(() => (window as any).toolbar.editValue('B3', 30))
    expect(
      await page.evaluate(() => (window as any).toolbar.axisVisibility('rows', [1, 2])),
    ).toEqual([true, false])
    await band.getByRole('button', { name: 'Reapply', exact: true }).click()
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.revision())).toBe(4)
    expect(
      await page.evaluate(() => (window as any).toolbar.axisVisibility('rows', [1, 2])),
    ).toEqual([true, true])
    await page.evaluate(() => (window as any).toolbar.holdReads())
    await band.getByRole('button', { name: 'Reapply', exact: true }).click()
    await page.evaluate(() => {
      ;(window as any).filterFlush = 'waiting'
      void (window as any).toolbar.flush().then((ok: boolean) => {
        ;(window as any).filterFlush = ok ? 'ready' : 'failed'
      })
      void (window as any).toolbar.editValue('B3', 40)
    })
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.revision())).toBe(5)
    expect(await page.evaluate(() => (window as any).filterFlush)).toBe('waiting')
    await page.evaluate(() => (window as any).toolbar.releaseReads())
    await expect(page.getByRole('alert')).toContainText('document changed')
    await expect.poll(() => page.evaluate(() => (window as any).filterFlush)).toBe('failed')
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(5)
  } finally {
    if (!page.isClosed()) await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('native column-header filter popup applies host-backed values and clears filtered rows', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar()
    await (window as any).toolbar.createFilter()
  })
  const open = async () => {
    await expect
      .poll(() =>
        page.evaluate(() =>
          [...document.querySelectorAll('canvas')].some(
            (element) => element.getBoundingClientRect().height > 100,
          ),
        ),
      )
      .toBe(true)
    const position = await page.evaluate(() => {
      const canvas = [...document.querySelectorAll('canvas')].find(
        (element) => element.getBoundingClientRect().height > 100,
      )!
      const box = canvas.getBoundingClientRect()
      const cell = (window as any).toolbar.cellRect('A1')
      return { x: box.x + cell.right - 9, y: box.y + cell.bottom - 9 }
    })
    await page.mouse.click(position.x, position.y)
    await expect(
      page
        .locator('[data-u-comp="sheets-filter-panel"]')
        .filter({ has: page.locator('[data-u-comp="sheets-filter-panel-footer"]') }),
    ).toBeVisible()
  }
  try {
    await open()
    const panel = page
      .locator('[data-u-comp="sheets-filter-panel"]')
      .filter({ has: page.locator('[data-u-comp="sheets-filter-panel-footer"]') })
    await expect(panel.getByText('North', { exact: true })).toBeVisible()
    await expect(panel.getByText('South', { exact: true })).toBeVisible()
    await panel
      .locator('[data-u-comp="sheets-filter-panel-values-virtual"] label[data-u-comp="checkbox"]')
      .nth(1)
      .click()
    await page.evaluate(() => (window as any).toolbar.rejectNextSubmission('Filter commit failed'))
    await panel.getByRole('button', { name: 'Confirm', exact: true }).click()
    await expect(page.getByRole('alert')).toContainText('Filter commit failed')
    await expect(panel).toBeVisible()
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(1)
    expect(
      await panel
        .locator('[data-u-comp="sheets-filter-panel-values-virtual"] input')
        .nth(1)
        .isChecked(),
    ).toBe(false)
    await page.evaluate(() => (window as any).toolbar.holdSubmissions())
    await panel.getByRole('button', { name: 'Confirm', exact: true }).click()
    await panel.getByRole('button', { name: 'Confirm', exact: true }).click()
    await page.evaluate(() => {
      ;(window as any).popupFlush = 'waiting'
      void (window as any).toolbar.flush().then((ready: boolean) => {
        ;(window as any).popupFlush = ready ? 'ready' : 'failed'
      })
    })
    expect(await page.evaluate(() => (window as any).popupFlush)).toBe('waiting')
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(1)
    await page.evaluate(() => (window as any).toolbar.releaseSubmissions())
    await expect.poll(() => page.evaluate(() => (window as any).popupFlush)).toBe('ready')
    expect(await page.evaluate(() => (window as any).toolbar.submissions())).toBe(3)
    await page.evaluate(() => (window as any).toolbar.flush())
    await expect(page.getByRole('alert')).toHaveCount(0)
    await expect(panel).toBeHidden()
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.revision())).toBe(2)
    expect(
      await page.evaluate(
        () => (window as any).toolbar.snapshot().sheets[0].features.filterEdits.filter.columns,
      ),
    ).toEqual([{ colId: 0, values: ['North'], blank: false }])
    expect(
      await page.evaluate(() => (window as any).toolbar.axisVisibility('rows', [1, 2])),
    ).toEqual([true, false])
    await open()
    await panel.getByRole('button', { name: /Clear filter/i }).click()
    await expect(panel).toBeHidden()
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.revision())).toBe(3)
    expect(
      await page.evaluate(() => (window as any).toolbar.axisVisibility('rows', [1, 2])),
    ).toEqual([true, true])
  } finally {
    if (!page.isClosed()) await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('native filter popup reads offscreen source values and refuses an overtaken opening read', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(false, false, false, false, false, true)
    await (window as any).toolbar.createFilter('A1:B81')
  })
  const panel = page
    .locator('[data-u-comp="sheets-filter-panel"]')
    .filter({ has: page.locator('[data-u-comp="sheets-filter-panel-footer"]') })
  const clickHeader = async () => {
    await expect
      .poll(() =>
        page.evaluate(() =>
          [...document.querySelectorAll('canvas')].some(
            (element) => element.getBoundingClientRect().height > 100,
          ),
        ),
      )
      .toBe(true)
    const position = await page.evaluate(() => {
      const canvas = [...document.querySelectorAll('canvas')].find(
        (element) => element.getBoundingClientRect().height > 100,
      )!
      const box = canvas.getBoundingClientRect(),
        cell = (window as any).toolbar.cellRect('A1')
      return { x: box.x + cell.right - 9, y: box.y + cell.bottom - 9 }
    })
    await page.mouse.click(position.x, position.y)
  }
  try {
    await clickHeader()
    await expect(panel).toBeVisible()
    await expect(panel.getByText('2022', { exact: true })).toBeVisible()
    expect(
      await page.evaluate(() =>
        (window as any).toolbar
          .reads()
          .some(
            (range: any) => range.startRow === 0 && range.endRow === 80 && range.endColumn === 1,
          ),
      ),
    ).toBe(true)
    await panel.getByRole('button', { name: 'Cancel', exact: true }).click()
    await expect(panel).toBeHidden()
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(1)
    await page.evaluate(() => (window as any).toolbar.holdReads())
    await clickHeader()
    await expect
      .poll(() => page.evaluate(() => (window as any).toolbar.diagnostics().filterPopupWork))
      .toBe(true)
    await page.evaluate(() => {
      void (window as any).toolbar.editValue('B2', 999)
    })
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.revision())).toBe(2)
    await page.evaluate(() => (window as any).toolbar.releaseReads())
    await expect(page.getByRole('alert')).toContainText('document changed')
    await expect(panel).toBeHidden()
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(2)
    await page.evaluate(() => (window as any).toolbar.flush())
    await clickHeader()
    await expect(panel).toBeVisible()
    await expect(page.getByRole('alert')).toHaveCount(0)
    await panel.getByRole('button', { name: 'Cancel', exact: true }).click()
  } finally {
    if (!page.isClosed()) await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('native filter popup uses authoritative formula results and applies desktop conditions', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar()
    await (window as any).toolbar.createFilter()
    await (window as any).toolbar.seedFilterFormulaSource()
  })
  try {
    await expect
      .poll(() =>
        page.evaluate(() =>
          [...document.querySelectorAll('canvas')].some(
            (element) => element.getBoundingClientRect().height > 100,
          ),
        ),
      )
      .toBe(true)
    const position = await page.evaluate(() => {
      const canvas = [...document.querySelectorAll('canvas')].find(
        (element) => element.getBoundingClientRect().height > 100,
      )!
      const box = canvas.getBoundingClientRect(),
        cell = (window as any).toolbar.cellRect('A1')
      return { x: box.x + cell.right - 9, y: box.y + cell.bottom - 9 }
    })
    await page.mouse.click(position.x, position.y)
    const panel = page
      .locator('[data-u-comp="sheets-filter-panel"]')
      .filter({ has: page.locator('[data-u-comp="sheets-filter-panel-footer"]') })
    await expect(panel).toBeVisible()
    await expect(panel.getByText('42', { exact: true })).toBeVisible()
    expect(await page.evaluate(() => (window as any).toolbar.displayedCell('A2').f)).toBe('=Z900')
    await panel.getByText('By Conditions', { exact: true }).click()
    await panel.getByText('By Values', { exact: true }).click()
    await expect(panel.getByText('42', { exact: true })).toBeVisible()
    await panel.getByText('By Conditions', { exact: true }).click()
    const conditions = panel.locator('[data-u-comp="sheets-filter-panel-conditions-container"]')
    await conditions.locator('[data-u-comp="select"]').first().click()
    await page.getByRole('menuitemradio', { name: 'Greater Than', exact: true }).click()
    await conditions.getByRole('textbox').fill('20')
    await panel.getByRole('button', { name: 'Confirm', exact: true }).click()
    await page.evaluate(() => (window as any).toolbar.flush())
    await expect(page.getByRole('alert')).toHaveCount(0)
    await expect(panel).toBeHidden()
    expect(
      await page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].features.filterEdits),
    ).toMatchObject({
      hiddenRows: [2],
      filter: {
        columns: [{ colId: 0, customs: { filters: [{ operator: 'greaterThan', val: '20' }] } }],
      },
    })
  } finally {
    if (!page.isClosed()) await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('native filter popup sorts visible slots with host retry, Save waiting and row styles', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar()
    await (window as any).toolbar.seedVisibleSort()
    await (window as any).toolbar.createFilter('A1:B6')
  })
  const panel = page
    .locator('[data-u-comp="sheets-filter-panel"]')
    .filter({ has: page.locator('[data-u-comp="sheets-filter-panel-footer"]') })
  const open = async (address: string) => {
    await expect
      .poll(() =>
        page.evaluate(() =>
          [...document.querySelectorAll('canvas')].some(
            (element) => element.getBoundingClientRect().height > 100,
          ),
        ),
      )
      .toBe(true)
    const position = await page.evaluate((address) => {
      const canvas = [...document.querySelectorAll('canvas')].find(
        (element) => element.getBoundingClientRect().height > 100,
      )!
      const box = canvas.getBoundingClientRect(),
        cell = (window as any).toolbar.cellRect(address)
      return { x: box.x + cell.right - 9, y: box.y + cell.bottom - 9 }
    }, address)
    await page.mouse.click(position.x, position.y)
    await expect(panel).toBeVisible()
  }
  try {
    await open('A1')
    await panel
      .getByText('Drop', { exact: true })
      .locator('xpath=ancestor::div[label[@data-u-comp="checkbox"]][1]')
      .locator('label[data-u-comp="checkbox"]')
      .click()
    await panel.getByRole('button', { name: 'Confirm', exact: true }).click()
    await expect(panel).toBeHidden()
    await page.evaluate(() => (window as any).toolbar.hideRow(5))
    const revision = await page.evaluate(() => (window as any).toolbar.revision())
    await open('B1')
    await page.evaluate(() => (window as any).toolbar.rejectNextSubmission('Sort failed'))
    await panel.getByRole('button', { name: 'Ascending', exact: true }).click()
    await expect(page.getByRole('alert')).toContainText('Sort failed')
    await expect(panel).toBeVisible()
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(revision)
    await page.evaluate(() => (window as any).toolbar.holdSubmissions())
    await panel.getByRole('button', { name: 'Ascending', exact: true }).click()
    await page.evaluate(() => {
      ;(window as any).sortFlush = 'waiting'
      void (window as any).toolbar.flush().then((ready: boolean) => {
        ;(window as any).sortFlush = ready ? 'ready' : 'failed'
      })
    })
    await expect(panel).toBeVisible()
    expect(await page.evaluate(() => (window as any).sortFlush)).toBe('waiting')
    await page.evaluate(() => (window as any).toolbar.releaseSubmissions())
    await expect.poll(() => page.evaluate(() => (window as any).sortFlush)).toBe('ready')
    await expect(panel).toBeHidden()
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(revision + 1)
    expect(
      await page.evaluate(() => {
        const cells = (window as any).toolbar.snapshot().sheets[0].cells
        return ['B2', 'B3', 'B4', 'B5', 'B6'].map((address) => cells[address].rawValue)
      }),
    ).toEqual([10, 999, 20, 5, 30])
    expect(
      await page.evaluate(
        () => (window as any).toolbar.snapshot().sheets[0].cells.B2.style.fillColor,
      ),
    ).toBe('#FF0000')
    expect(
      await page.evaluate(() => (window as any).toolbar.axisVisibility('rows', [1, 2, 3, 4, 5])),
    ).toEqual([true, false, true, false, true])
    await open('B1')
    await panel.getByRole('button', { name: 'Descending', exact: true }).click()
    await page.evaluate(() => (window as any).toolbar.flush())
    await expect(panel).toBeHidden()
    expect(
      await page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells.B2.rawValue),
    ).toBe(30)
  } finally {
    if (!page.isClosed()) await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('native color filter supports fill/font, multiple choices, no fill, retry and reapply after a style edit', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar()
    await (window as any).toolbar.seedFilterColors()
    await (window as any).toolbar.createFilter('A1:A5')
  })
  const panel = page
    .locator('[data-u-comp="sheets-filter-panel"]')
    .filter({ has: page.locator('[data-u-comp="sheets-filter-panel-footer"]') })
  const colors = panel.locator('[data-u-comp="sheets-filter-panel-colors-container"]')
  const open = async () => {
    await expect
      .poll(() =>
        page.evaluate(() =>
          [...document.querySelectorAll('canvas')].some(
            (element) => element.getBoundingClientRect().height > 100,
          ),
        ),
      )
      .toBe(true)
    const point = await page.evaluate(() => {
      const canvas = [...document.querySelectorAll('canvas')].find(
        (element) => element.getBoundingClientRect().height > 100,
      )!
      const box = canvas.getBoundingClientRect(),
        cell = (window as any).toolbar.cellRect('A1')
      return { x: box.x + cell.right - 9, y: box.y + cell.bottom - 9 }
    })
    await page.mouse.click(point.x, point.y)
    await expect(panel).toBeVisible()
    if (!(await colors.count())) await panel.getByText('By Colors', { exact: true }).click()
    await expect(colors).toBeVisible()
  }
  const confirm = async () => {
    await panel.getByRole('button', { name: 'Confirm', exact: true }).click()
    await page.evaluate(() => (window as any).toolbar.flush())
    await expect(page.getByRole('alert')).toHaveCount(0)
    await expect(panel).toBeHidden()
  }
  try {
    await open()
    await colors.locator('button[style*="255, 0, 0"]').click()
    await page.evaluate(() => (window as any).toolbar.rejectNextSubmission('Color filter failed'))
    await panel.getByRole('button', { name: 'Confirm', exact: true }).click()
    await expect(page.getByRole('alert')).toContainText('Color filter failed')
    await expect(panel).toBeVisible()
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(1)
    await confirm()
    expect(
      await page.evaluate(() => (window as any).toolbar.axisVisibility('rows', [1, 2, 3, 4])),
    ).toEqual([true, false, false, false])
    await page.evaluate(() => (window as any).toolbar.editFill('A3', '#FF0000'))
    await page.locator('#ribbon-tab-data').click()
    await page
      .locator('.workbook-shared-sort-filter')
      .getByRole('button', { name: 'Reapply', exact: true })
      .click()
    await page.evaluate(() => (window as any).toolbar.flush())
    expect(
      await page.evaluate(() => (window as any).toolbar.axisVisibility('rows', [1, 2, 3, 4])),
    ).toEqual([true, true, false, false])
    await open()
    await colors.locator('.univerjs-icon-ban-icon').click()
    await confirm()
    expect(
      await page.evaluate(
        () =>
          (window as any).toolbar.snapshot().sheets[0].features.filterEdits.filter.columns[0]
            .colors,
      ),
    ).toEqual({ fill: ['#FF0000', null] })
    expect(
      await page.evaluate(() => (window as any).toolbar.axisVisibility('rows', [1, 2, 3, 4])),
    ).toEqual([true, true, true, false])
    await open()
    await colors
      .locator('.univer-grid-cols-8')
      .nth(1)
      .locator('[style*="color: rgb(0, 0, 255)"]')
      .click()
    await confirm()
    expect(
      await page.evaluate(
        () =>
          (window as any).toolbar.snapshot().sheets[0].features.filterEdits.filter.columns[0]
            .colors,
      ),
    ).toEqual({ font: ['#0000FF'] })
    expect(
      await page.evaluate(() => (window as any).toolbar.axisVisibility('rows', [1, 2, 3, 4])),
    ).toEqual([true, false, false, true])
    await open()
    await panel.getByRole('button', { name: 'Clear Filter', exact: true }).click()
    await page.evaluate(() => (window as any).toolbar.flush())
    await expect(panel).toBeHidden()
    expect(
      await page.evaluate(() => (window as any).toolbar.axisVisibility('rows', [1, 2, 3, 4])),
    ).toEqual([true, true, true, true])
  } finally {
    if (!page.isClosed()) await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('shared Data Tools menu splits four delimiters at non-A offsets with sparse blanks and no inserted columns', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(true)
  })
  try {
    await page.locator('#ribbon-tab-data').click()
    for (const [separator, name] of [
      [',', 'Comma'],
      [';', 'Semicolon'],
      [' ', 'Space'],
      ['\t', 'Tab'],
    ]) {
      await page.evaluate(
        (separator) => (window as any).toolbar.seedTextSplit(`left${separator}right`),
        separator,
      )
      const before = await page.evaluate(() => (window as any).toolbar.revision())
      await page.getByRole('button', { name: 'Text to Columns', exact: true }).click()
      await page.getByRole('option', { name: `Split by ${name}`, exact: true }).click()
      await expect
        .poll(() => page.evaluate(() => (window as any).toolbar.revision()))
        .toBe(before + 1)
      const state = await page.evaluate(() => (window as any).toolbar.snapshot())
      expect(state.structuralLog).toEqual([])
      expect(state.sheets[0].cells.A1.formula).toBe('=D4')
      for (const row of [4, 6]) {
        expect(state.sheets[0].cells[`B${row}`].rawValue).toBe('left')
        expect(state.sheets[0].cells[`C${row}`].rawValue).toBe('right')
      }
      expect(state.sheets[0].cells.B5?.rawValue ?? null).toBeNull()
    }
    await expect(page.getByRole('alert')).toHaveCount(0)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('Text to Columns Save waits for source reads and commit while duplicate clicks are rejected', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    const toolbar = ((window as any).toolbar = await mountManualToolbar(true))
    await toolbar.seedTextSplit()
    toolbar.holdReads()
    ;(window as any).splitWork = toolbar.splitText(2)
    ;(window as any).saveDone = false
    ;(window as any).saveWork = toolbar.flush().then((ok: boolean) => {
      ;(window as any).saveDone = true
      return ok
    })
  })
  try {
    await expect(page.evaluate(() => (window as any).toolbar.splitText(2))).rejects.toThrow(
      'already being applied',
    )
    expect(await page.evaluate(() => (window as any).saveDone)).toBe(false)
    await page.evaluate(() => {
      ;(window as any).toolbar.holdSubmissions()
      ;(window as any).toolbar.releaseReads()
    })
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.submissions())).toBe(1)
    expect(await page.evaluate(() => (window as any).saveDone)).toBe(false)
    await page.evaluate(() => (window as any).toolbar.releaseSubmissions())
    expect(await page.evaluate(() => (window as any).saveWork)).toBe(true)
    expect(await page.evaluate(() => (window as any).splitWork)).toBe(4)
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(1)
  } finally {
    await page.evaluate(() => {
      ;(window as any).toolbar.releaseReads()
      ;(window as any).toolbar.releaseSubmissions()
      ;(window as any).toolbar.dispose()
    })
  }
})

test('Text to Columns rejects a failed commit without partial writes and allows retry', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    const toolbar = ((window as any).toolbar = await mountManualToolbar(true))
    await toolbar.seedTextSplit()
  })
  try {
    const before = await page.evaluate(() => (window as any).toolbar.snapshot())
    await page.evaluate(() => {
      const toolbar = (window as any).toolbar
      toolbar.rejectNextSubmission('Text split rejected by host')
      toolbar.holdSubmissions()
      ;(window as any).splitFailure = toolbar.splitText(2).then(
        () => '',
        (error: Error) => error.message,
      )
      ;(window as any).failedSave = toolbar.flush()
    })
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.submissions())).toBe(1)
    await page.evaluate(() => (window as any).toolbar.releaseSubmissions())
    expect(await page.evaluate(() => (window as any).splitFailure)).toContain(
      'Text split rejected by host',
    )
    expect(await page.evaluate(() => (window as any).failedSave)).toBe(false)
    expect(await page.evaluate(() => (window as any).toolbar.snapshot())).toEqual(before)
    await page.evaluate(() => (window as any).toolbar.splitText(2))
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(1)
  } finally {
    await page.evaluate(() => {
      ;(window as any).toolbar.releaseSubmissions()
      ;(window as any).toolbar.dispose()
    })
  }
})

test('shared Forecast and Outline menus retain summary marks after rejection and publish only accepted settings', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(true)
    ;(window as any).toolbar.select('B2')
  })
  try {
    await page.locator('#ribbon-tab-data').click()
    const group = page.getByRole('button', { name: 'Group', exact: true })
    await group.click()
    await expect(
      page.getByRole('option', { name: '✓ Summary rows below detail', exact: true }),
    ).toBeVisible()
    await expect(
      page.getByRole('option', { name: '✓ Summary columns to right of detail', exact: true }),
    ).toBeVisible()
    await page.evaluate(() => {
      ;(window as any).toolbar.holdSubmissions()
      ;(window as any).toolbar.rejectNextSubmission('outline setting rejected')
    })
    await page.getByRole('option', { name: '✓ Summary rows below detail', exact: true }).click()
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.submissions())).toBe(1)
    await page.evaluate(() => {
      ;(window as any).saveFinished = false
      ;(window as any).settingsSave = (window as any).toolbar.flush().then((ok: boolean) => {
        ;(window as any).saveFinished = true
        return ok
      })
    })
    expect(await page.evaluate(() => (window as any).saveFinished)).toBe(false)
    await page.evaluate(() => (window as any).toolbar.releaseSubmissions())
    expect(await page.evaluate(() => (window as any).settingsSave)).toBe(false)
    await expect(page.getByRole('alert')).toContainText('outline setting rejected')
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(0)
    await pickOutlineCommand(page, '✓ Summary rows below detail')
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.revision())).toBe(1)
    await group.click()
    await expect(
      page.getByRole('option', { name: 'Summary rows below detail', exact: true }),
    ).toBeVisible()
    await expect(
      page.getByRole('option', { name: '✓ Summary columns to right of detail', exact: true }),
    ).toBeVisible()
    await page.keyboard.press('Escape')
    await expect(group).toBeFocused()
    await page.getByRole('button', { name: 'What-If Analysis', exact: true }).click()
    await page.getByRole('option', { name: 'Goal Seek', exact: true }).click()
    const dialog = page.getByRole('dialog', { name: 'Goal Seek', exact: true })
    await expect(dialog.getByRole('textbox', { name: 'Set cell', exact: true })).toHaveValue('B2')
    await dialog.getByRole('button', { name: 'Close', exact: true }).click()
    await expect(dialog).toHaveCount(0)
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(1)
  } finally {
    await page.evaluate(() => {
      ;(window as any).toolbar.releaseSubmissions()
      ;(window as any).toolbar.dispose()
    })
  }
})

test('shared Subtotal dialog retains failed options, waits for Save and commits insertions, formulas and outline together', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(true)
    await (window as any).toolbar.seedSubtotal()
  })
  try {
    await page.locator('#ribbon-tab-data').click()
    await page.getByRole('button', { name: 'Subtotal', exact: true }).click()
    const dialog = page.getByRole('dialog', { name: 'Subtotal', exact: true })
    await expect(
      dialog.getByRole('button', { name: 'At each change in', exact: true }),
    ).toContainText('Region')
    await expect(
      dialog.getByRole('button', { name: 'Add subtotal to', exact: true }),
    ).toContainText('Amount')
    await dialog.getByRole('button', { name: 'Use function', exact: true }).click()
    await dialog.getByRole('option', { name: 'Average', exact: true }).click()
    await page.evaluate(() =>
      (window as any).toolbar.rejectNextSubmission('subtotal host rejection'),
    )
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog.getByRole('alert')).toContainText('subtotal host rejection')
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(0)
    await expect(dialog.getByRole('button', { name: 'Use function', exact: true })).toContainText(
      'Average',
    )
    await page.evaluate(() => (window as any).toolbar.holdSubmissions())
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog.getByRole('button', { name: 'Cancel', exact: true })).toBeDisabled()
    await page.keyboard.press('Escape')
    await expect(dialog).toBeVisible()
    await page.evaluate(() => {
      ;(window as any).saveDone = false
      ;(window as any).saveWork = (window as any).toolbar.flush().then((ok: boolean) => {
        ;(window as any).saveDone = true
        return ok
      })
    })
    expect(await page.evaluate(() => (window as any).saveDone)).toBe(false)
    await page.evaluate(() => (window as any).toolbar.releaseSubmissions())
    expect(await page.evaluate(() => (window as any).saveWork)).toBe(true)
    await expect(dialog).toHaveCount(0)
    const state = await page.evaluate(() => (window as any).toolbar.snapshot())
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(1)
    expect(state.structuralLog).toHaveLength(3)
    expect(state.sheets[0].cells.B6.rawValue).toBe('North Total')
    expect(state.sheets[0].cells.C6.formula).toBe('=SUBTOTAL(1,$C$4:$C$5)')
    expect(state.sheets[0].cells.C9.formula).toBe('=SUBTOTAL(1,$C$4:$C$8)')
    expect(state.sheets[0].cells.A11.rawValue).toBe(42)
    await expect(page.locator('.outline-gutter-rows')).toBeVisible()
  } finally {
    await page.evaluate(() => {
      ;(window as any).toolbar.releaseSubmissions()
      ;(window as any).toolbar.dispose()
    })
  }
})

test('shared Subtotal cancels without writes, translates invalid columns and refuses an overtaken draft', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(true)
    await (window as any).toolbar.seedSubtotal()
  })
  try {
    await page.locator('#ribbon-tab-data').click()
    const open = page.getByRole('button', { name: 'Subtotal', exact: true })
    await open.click()
    const dialog = page.getByRole('dialog', { name: 'Subtotal', exact: true })
    await dialog.getByRole('button', { name: 'Cancel', exact: true }).click()
    expect(await page.evaluate(() => (window as any).toolbar.submissions())).toBe(0)
    await open.click()
    await dialog.getByRole('button', { name: 'Add subtotal to', exact: true }).click()
    await dialog.getByRole('option', { name: 'Region', exact: true }).click()
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog.getByRole('alert')).toHaveText('Pick different group and value columns.')
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(0)
    await dialog.getByRole('button', { name: 'Add subtotal to', exact: true }).click()
    await dialog.getByRole('option', { name: 'Amount', exact: true }).click()
    await page.evaluate(() => (window as any).toolbar.editValue('A1', 'Concurrent edit'))
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog.getByRole('alert')).toContainText('document changed')
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(1)
    expect(
      (await page.evaluate(() => (window as any).toolbar.snapshot())).structuralLog,
    ).toHaveLength(0)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('shared Consolidate exposes all desktop functions and retains reference-list rules', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(true)
    ;(window as any).toolbar.select('D1')
  })
  try {
    await page.locator('#ribbon-tab-data').click()
    const open = page.getByRole('button', { name: 'Consolidate', exact: true })
    await open.click()
    const dialog = page.getByRole('dialog', { name: 'Consolidate', exact: true })
    const fn = dialog.getByRole('button', { name: 'Function', exact: true })
    await fn.click()
    await expect(dialog.getByRole('option')).toHaveText(['Sum', 'Count', 'Average', 'Max', 'Min'])
    await dialog.getByRole('option', { name: 'Max', exact: true }).click()
    await dialog.getByRole('checkbox', { name: 'Use labels in left column', exact: true }).check()
    await expect(fn).toContainText('Sum')
    await fn.click()
    await expect(dialog.getByRole('option')).toHaveText(['Sum', 'Count', 'Average'])
    await page.keyboard.press('Escape')
    const reference = dialog.locator('input:not([type=checkbox])')
    await reference.fill('invalid reference')
    await dialog.getByRole('button', { name: 'Add', exact: true }).click()
    await expect(dialog.getByRole('alert')).toContainText('invalid reference')
    await reference.fill('Data!A2:B3')
    await reference.press('Enter')
    const list = dialog.getByRole('list', { name: 'All references', exact: true })
    await expect(list.getByRole('listitem')).toHaveCount(1)
    await reference.fill('Data!A2:B3')
    await dialog.getByRole('button', { name: 'Add', exact: true }).click()
    await expect(list.getByRole('listitem')).toHaveCount(1)
    await list.getByRole('button').click()
    await expect(list).toHaveCount(0)
    await dialog.getByRole('button', { name: 'Cancel', exact: true }).click()
    expect(await page.evaluate(() => (window as any).toolbar.submissions())).toBe(0)
    for (const [index, [name, formula]] of [
      ['Sum', 'SUM'],
      ['Count', 'COUNTA'],
      ['Average', 'AVERAGE'],
      ['Max', 'MAX'],
      ['Min', 'MIN'],
    ].entries()) {
      const target = `${String.fromCharCode(68 + index)}1`
      await page.evaluate((target) => (window as any).toolbar.select(target), target)
      await open.click()
      await reference.fill('Data!B2:B3')
      await fn.click()
      await dialog.getByRole('option', { name, exact: true }).click()
      await dialog.getByRole('button', { name: 'OK', exact: true }).click()
      await expect(dialog).toHaveCount(0)
      expect(
        (await page.evaluate(() => (window as any).toolbar.snapshot())).sheets[0].cells[target]
          .formula,
      ).toBe(`=${formula}(Data!B2)`)
    }
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('shared Consolidate skips sparse blank labels and keeps pending reads and commit inside Save', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(true)
    await (window as any).toolbar.editValue('B4', 42)
    ;(window as any).toolbar.select('D1')
  })
  try {
    await page.locator('#ribbon-tab-data').click()
    await page.getByRole('button', { name: 'Consolidate', exact: true }).click()
    const dialog = page.getByRole('dialog', { name: 'Consolidate', exact: true })
    await dialog.locator('input:not([type=checkbox])').fill('Data!A2:B4')
    await dialog.getByRole('checkbox').check()
    await page.evaluate(() => (window as any).toolbar.holdReads())
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog.getByRole('button', { name: 'Cancel', exact: true })).toBeDisabled()
    await expect(dialog.getByRole('button', { name: 'Function', exact: true })).toBeDisabled()
    await page.keyboard.press('Escape')
    await expect(dialog).toBeVisible()
    await page.evaluate(() => {
      ;(window as any).saveDone = false
      ;(window as any).saveWork = (window as any).toolbar.flush().then((ok: boolean) => {
        ;(window as any).saveDone = true
        return ok
      })
    })
    expect(await page.evaluate(() => (window as any).saveDone)).toBe(false)
    await page.evaluate(() => {
      ;(window as any).toolbar.holdSubmissions()
      ;(window as any).toolbar.releaseReads()
    })
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.submissions())).toBe(2)
    expect(await page.evaluate(() => (window as any).saveDone)).toBe(false)
    await page.evaluate(() => (window as any).toolbar.releaseSubmissions())
    expect(await page.evaluate(() => (window as any).saveWork)).toBe(true)
    await expect(dialog).toHaveCount(0)
    const state = await page.evaluate(() => (window as any).toolbar.snapshot())
    expect(state.sheets[0].cells.D1.rawValue).toBe('North')
    expect(state.sheets[0].cells.D2.rawValue).toBe('South')
    expect(state.sheets[0].cells.E1.formula).toBe('=SUMIF(Data!$A$2:$A$4,$D1,Data!B2:B4)')
    expect(state.sheets[0].cells.D3).toBeUndefined()
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(2)
  } finally {
    await page.evaluate(() => {
      ;(window as any).toolbar.releaseReads()
      ;(window as any).toolbar.releaseSubmissions()
      ;(window as any).toolbar.dispose()
    })
  }
})

test('shared Consolidate refuses pending formula labels without publishing partial output', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    const toolbar = ((window as any).toolbar = await mountManualToolbar(true))
    await toolbar.putFormula('A2', '=B2')
    toolbar.select('D1')
  })
  try {
    const before = await page.evaluate(() => (window as any).toolbar.snapshot())
    await page.locator('#ribbon-tab-data').click()
    await page.getByRole('button', { name: 'Consolidate', exact: true }).click()
    const dialog = page.getByRole('dialog', { name: 'Consolidate', exact: true })
    await dialog.locator('input:not([type=checkbox])').fill('Data!A2:B3')
    await dialog.getByRole('checkbox').check()
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog.getByRole('alert')).toContainText('Wait for source formulas')
    expect(await page.evaluate(() => (window as any).toolbar.snapshot())).toEqual(before)
    await expect(dialog.getByRole('button', { name: 'Cancel', exact: true })).toBeEnabled()
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})
