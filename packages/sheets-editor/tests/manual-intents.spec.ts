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
    await page.getByRole('tab', { name: 'Insert', exact: true }).click()
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
    await page.locator('.workbook-additional-controls > summary').click()
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
    await page.locator('.workbook-additional-controls > summary').click()
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
    await page
      .getByRole('tablist', { name: 'Ribbon', exact: true })
      .getByRole('tab', { name: 'Home', exact: true })
      .focus()
    await page.keyboard.press('ArrowRight')
    await expect(
      page
        .getByRole('tablist', { name: 'Ribbon', exact: true })
        .getByRole('tab', { name: 'Insert', exact: true }),
    ).toBeFocused()
    await tab('Page Layout')
    await page.locator('select[aria-label="Orientation"]').selectOption('landscape')
    await expect
      .poll(async () => (await snapshot()).sheets[0].features.pageSetupEdits.orientation)
      .toBe('landscape')
    await page.getByLabel('Paper size', { exact: true }).selectOption('8')
    await expect
      .poll(async () => (await snapshot()).sheets[0].features.pageSetupEdits.paperSize)
      .toBe(8)
    await page.getByRole('button', { name: 'Set selection as print area', exact: true }).click()
    await expect
      .poll(async () => (await snapshot()).sheets[0].features.pageSetupEdits.printArea)
      .toBe('A1:B3')
    await tab('Formulas')
    await page.evaluate(() => (window as any).toolbar.select('B2:B3'))
    await page.getByRole('button', { name: 'Insert summary below selection', exact: true }).click()
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.value('B4'))).toBe(90)
    await tab('Data')
    await page.evaluate(() => (window as any).toolbar.select('A1:B3'))
    await page.getByRole('button', { name: 'Sort descending', exact: true }).click()
    await expect.poll(async () => (await snapshot()).sheets[0].cells.A2.rawValue).toBe('South')
    expect((await snapshot()).sheets[0].cells.B2.rawValue).toBe(-30)
    await page.getByRole('button', { name: 'Filter selected table', exact: true }).click()
    await expect
      .poll(async () => (await snapshot()).sheets[0].features.filterEdits.filter.range.endRow)
      .toBe(2)
    await page
      .getByRole('textbox', {
        name: 'Keep values (one per line; empty line matches blanks)',
        exact: true,
      })
      .fill('South')
    await page.getByRole('button', { name: 'Apply filter', exact: true }).click()
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.rowVisible(2))).toBe(false)
    await page.getByRole('button', { name: 'Clear filter', exact: true }).click()
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.rowVisible(2))).toBe(true)
    await tab('Review')
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
    await page.getByRole('button', { name: 'Freeze top row', exact: true }).click()
    await expect
      .poll(() => page.evaluate(() => (window as any).toolbar.view().freeze.ySplit))
      .toBe(1)
    await page.getByLabel('Zoom', { exact: true }).selectOption('125')
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.view().zoom)).toBe(1.25)
    await page.getByRole('checkbox', { name: 'Gridlines', exact: true }).click()
    await expect
      .poll(() => page.evaluate(() => (window as any).toolbar.view().gridlines))
      .toBe(false)
    await expect(page.getByRole('checkbox', { name: 'Gridlines', exact: true })).not.toBeChecked()
    await expect(page.getByRole('alert')).toHaveCount(0)
    await tab('Home')
    await page.screenshot({ path: test.info().outputPath('excel-ribbon.png') })
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})
