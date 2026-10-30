import { test, expect } from '@playwright/test'
import { createServer, type ViteDevServer } from 'vite'
import { fileURLToPath } from 'node:url'
test.setTimeout(60_000)
let server: ViteDevServer
let base: string

test.beforeAll(async () => {
  server = await createServer({
    configFile: false,
    cacheDir: fileURLToPath(new URL('../../../node_modules/.vite-home-ribbon', import.meta.url)),
    root: fileURLToPath(new URL('../../..', import.meta.url)),
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

test('Home relative indentation, decimal steps and adjacent directional fills use authoritative batches', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(true)
    await (window as any).toolbar.seedMixedIndent()
    ;(window as any).toolbar.select('A2:B2')
  })
  try {
    const home = page.locator('.workbook-desktop-header')
    const indents = () =>
      page.evaluate(() => {
        const cells = (window as any).toolbar.snapshot().sheets[0].cells
        return [cells.A2.style.indent, cells.B2.style.indent]
      })
    const revision = () => page.evaluate(() => (window as any).toolbar.revision())
    const before = await revision()
    await home.getByRole('button', { name: 'Increase Indent', exact: true }).click()
    await expect.poll(indents).toEqual([2, 4])
    expect(await revision()).toBe(before + 1)
    await home.getByRole('button', { name: 'Decrease Indent', exact: true }).click()
    await expect.poll(indents).toEqual([1, 3])
    await page.evaluate(() => (window as any).toolbar.select('B2:B3'))
    const beforeDecimal = await revision()
    await home.locator('[data-tip="Increase decimal"]').click()
    await expect
      .poll(() =>
        page.evaluate(() => {
          const cells = (window as any).toolbar.snapshot().sheets[0].cells
          return [cells.B2.style.numberFormat, cells.B3.style.numberFormat]
        }),
      )
      .toEqual(['0.0', '0.0'])
    expect(await revision()).toBe(beforeDecimal + 1)
    await home.locator('[data-tip="Decrease decimal"]').click()
    await expect
      .poll(() =>
        page.evaluate(() => {
          const cells = (window as any).toolbar.snapshot().sheets[0].cells
          return [cells.B2.style.numberFormat, cells.B3.style.numberFormat]
        }),
      )
      .toEqual(['0', '0'])
    await page.evaluate(() => (window as any).toolbar.select('A3:B3'))
    await home.getByRole('button', { name: 'Fill', exact: true }).click()
    await page.getByRole('option', { name: 'Fill Down', exact: true }).click()
    await expect
      .poll(() =>
        page.evaluate(() => {
          const cells = (window as any).toolbar.snapshot().sheets[0].cells
          return [cells.A3.rawValue, cells.B3.rawValue]
        }),
      )
      .toEqual(['North', 120])
    await expect
      .poll(() =>
        page.evaluate(() => {
          const cells = (window as any).toolbar.snapshot().sheets[0].cells
          return [cells.A3.style.indent, cells.B3.style.indent, cells.B3.styleReset]
        }),
      )
      .toEqual([1, 3, true])
    await expect
      .poll(() =>
        page.evaluate(() => {
          const styles = (window as any).toolbar.styles('A3:B3')
          return [styles.A3.pd.l, styles.B3.pd.l]
        }),
      )
      .toEqual([8, 24])
    await page.evaluate(() => (window as any).toolbar.select('A1:B1'))
    await home.getByRole('button', { name: 'Fill', exact: true }).click()
    await page.getByRole('option', { name: 'Fill Up', exact: true }).click()
    await expect
      .poll(() =>
        page.evaluate(() => {
          const cells = (window as any).toolbar.snapshot().sheets[0].cells
          return [cells.A1.rawValue, cells.B1.rawValue]
        }),
      )
      .toEqual(['North', 120])
    // The blank B4 donor must clear B3's copied indent and number format.
    await page.evaluate(() => (window as any).toolbar.select('B3'))
    await home.getByRole('button', { name: 'Fill', exact: true }).click()
    await page.getByRole('option', { name: 'Fill Up', exact: true }).click()
    await expect
      .poll(() =>
        page.evaluate(() => {
          const cell = (window as any).toolbar.snapshot().sheets[0].cells.B3
          return [cell.rawValue, cell.styleReset, cell.style]
        }),
      )
      .toEqual([null, true, {}])
    await expect
      .poll(() =>
        page.evaluate(() => {
          const style = (window as any).toolbar.styles('B3').B3
          return Math.round(style.pd.l / 8)
        }),
      )
      .toBe(0)
    await expect(page.getByRole('alert')).toHaveCount(0)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('Home clear choices preserve unrelated cell data and submit Clear All atomically', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(true)
    await (window as any).toolbar.seedClearTargets()
    ;(window as any).toolbar.select('B2')
  })
  const home = page.locator('.workbook-desktop-header')
  const clear = async (name: string) => {
    await home.getByRole('button', { name: 'Clear', exact: true }).click()
    await page.getByRole('option', { name, exact: true }).click()
  }
  const current = () => page.evaluate(() => (window as any).toolbar.snapshot().sheets[0])
  try {
    await clear('Clear Formats')
    await expect.poll(async () => (await current()).cells.B2.styleReset).toBe(true)
    expect((await current()).cells.B2.rawValue).toBe(120)
    expect((await current()).cells.B2.style.bold).toBeUndefined()
    expect((await current()).features.noteEdits.B2).toBe('clear note')
    await clear('Clear Comments and Notes')
    await expect.poll(async () => (await current()).features.noteEdits.B2).toBeNull()
    expect((await current()).features.hyperlinkEdits.B2).toBe('https://example.com/clear')
    await clear('Clear Hyperlinks')
    await expect.poll(async () => (await current()).features.hyperlinkEdits.B2).toBeNull()
    expect((await current()).cells.B2.rawValue).toBe(120)
    await page.evaluate(() => (window as any).toolbar.seedClearTargets())
    const before = await page.evaluate(() => (window as any).toolbar.revision())
    await clear('Clear All')
    await expect.poll(async () => (await current()).cells.B2.rawValue).toBeNull()
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(before + 1)
    const after = await current()
    expect(after.cells.B2.style.bold).toBeUndefined()
    expect(after.features.noteEdits.B2).toBeNull()
    expect(after.features.hyperlinkEdits.B2).toBeNull()
    expect(after.features.noteEdits.A2).toBe('keep note')
    expect(after.features.hyperlinkEdits.A2).toBe('https://example.com/keep')
    expect(after.cells.A2.rawValue).toBe('North')
    await expect(page.getByRole('alert')).toHaveCount(0)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('desktop axis dialogs use points and character units, cancel and reject stale applies', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(true)
    ;(window as any).toolbar.select('B2:C3')
  })
  const home = page.locator('.workbook-desktop-header')
  const open = async (label: string) => {
    await home.getByRole('button', { name: 'Format', exact: true }).click()
    await page.getByRole('option', { name: label + '…', exact: true }).click()
    return page.getByRole('dialog', { name: label, exact: true })
  }
  try {
    let dialog = await open('Row Height')
    await dialog.getByRole('textbox').fill('25,3')
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog).toHaveCount(0)
    await expect
      .poll(() =>
        page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].features.rowHeights),
      )
      .toMatchObject({ '2': 25.5, '3': 25.5 })
    dialog = await open('Column Width')
    await dialog.getByRole('textbox').fill('256')
    await expect(dialog.getByRole('button', { name: 'OK', exact: true })).toBeDisabled()
    await dialog.getByRole('textbox').fill('10')
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(dialog).toHaveCount(0)
    await expect
      .poll(() =>
        page.evaluate(
          () => (window as any).toolbar.snapshot().sheets[0].features.columnWidthsCharacters,
        ),
      )
      .toMatchObject({ B: 10, C: 10 })
    await expect
      .poll(() =>
        page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].features.colWidths),
      )
      .toMatchObject({ B: 80, C: 80 })
    const beforeCancel = await page.evaluate(() => (window as any).toolbar.revision())
    dialog = await open('Row Height')
    await dialog.getByRole('textbox').fill('40')
    await dialog.getByRole('button', { name: 'Cancel', exact: true }).click()
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(beforeCancel)
    dialog = await open('Column Width')
    await dialog.getByRole('textbox').fill('12')
    await page.evaluate(() => (window as any).toolbar.editNote('Concurrent axis update'))
    await dialog.getByRole('button', { name: 'OK', exact: true }).click()
    await expect(page.getByRole('alert')).toContainText('document changed')
    await expect(dialog).toBeVisible()
    await dialog.getByRole('button', { name: 'Cancel', exact: true }).click()
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('desktop AutoFit reads original cells outside the viewport and sizes whole selected axes atomically', async ({
  page,
}) => {
  const errors: string[] = []
  page.on('pageerror', (error) => errors.push(String(error)))
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(true, false, true)
    ;(window as any).toolbar.select('A2')
  })
  const home = page.locator('.workbook-desktop-header')
  const fit = async (name: string) => {
    await home.getByRole('button', { name: 'Format', exact: true }).click()
    await page.getByRole('option', { name, exact: true }).click()
  }
  try {
    const before = await page.evaluate(() => ({
      revision: (window as any).toolbar.revision(),
      snapshot: (window as any).toolbar.snapshot(),
      reads: (window as any).toolbar.reads(),
    }))
    expect(before.reads.some((range: any) => range.startRow <= 79 && range.endRow >= 79)).toBe(
      false,
    )
    await fit('AutoFit Row Height')
    await expect
      .poll(() =>
        page.evaluate(
          () => (window as any).toolbar.snapshot().sheets[0].features.rowHeights?.['2'] ?? 0,
        ),
      )
      .toBeGreaterThan(30)
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(before.revision + 1)
    // D2 is outside the selected column and exists only in the immutable source.
    expect(
      (await page.evaluate(() => (window as any).toolbar.snapshot())).sheets[0].cells.D2,
    ).toBeUndefined()
    await fit('AutoFit Column Width')
    await expect
      .poll(() =>
        page.evaluate(
          () => (window as any).toolbar.snapshot().sheets[0].features.colWidths?.A ?? 0,
        ),
      )
      .toBeGreaterThan(200)
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(before.revision + 2)
    const after = await page.evaluate(() => ({
      snapshot: (window as any).toolbar.snapshot(),
      reads: (window as any).toolbar.reads(),
    }))
    expect(after.reads.some((range: any) => range.startRow <= 79 && range.endRow >= 79)).toBe(true)
    expect(after.snapshot.sheets[0].cells.A80).toBeUndefined()
    expect(after.snapshot.sheets[0].cells.A2.rawValue).toBe(
      before.snapshot.sheets[0].cells.A2.rawValue,
    )
    await expect(page.getByRole('alert')).toHaveCount(0)
    expect(errors).toEqual([])
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('AutoFit rejects measurements from an original-data read overtaken by another edit', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(true, false, true)
    ;(window as any).toolbar.select('A1')
    ;(window as any).toolbar.holdReads()
  })
  try {
    const home = page.locator('.workbook-desktop-header')
    const before = await page.evaluate(() => (window as any).toolbar.revision())
    await home.getByRole('button', { name: 'Format', exact: true }).click()
    await page.getByRole('option', { name: 'AutoFit Column Width', exact: true }).click()
    await expect
      .poll(() =>
        page.evaluate(() =>
          (window as any).toolbar
            .reads()
            .some((range: any) => range.startRow <= 79 && range.endRow >= 79),
        ),
      )
      .toBe(true)
    await page.evaluate(() => {
      ;(window as any).pendingEdit = (window as any).toolbar.editNote('Concurrent AutoFit update')
    })
    await expect
      .poll(() => page.evaluate(() => (window as any).toolbar.revision()))
      .toBe(before + 1)
    await page.evaluate(async () => {
      ;(window as any).toolbar.releaseReads()
      await (window as any).pendingEdit
    })
    await expect(page.getByRole('alert')).toContainText('document changed')
    const state = await page.evaluate(() => (window as any).toolbar.snapshot())
    expect(state.sheets[0].features.colWidths?.A).toBeUndefined()
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(before + 1)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('desktop Go To navigates references and names across sheets without authored edits', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(true, false, false, true)
  })
  const home = page.locator('.workbook-desktop-header')
  const open = async () => {
    await home.getByRole('button', { name: /Go To/ }).click()
    return page.getByRole('dialog', { name: 'Go To', exact: true })
  }
  const selection = () => page.evaluate(() => (window as any).toolbar.selection().selection)
  const before = await page.evaluate(() => (window as any).toolbar.snapshot())
  const revision = await page.evaluate(() => (window as any).toolbar.revision())
  try {
    let dialog = await open()
    await dialog.getByRole('textbox', { name: 'Reference', exact: true }).fill('A0')
    await dialog.getByRole('button', { name: 'Go', exact: true }).click()
    await expect(dialog.getByRole('alert')).toContainText('A0')
    await dialog.getByRole('textbox', { name: 'Reference', exact: true }).fill('Missing!A1')
    await dialog.getByRole('button', { name: 'Go', exact: true }).click()
    await expect(dialog.getByRole('alert')).toContainText('sheet was not found')
    await dialog.getByRole('button', { name: /GoTarget/ }).dblclick()
    await expect(dialog).toHaveCount(0)
    await expect.poll(selection).toMatchObject({ sheetId: 'toolbar-other', address: 'A1:A1' })
    dialog = await open()
    await dialog.getByRole('textbox', { name: 'Reference', exact: true }).fill('Data!Q2000')
    await dialog.getByRole('button', { name: 'Go', exact: true }).click()
    await expect
      .poll(async () => ({
        open: await dialog.count(),
        errors: await dialog.getByRole('alert').allTextContents(),
      }))
      .toEqual({ open: 0, errors: [] })
    await expect.poll(selection).toMatchObject({ sheetId: 'toolbar-sheet', address: 'Q2000:Q2000' })
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(revision)
    expect(await page.evaluate(() => (window as any).toolbar.snapshot())).toEqual(before)
    dialog = await open()
    await dialog.getByRole('textbox', { name: 'Reference', exact: true }).fill('MovingTarget')
    await dialog.getByRole('button', { name: 'Go', exact: true }).click()
    await expect(dialog.getByRole('alert')).toContainText('navigable defined name')
    await dialog.getByRole('button', { name: 'Cancel', exact: true }).click()
    await expect.poll(selection).toMatchObject({ address: 'Q2000:Q2000' })
    dialog = await open()
    await dialog.getByRole('textbox', { name: 'Reference', exact: true }).fill('Data!Q2000')
    await dialog.getByRole('button', { name: 'Go', exact: true }).click()
    await expect(dialog).toHaveCount(0)
    await page.evaluate(
      () => new Promise<void>((resolve) => requestAnimationFrame(() => resolve())),
    )
    await page.keyboard.type('42')
    await page.keyboard.press('Enter')
    await expect
      .poll(() =>
        page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells.Q2000?.rawValue),
      )
      .toBe(42)
    expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(revision + 1)
  } finally {
    await page.evaluate(() => (window as any).toolbar.dispose())
  }
})

test('shared Insert Symbol waits for one host commit, keeps the dialog for repeated inserts and rejects an overtaken target', async ({
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
    ;(window as any).toolbar.select('B2:C3')
  })
  await page.getByRole('tab', { name: 'Insert', exact: true }).click()
  await page.locator('.workbook-desktop-header [data-tip="Symbol"]').click()
  const dialog = page.getByRole('dialog', { name: 'Symbol', exact: true })
  await dialog.getByRole('button', { name: 'Greek', exact: true }).click()
  await page.evaluate(() => (window as any).toolbar.holdReads())
  const symbol = dialog.getByRole('button', { name: 'Ω', exact: true })
  await symbol.click()
  await expect(symbol).toBeDisabled()
  await symbol.dispatchEvent('click')
  await expect(dialog.getByRole('button', { name: 'Close', exact: true })).toBeDisabled()
  await page.keyboard.press('Escape')
  await expect(dialog).toBeVisible()
  await page.evaluate(() => {
    ;(window as any).symbolFlushResult = null
    ;(window as any).symbolFlushPromise = (window as any).toolbar.flush().then((ready: boolean) => {
      ;(window as any).symbolFlushResult = ready
    })
  })
  expect(await page.evaluate(() => (window as any).symbolFlushResult)).toBeNull()
  expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(0)
  expect(
    await page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells.B2.rawValue),
  ).toBe(120)
  await page.evaluate(() => (window as any).toolbar.releaseReads())
  await page.evaluate(() => (window as any).symbolFlushPromise)
  expect(await page.evaluate(() => (window as any).symbolFlushResult)).toBe(true)
  await expect(symbol).toBeEnabled()
  await expect
    .poll(() => page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells.B2.rawValue))
    .toBe('120Ω')
  expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(1)
  expect(
    await page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells.C3),
  ).toBeUndefined()
  await dialog.getByRole('button', { name: 'α', exact: true }).click()
  await expect
    .poll(() => page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells.B2.rawValue))
    .toBe('120Ωα')
  expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(2)
  await page.evaluate(() => (window as any).toolbar.editNote('Concurrent note'))
  const revision = await page.evaluate(() => (window as any).toolbar.revision())
  await symbol.click()
  await expect(page.getByRole('alert')).toContainText('document changed')
  expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(revision)
  expect(
    await page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells.B2.rawValue),
  ).toBe('120Ωα')
  await expect(dialog).toBeVisible()
  await dialog.getByRole('button', { name: 'Close', exact: true }).click()
  await expect(dialog).toHaveCount(0)
  expect(errors).toEqual([])
})

test('shared Header & Footer preloads source sections, inserts fields and validates atomic frozen settings', async ({
  page,
}) => {
  const errors: string[] = []
  page.on('pageerror', (error) => errors.push(String(error)))
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(true, false, false, false, true)
    ;(window as any).toolbar.select('B2')
  })
  await page.getByRole('tab', { name: 'Insert', exact: true }).click()
  const open = () => page.locator('.workbook-desktop-header [data-tip="Header & Footer"]').click()
  const dialog = page.getByRole('dialog', { name: 'Header & Footer', exact: true })
  await open()
  await expect(dialog.getByRole('textbox', { name: 'Header · Left', exact: true })).toHaveValue(
    'Original left',
  )
  await expect(dialog.getByRole('textbox', { name: 'Header · Center', exact: true })).toHaveValue(
    'Page &P of &N',
  )
  await dialog.getByRole('button', { name: 'Cancel', exact: true }).click()
  expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(0)
  await open()
  const right = dialog.getByRole('textbox', { name: 'Header · Right', exact: true })
  await right.fill('AB')
  await right.evaluate((element: HTMLInputElement) => element.setSelectionRange(1, 1))
  await dialog.getByRole('button', { name: 'Page Number &P', exact: true }).click()
  await expect(right).toHaveValue('A&PB')
  await dialog.getByRole('button', { name: 'OK', exact: true }).click()
  await expect(dialog).toHaveCount(0)
  expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(1)
  expect(
    await page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].features.pageSetupEdits),
  ).toMatchObject({
    header: { left: 'Original left', center: 'Page &P of &N', right: 'A&PB' },
    footer: { right: 'Confidential' },
  })
  await open()
  await expect(right).toHaveValue('A&PB')
  await right.fill('x'.repeat(256))
  await dialog.getByRole('button', { name: 'OK', exact: true }).click()
  await expect(dialog.getByRole('alert')).toBeVisible()
  expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(1)
  for (const input of await dialog.getByRole('textbox').all()) await input.fill('')
  await dialog.getByRole('button', { name: 'OK', exact: true }).click()
  await expect(dialog).toHaveCount(0)
  expect(
    await page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].features.pageSetupEdits),
  ).toMatchObject({ header: null, footer: null })
  await open()
  for (const input of await dialog.getByRole('textbox').all()) await expect(input).toHaveValue('')
  await page.evaluate(() => (window as any).toolbar.editNote('Concurrent note'))
  await right.fill('Stale settings')
  await dialog.getByRole('button', { name: 'OK', exact: true }).click()
  await expect(dialog.getByRole('alert')).toContainText('document changed')
  expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(3)
  await dialog.getByRole('button', { name: 'Cancel', exact: true }).click()
  expect(
    await page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells.B2.rawValue),
  ).toBe(120)
  expect(errors).toEqual([])
})

test('shared Insert chart gallery reads source data beyond the viewport and renders all desktop plot families', async ({
  page,
}) => {
  const errors: string[] = []
  page.on('pageerror', (error) => errors.push(String(error)))
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(true, false, false, false, false, true)
    await (window as any).toolbar.reveal('A78:C81')
  })
  await page.getByRole('tab', { name: 'Insert', exact: true }).click()
  const labels = [
    'Column',
    'Bar',
    'Line',
    'Area',
    'Pie',
    'Scatter',
    'Radar',
    'Doughnut',
    'Combo (column + line)',
  ]
  for (const [index, label] of labels.entries()) {
    await page
      .locator('.workbook-desktop-header')
      .getByRole('button', { name: `${label} chart — from selection`, exact: true })
      .click()
    await expect.poll(() => page.evaluate(() => (window as any).toolbar.revision())).toBe(index + 1)
    await expect(page.locator('.xlsx-print-visual .xlsx-chart')).toHaveCount(index + 1)
  }
  const charts = await page.evaluate(
    () => (window as any).toolbar.snapshot().sheets[0].features.charts,
  )
  expect(charts.map((visual: any) => visual.chart.chartTypes)).toEqual([
    ['barChart'],
    ['barChart'],
    ['lineChart'],
    ['areaChart'],
    ['pieChart'],
    ['scatterChart'],
    ['radarChart'],
    ['doughnutChart'],
    ['barChart', 'lineChart'],
  ])
  for (const chart of charts) {
    expect(chart.chart.series.map((series: any) => series.values)).toEqual([
      [12, 8, 18],
      [6, 10, 14],
    ])
    expect(chart.chart.series[0].categories).toEqual(['2020', '2021', '2022'])
    expect(chart.chart.series[0].valuesRef).toBe("'Data'!$B$79:$B$81")
    expect(chart.chart.series[0].categoriesRef).toBe("'Data'!$A$79:$A$81")
    expect(chart.chart.dataLabels).toBeUndefined()
  }
  expect(charts[0].anchor).toMatchObject({ fromRow: 77, fromColumn: 4, toRow: 92, toColumn: 11 })
  expect(charts[1].chart.barDirection).toBe('bar')
  expect(await page.locator('.xlsx-print-visual svg').count()).toBe(9)
  expect(errors).toEqual([])
})

test('shared Recommended Charts keeps frozen data and awaits accepted insertion before closing or flushing', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(true, false, false, false, false, true)
    await (window as any).toolbar.reveal('A78:C81')
  })
  await page.getByRole('tab', { name: 'Insert', exact: true }).click()
  const open = () =>
    page.locator('.workbook-desktop-header [data-tip="Recommended Charts"]').click()
  const dialog = page.getByRole('dialog', { name: 'Recommended Charts', exact: true })
  await open()
  await expect(dialog).toBeVisible()
  await expect(dialog.locator('.recommended-charts-grid button').first()).toContainText('Line')
  await dialog.getByRole('button', { name: 'Close', exact: true }).click()
  expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(0)
  await open()
  await page.evaluate(() => (window as any).toolbar.holdReads())
  const choice = dialog.locator('.recommended-charts-grid button').first()
  await choice.click()
  await expect(choice).toBeDisabled()
  await choice.dispatchEvent('click')
  await page.keyboard.press('Escape')
  await expect(dialog).toBeVisible()
  await page.evaluate(() => {
    ;(window as any).chartFlushResult = null
    ;(window as any).chartFlush = (window as any).toolbar.flush().then((ready: boolean) => {
      ;(window as any).chartFlushResult = ready
    })
  })
  expect(await page.evaluate(() => (window as any).chartFlushResult)).toBeNull()
  expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(0)
  await page.evaluate(() => (window as any).toolbar.releaseReads())
  await page.evaluate(() => (window as any).chartFlush)
  await expect(dialog).toHaveCount(0)
  expect(await page.evaluate(() => (window as any).chartFlushResult)).toBe(true)
  expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(1)
  await open()
  await page.evaluate(() => (window as any).toolbar.editNote('Concurrent note'))
  await choice.click()
  await expect(dialog.getByRole('alert')).toContainText('document changed')
  expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(2)
  await dialog.getByRole('button', { name: 'Close', exact: true }).click()
})

test('Insert charts reject nonnumeric, incomplete-formula and oversized data without poisoning later valid insertion', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(true, false, false, false, false, true)
    ;(window as any).toolbar.select('A1')
  })
  await page.getByRole('tab', { name: 'Insert', exact: true }).click()
  const insert = () =>
    page
      .locator('.workbook-desktop-header')
      .getByRole('button', { name: 'Column chart — from selection', exact: true })
      .click()
  await insert()
  await expect(page.getByRole('alert')).toContainText('numeric')
  expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(0)
  await page.evaluate(() => (window as any).toolbar.select('D78'))
  await insert()
  await expect(page.getByRole('alert')).toContainText('formulas are still pending')
  expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(0)
  await page.evaluate(() => (window as any).toolbar.select('A1:T101'))
  const reads = await page.evaluate(() => (window as any).toolbar.reads().length)
  await insert()
  await expect(page.getByRole('alert')).toContainText('at most 2000')
  expect(
    await page.evaluate(
      (start) =>
        (window as any).toolbar
          .reads()
          .slice(start)
          .filter(
            (range: any) =>
              range.startRow === 0 &&
              range.endRow === 100 &&
              range.startColumn === 0 &&
              range.endColumn === 19,
          ),
      reads,
    ),
  ).toEqual([])
  expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(0)
  await page.evaluate(() => (window as any).toolbar.reveal('A78:C81'))
  await insert()
  await expect.poll(() => page.evaluate(() => (window as any).toolbar.revision())).toBe(1)
  await expect(page.getByRole('alert')).toHaveCount(0)
})

test('desktop floating source chart moves, resizes and deletes through authoritative drawing edits', async ({
  page,
}) => {
  await page.goto(base)
  await page.addStyleTag({ content: 'body { padding: 0 20px; }' })
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(
      true,
      false,
      false,
      false,
      false,
      false,
      true,
    )
  })
  expect(await page.evaluate(() => getComputedStyle(document.body).paddingLeft)).toBe('0px')
  const frame = page.locator('[data-print-visual="original-chart"] .shape-editable')
  await expect(frame).toBeVisible()
  const cells = await page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells)
  const before = await frame.boundingBox()
  expect(before).not.toBeNull()
  const position = await frame.locator('figcaption').boundingBox()
  await page.mouse.move(position!.x + 50, position!.y + 8)
  await page.mouse.down()
  await page.mouse.move(position!.x + 90, position!.y + 38, { steps: 8 })
  await page.evaluate(() => {
    ;(window as any).dragFlushResult = null
    ;(window as any).dragFlush = (window as any).toolbar.flush().then((ready: boolean) => {
      ;(window as any).dragFlushResult = ready
    })
  })
  expect(await page.evaluate(() => (window as any).dragFlushResult)).toBeNull()
  await page.mouse.up()
  await page.evaluate(() => (window as any).dragFlush)
  expect(await page.evaluate(() => (window as any).dragFlushResult)).toBe(true)
  await expect.poll(() => page.evaluate(() => (window as any).toolbar.revision())).toBe(1)
  await expect.poll(async () => Math.round((await frame.boundingBox())!.x - before!.x)).toBe(40)
  await expect.poll(async () => Math.round((await frame.boundingBox())!.y - before!.y)).toBe(30)
  const moved = await frame.boundingBox()
  const corner = await frame.locator('[data-corner="se"]').boundingBox()
  await page.mouse.move(corner!.x + corner!.width / 2, corner!.y + corner!.height / 2)
  await page.mouse.down()
  await page.mouse.move(corner!.x + corner!.width / 2 + 45, corner!.y + corner!.height / 2 + 35, {
    steps: 8,
  })
  await page.mouse.up()
  await expect.poll(() => page.evaluate(() => (window as any).toolbar.revision())).toBe(2)
  await expect
    .poll(async () => Math.round((await frame.boundingBox())!.width - moved!.width))
    .toBe(45)
  await expect
    .poll(async () => Math.round((await frame.boundingBox())!.height - moved!.height))
    .toBe(35)
  expect(await page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].cells)).toEqual(
    cells,
  )
  const source = await page.evaluate(
    () => (window as any).toolbar.snapshot().features.originalMetadata.visuals[0].anchor,
  )
  const patch = await page.evaluate(
    () =>
      (window as any).toolbar.snapshot().sheets[0].features.visualEdits['original-chart'].anchor,
  )
  expect(patch).not.toEqual(source)
  await frame.press('Delete')
  await expect(frame).toHaveCount(0)
  expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(3)
  expect(
    await page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].features.visualEdits),
  ).toEqual({ 'original-chart': { remove: true } })
})

test('a source chart drag overtaken by a concurrent transaction rejects the old anchor', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(
      true,
      false,
      false,
      false,
      false,
      false,
      true,
    )
  })
  const frame = page.locator('[data-print-visual="original-chart"] .shape-editable')
  await expect(frame).toBeVisible()
  const heading = await frame.locator('figcaption').boundingBox()
  await page.mouse.move(heading!.x + 50, heading!.y + 8)
  await page.mouse.down()
  await page.mouse.move(heading!.x + 90, heading!.y + 38, { steps: 4 })
  await page.evaluate(() => {
    ;(window as any).concurrentVisual = (window as any).toolbar.editNote(
      'Concurrent while dragging',
    )
  })
  await page.mouse.up()
  await page.evaluate(() => (window as any).concurrentVisual)
  await expect(page.getByRole('alert')).toContainText('document changed during the visual edit')
  expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(1)
  expect(
    await page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].features.visualEdits),
  ).toBeUndefined()
})

test('Escape cancels a floating chart drag and releases editor flushing without a write', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(
      true,
      false,
      false,
      false,
      false,
      false,
      true,
    )
  })
  const frame = page.locator('[data-print-visual="original-chart"] .shape-editable')
  const heading = (await frame.locator('figcaption').boundingBox())!
  await page.mouse.move(heading.x + 50, heading.y + 8)
  await page.mouse.down()
  await page.mouse.move(heading.x + 90, heading.y + 38, { steps: 4 })
  await expect(page.locator('.shape-drag-ghost')).toHaveCount(1)
  await page.keyboard.press('Escape')
  expect(await page.evaluate(() => (window as any).toolbar.flush())).toBe(true)
  await page.mouse.up()
  await expect(page.locator('.shape-drag-ghost')).toHaveCount(0)
  expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(0)
})

test('Shift-selecting original and session charts deletes both with one authoritative transaction', async ({
  page,
}) => {
  await page.goto(base)
  await page.evaluate(async () => {
    const { mountManualToolbar } = await import(
      /* @vite-ignore */ String('/packages/sheets-editor/tests/fixtures/manual-toolbar.tsx')
    )
    ;(window as any).toolbar = await mountManualToolbar(
      true,
      false,
      false,
      false,
      false,
      false,
      true,
    )
    ;(window as any).toolbar.select('A1:B3')
  })
  const original = page.locator('[data-print-visual="original-chart"] .shape-editable')
  await original.focus()
  await page.getByRole('tab', { name: 'Insert', exact: true }).click()
  await page
    .locator('.workbook-desktop-header')
    .getByRole('button', { name: 'Column chart — from selection', exact: true })
    .click()
  await expect.poll(() => page.evaluate(() => (window as any).toolbar.revision())).toBe(1)
  const added = page.locator(
    '.xlsx-print-visual:not([data-print-visual="original-chart"]) .shape-editable',
  )
  await added.locator('figcaption').click({ modifiers: ['Shift'] })
  await expect(page.locator('.shape-editable.selected')).toHaveCount(2)
  await added.press('Delete')
  await expect(page.locator('.xlsx-print-visual')).toHaveCount(0)
  expect(await page.evaluate(() => (window as any).toolbar.revision())).toBe(2)
  expect(
    await page.evaluate(() => (window as any).toolbar.snapshot().sheets[0].features.visualEdits),
  ).toEqual({ 'original-chart': { remove: true } })
})

test('desktop window shortcuts and Format Painter route through the shared command table', async ({
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
  const snapshot = () => page.evaluate(() => (window as any).toolbar.snapshot())
  const revision = () => page.evaluate(() => (window as any).toolbar.revision())
  await page.getByRole('tab', { name: 'Home', exact: true }).click()
  // Ctrl+Shift+5: percent format, the desktop's number-format chord.
  await page.keyboard.press('Control+Shift+Digit5')
  await expect.poll(async () => (await snapshot()).sheets[0].cells.B2.numberFormat).toBe('0%')
  // Ctrl+1 opens Format Cells; Ctrl+T opens Create Table.
  await page.keyboard.press('Control+Digit1')
  const format = page.getByRole('dialog', { name: 'Format Cells', exact: true })
  await expect(format).toBeVisible()
  await format.getByRole('button', { name: 'Cancel', exact: true }).click()
  await page.getByRole('tab', { name: 'Home', exact: true }).click()
  await page.keyboard.press('Control+KeyT')
  const table = page.getByRole('dialog', { name: 'Create Table', exact: true })
  await expect(table).toBeVisible()
  await table.getByRole('button', { name: 'Cancel', exact: true }).click()
  // Shift+F11 appends the next free SheetN.
  await page.getByRole('tab', { name: 'Home', exact: true }).click()
  const before = await revision()
  await page.keyboard.press('Shift+F11')
  await expect.poll(revision).toBe(before + 1)
  expect((await snapshot()).sheets.map((sheet: { name: string }) => sheet.name)).toContain('Sheet1')
  // Format Painter: B2's formats paint onto B3 when the target selection is released.
  await page.evaluate(() => {
    ;(window as any).toolbar.selectSheet('Data')
    ;(window as any).toolbar.select('B2')
  })
  const painter = page
    .locator('.workbook-desktop-header')
    .getByRole('button', { name: 'Format Painter', exact: true })
  await painter.click()
  await expect(painter).toHaveAttribute('aria-pressed', 'true')
  const painted = await revision()
  await page.evaluate(() => {
    ;(window as any).toolbar.select('B3')
    document
      .querySelector('canvas[id^="univer-sheet-main-canvas"]')!
      .dispatchEvent(new PointerEvent('pointerup', { bubbles: true, button: 0 }))
  })
  await expect.poll(revision).toBe(painted + 1)
  expect((await snapshot()).sheets[0].cells.B3.numberFormat).toBe('0%')
  await expect(painter).toHaveAttribute('aria-pressed', 'false')
  expect(errors).toEqual([])
})

test('Home Insert Cells and Fill Series write canonical edits through the desktop dialogs', async ({
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
  const cells = async () =>
    (await page.evaluate(() => (window as any).toolbar.snapshot())).sheets[0].cells
  const home = page.locator('.workbook-desktop-header')
  await home.getByRole('button', { name: 'Insert Cells…', exact: true }).click()
  const insert = page.getByRole('dialog', { name: 'Insert', exact: true })
  await insert.getByRole('button', { name: 'OK', exact: true }).click()
  await expect.poll(async () => (await cells()).B4?.rawValue).toBe(-30)
  expect((await cells()).B3?.rawValue).toBe(120)
  expect((await cells()).B2?.rawValue ?? null).toBe(null)
  expect((await cells()).A2?.rawValue).toBe('North')
  await page.evaluate(() => (window as any).toolbar.select('B3'))
  await home.getByRole('button', { name: 'Fill', exact: true }).click()
  await page.getByRole('listbox').getByRole('option', { name: 'Series…', exact: true }).click()
  const series = page.getByRole('dialog', { name: 'Series', exact: true })
  await series.getByLabel('Columns', { exact: true }).check()
  await series.getByLabel('Step value', { exact: true }).fill('5')
  await series.getByLabel('Stop value', { exact: true }).fill('130')
  await series.getByRole('button', { name: 'OK', exact: true }).click()
  await expect.poll(async () => (await cells()).B5?.rawValue).toBe(130)
  expect((await cells()).B4?.rawValue).toBe(125)
})

test('Conditional Formatting opens the shared rule manager and saves panel rules as one session edit', async ({
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
    ;(window as any).toolbar.select('B2:B3')
  })
  const revision = () => page.evaluate(() => (window as any).toolbar.revision())
  await page
    .locator('.workbook-desktop-header')
    .getByRole('button', { name: 'Conditional Formatting', exact: true })
    .click()
  await expect(page.locator('span', { hasText: /^Manage\s*$/ }).first()).toBeVisible()
  const before = await revision()
  // A rule added through Univer's own command path, as the panel's Apply does.
  await page.evaluate(() => {
    const sheet = (window as any).toolbar.facade().getActiveWorkbook().getActiveSheet()
    const rule = sheet
      .newConditionalFormattingRule()
      .whenNumberGreaterThan(100)
      .setBackground('#FFC7CE')
      .setRanges([sheet.getRange('B2:B3').getRange()])
      .build()
    sheet.addConditionalFormattingRule(rule)
  })
  await expect.poll(revision).toBe(before + 1)
  const edits = await page.evaluate(
    () => (window as any).toolbar.snapshot().sheets[0].features.cfEdits,
  )
  expect(edits.rules).toHaveLength(1)
  expect(edits.rules[0].ranges).toEqual([{ startRow: 1, endRow: 2, startColumn: 1, endColumn: 1 }])
  expect(edits.rules[0].rule.type).toBe('highlightCell')
  // The projected rule survives the reinstall that follows the commit.
  await expect
    .poll(() =>
      page.evaluate(
        () =>
          (window as any).toolbar
            .facade()
            .getActiveWorkbook()
            .getActiveSheet()
            .getConditionalFormattingRules().length,
      ),
    )
    .toBe(1)
  expect(errors).toEqual([])
})

test('Excel keyboard layer hides rows and opens Insert Cells through session edits', async ({
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
  })
  const canvas = page.locator('canvas[id^="univer-sheet-main-canvas"]')
  await canvas.click({ position: { x: 200, y: 200 } })
  await page.evaluate(() => (window as any).toolbar.select('A2'))
  const revision = () => page.evaluate(() => (window as any).toolbar.revision())
  const before = await revision()
  await page.keyboard.press('Control+Digit9')
  await expect.poll(revision).toBe(before + 1)
  const edits = await page.evaluate(
    () => (window as any).toolbar.snapshot().sheets[0].features.visibilityEdits,
  )
  expect(edits).toEqual([{ axis: 'rows', start: 1, count: 1, hidden: true }])
  await page.evaluate(() => (window as any).toolbar.select('B3'))
  await page.keyboard.press('Control+Shift+Equal')
  await expect(page.getByRole('dialog', { name: 'Insert', exact: true })).toBeVisible()
  expect(errors).toEqual([])
})

test('the desktop cell context menu routes built-in and shared entries to session edits and dialogs', async ({
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
  })
  const canvas = page.locator('canvas[id^="univer-sheet-main-canvas"]')
  await canvas.click({ position: { x: 120, y: 50 }, button: 'right' })
  await page.locator('section.univer-popup button', { hasText: 'Format Cells…' }).click()
  const format = page.getByRole('dialog', { name: 'Format Cells', exact: true })
  await expect(format).toBeVisible()
  await format.getByRole('button', { name: 'Cancel', exact: true }).click()
  const revision = () => page.evaluate(() => (window as any).toolbar.revision())
  const snapshot = () => page.evaluate(() => (window as any).toolbar.snapshot())
  const run = (id: string, params?: unknown) =>
    page.evaluate(([id, params]) => (window as any).toolbar.facade().executeCommand(id, params), [
      id,
      params,
    ] as const)
  await page.evaluate(() => (window as any).toolbar.select('A2'))
  let before = await revision()
  expect(await run('sheet.command.insert-multi-rows-above', { value: 2 })).toBe(true)
  await expect.poll(revision).toBe(before + 1)
  expect((await snapshot()).sheets[0].cells.A4?.rawValue).toBe('North')
  await page.evaluate(() => (window as any).toolbar.select('A4:B4'))
  before = await revision()
  expect(await run('sheet.command.clear-selection-content')).toBe(true)
  await expect.poll(revision).toBe(before + 1)
  expect((await snapshot()).sheets[0].cells.A4?.rawValue ?? null).toBe(null)
  before = await revision()
  expect(await run('sheet.command.hide-row-confirm')).toBe(true)
  await expect.poll(revision).toBe(before + 1)
  expect(errors).toEqual([])
})
