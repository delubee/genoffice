// @vitest-environment jsdom
import { act, createElement, type ComponentProps } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { LocaleScope } from '../src/renderer/i18n/locale'
import { SheetsRibbon, ribbonTabs } from '../src/renderer/SheetsRibbon'

Object.assign(globalThis, { IS_REACT_ACT_ENVIRONMENT: true })

// No preload globals: the embedded host supplies chat and commands through props.
const commands = vi.fn()
const aiRun = vi.fn()
const props: ComponentProps<typeof SheetsRibbon> = {
  activeTab: 'Home',
  selectionFormat: null,
  formatPainterActive: false,
  sheetHasContent: true,
  sheetProtected: false,
  workbookProtected: false,
  formulaBarVisible: true,
  crossHighlightVisible: false,
  pageLayout: { showGridlines: true, showHeadings: true, rightToLeft: false },
  selectedChart: null,
  selectedShape: null,
  selectedTable: null,
  outlineSummary: { below: true, right: true },
  onCommand: commands,
  onListNames: () => ['Revenue'],
  calcManual: false,
  onAiRun: aiRun,
  aiOpen: false,
  onAiToggle: vi.fn(),
  onRefreshPivot: () => null,
  onIsSelectionInPivot: () => false,
}
let root: Root | undefined
async function render(patch: Partial<typeof props> = {}, lang: 'en' | 'zh' = 'en') {
  if (!root) {
    document.body.innerHTML = '<main id="ribbon-test"></main>'
    root = createRoot(document.getElementById('ribbon-test')!)
  }
  await act(async () => {
    root!.render(
      createElement(LocaleScope, {
        lang,
        children: createElement(SheetsRibbon, { ...props, ...patch }),
      }),
    )
  })
}
async function click(selector: string) {
  const button = document.querySelector<HTMLElement>(selector)
  expect(button, selector).not.toBeNull()
  await act(async () => button!.click())
}
afterEach(async () => {
  if (root) await act(async () => root!.unmount())
  root = undefined
  document.body.innerHTML = ''
  vi.clearAllMocks()
})

describe('desktop ribbon embedded without Electron', () => {
  it('renders all standard and contextual tabs without a desktop preload', async () => {
    for (const activeTab of [
      ...ribbonTabs,
      'Chart Design',
      'Table Design',
      'Shape Format',
      'Picture Format',
    ] as const) {
      await render({ activeTab })
      expect(document.querySelector('.ribbon'), activeTab).not.toBeNull()
      expect(document.querySelectorAll('button').length, activeTab).toBeGreaterThan(0)
    }
  })

  it('passes formatting and menu parameters to the injected command host', async () => {
    await render()
    await click('[data-tip="Bold"]')
    expect(commands).toHaveBeenLastCalledWith('bold')
    await click('[aria-label="Number format"]')
    const percent = [...document.querySelectorAll<HTMLElement>('[role="option"]')].find(
      (option) => option.textContent === 'Percentage',
    )
    expect(percent).toBeDefined()
    await act(async () => percent!.click())
    expect(commands).toHaveBeenLastCalledWith('format:0.00%')
    expect(document.querySelector('[role="listbox"]')).toBeNull()
  })

  it('echoes selection and view state instead of treating tools as stateless buttons', async () => {
    await render({ activeTab: 'View' })
    expect(document.querySelector('[data-tip="Gridlines"] .check-box')?.textContent).toBe('✓')
    await click('[data-tip="Gridlines"]')
    expect(commands).toHaveBeenLastCalledWith('toggle-gridlines')
    await render({ activeTab: 'View', pageLayout: { ...props.pageLayout, showGridlines: false } })
    expect(document.querySelector('[data-tip="Gridlines"] .check-box')?.textContent).toBe('')
  })

  it('uses the host controlled locale and closes popovers on Escape', async () => {
    await render({}, 'zh')
    expect(document.querySelector('[data-tip="加粗"]')).not.toBeNull()
    await render({}, 'en')
    expect(document.querySelector('[data-tip="Bold"]')).not.toBeNull()
    await click('[aria-label="Number format"]')
    expect(document.querySelector('[role="listbox"]')).not.toBeNull()
    await act(async () => {
      window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }))
    })
    expect(document.querySelector('[role="listbox"]')).toBeNull()
    expect(document.documentElement.classList.contains('genoffice-popover-open')).toBe(false)
  })
})
