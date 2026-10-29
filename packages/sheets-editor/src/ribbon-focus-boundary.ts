import { DOCS_NORMAL_EDITOR_UNIT_ID_KEY } from '@univerjs/core'
import { DocSelectionRenderService, type IRenderManagerService } from '@univerjs/preset-sheets-core'

/** Engine synchronization must preserve focused ribbon tabs, dialogs and pane controls. */
export function installRibbonFocusBoundary(renders: IRenderManagerService, grid: HTMLElement) {
  let tab: HTMLElement | null = null
  const root = grid.closest('.workbook-editor')
  const owner = grid.ownerDocument
  const focusIn = (event: FocusEvent) => {
    const target = event.target
    tab =
      root &&
      target instanceof HTMLElement &&
      target.closest('.workbook-editor') === root &&
      ((target.getAttribute('role') === 'tab' && target.closest('.workbook-ribbon-tabs')) ||
        target.closest('.chart-format-pane, .shape-text-editor'))
        ? target
        : null
  }
  // A real grid click releases tab ownership before the engine's pointer handlers.
  const pointerDown = (event: PointerEvent) => {
    if (!(event.target instanceof Node) || !tab?.contains(event.target)) tab = null
  }
  owner.addEventListener('focusin', focusIn, true)
  owner.addEventListener('pointerdown', pointerDown, true)
  const restorers = new Map<string, () => void>()
  const attach = (render: ReturnType<IRenderManagerService['getRenderById']>) => {
    if (!render || render.unitId !== DOCS_NORMAL_EDITOR_UNIT_ID_KEY || restorers.has(render.unitId))
      return
    const selection = render.with(DocSelectionRenderService)
    const original = selection.focus
    const guarded = () => {
      if (tab?.isConnected && owner.activeElement === tab) return
      // Dialogs and side panes own focus while open; engine sync must not move
      // typing into the hidden cell editor (Go To, Fields pane, chart pane, …).
      const active = owner.activeElement
      if (
        active instanceof HTMLElement &&
        !grid.contains(active) &&
        active.closest('[role="dialog"], [role="complementary"], .chart-format-pane')
      )
        return
      original.call(selection)
    }
    selection.focus = guarded
    restorers.set(render.unitId, () => {
      if (selection.focus === guarded) selection.focus = original
    })
  }
  const created = renders.created$.subscribe(attach)
  const disposed = renders.disposed$.subscribe((id) => {
    restorers.get(id)?.()
    restorers.delete(id)
  })
  attach(renders.getRenderById(DOCS_NORMAL_EDITOR_UNIT_ID_KEY))
  return {
    release() {
      tab = null
    },
    dispose() {
      owner.removeEventListener('focusin', focusIn, true)
      owner.removeEventListener('pointerdown', pointerDown, true)
      created.unsubscribe()
      disposed.unsubscribe()
      for (const restore of restorers.values()) restore()
      restorers.clear()
    },
  }
}
