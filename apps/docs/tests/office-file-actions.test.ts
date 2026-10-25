import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { expect, it, vi } from 'vitest'
import { OfficeFileActions, OFFICE_ACTION_LABELS } from '@genoffice/ui/chrome'

it('routes each file/chat menu action, disables concurrent clicks and reports failures', async () => {
  const node = document.createElement('div')
  document.body.append(node)
  const root = createRoot(node)
  let release!: () => void
  const open = vi.fn(
    () =>
      new Promise<void>((resolve) => {
        release = resolve
      }),
  )
  const save = vi.fn(async () => {
    throw new Error('disk changed')
  })
  const copy = vi.fn(async () => {})
  const chat = vi.fn(async () => {})
  const error = vi.fn()
  try {
    await act(async () =>
      root.render(
        createElement(OfficeFileActions, {
          labels: OFFICE_ACTION_LABELS.en,
          onOpen: open,
          onSave: save,
          onSaveAs: copy,
          onChat: chat,
          onError: error,
        }),
      ),
    )
    const buttons = [...node.querySelectorAll('button')]
    expect(buttons).toHaveLength(5)
    await act(async () => buttons[0]!.click())
    expect(buttons.every((button) => button.disabled)).toBe(true)
    await act(async () => buttons[1]!.click())
    expect(save).not.toHaveBeenCalled()
    await act(async () => release())
    await act(async () => buttons[1]!.click())
    expect(error).toHaveBeenCalledWith(expect.objectContaining({ message: 'disk changed' }))
    expect(buttons.every((button) => !button.disabled)).toBe(true)
    await act(async () => buttons[2]!.click())
    await act(async () => buttons[3]!.click())
    await act(async () => buttons[4]!.click())
    expect(copy).toHaveBeenCalledOnce()
    expect(chat.mock.calls).toEqual([['document'], ['selection']])
  } finally {
    await act(async () => root.unmount())
    node.remove()
  }
})
