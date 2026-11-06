import { expect, it, vi } from 'vitest'
import {
  officeActionHandler,
  type OfficeActionPorts,
  type OfficeChatBinding,
} from '../src/extension/office-action-handler'

function fixture() {
  const events: string[] = []
  const binding: OfficeChatBinding = {
    documentUri: 'file:///workspace/book.xlsx',
    bindingId: 'binding-frozen',
    revision: 7,
    kind: 'xlsx',
    scope: 'selection',
    sheetId: 'sheet-1',
    range: 'B2:D8',
  }
  const ports: OfficeActionPorts = {
    available: () => true,
    active: () => true,
    flush: vi.fn(async () => {
      events.push('flush')
    }),
    open: vi.fn(async () => {
      events.push('open')
    }),
    save: vi.fn(async (copy) => {
      events.push(copy ? 'saveAs' : 'save')
    }),
    bind: vi.fn(async (scope) => {
      events.push('bind')
      return { ...binding, scope }
    }),
    chat: vi.fn(async () => {
      events.push('chat')
    }),
  }
  return { ports, events, handler: officeActionHandler(ports) }
}

it.each(['open', 'save', 'saveAs'] as const)(
  'flushes the captured editor before native %s',
  async (action) => {
    const { handler, events } = fixture()
    expect(await handler({ action })).toEqual({ completed: true })
    expect(events).toEqual(['flush', action])
  },
)

it('routes a frozen workbook binding to partial native chat and a preset to an immediate request', async () => {
  const { handler, ports, events } = fixture()
  await handler({ action: 'chat', scope: 'selection' })
  expect(events).toEqual(['flush', 'bind', 'chat'])
  const [query, partial] = vi.mocked(ports.chat).mock.calls[0]!
  expect(partial).toBe(true)
  expect(query).toContain('"bindingId":"binding-frozen"')
  expect(query).toContain('"range":"B2:D8"')
  expect(query).toContain('expectedRevision=7')
  expect(query).toContain('genoffice_workbook_context')
  expect(ports.bind).toHaveBeenCalledWith('selection')
  await handler({ action: 'chat', instruction: 'Summarize the workbook', scope: 'document' })
  expect(ports.chat).toHaveBeenLastCalledWith(
    expect.stringContaining('User request:\nSummarize the workbook'),
    false,
  )
})

it('keeps the existing DOCX binding instead of rebinding from focus after chat opens', async () => {
  const { handler, ports } = fixture()
  ports.bind = vi.fn<OfficeActionPorts['bind']>(async (scope) => ({
    kind: 'docx',
    scope,
    documentUri: 'file:///workspace/document.docx',
    bindingId: 'docx-frozen',
    revision: 12,
  }))
  await handler({ action: 'chat', scope: 'selection', instruction: 'Rewrite this' })
  expect(ports.chat).toHaveBeenCalledWith(
    expect.stringContaining('genoffice_docx_read with this bindingId and expectedRevision=12'),
    false,
  )
  expect(vi.mocked(ports.chat).mock.calls[0]![0]).toContain('"bindingId":"docx-frozen"')
})

it('rejects injected URIs, command names, oversized prompts and unknown actions before flushing', async () => {
  const { handler, events } = fixture()
  for (const request of [
    { action: 'save', uri: 'file:///other.xlsx' },
    { action: 'open', command: 'terminal.execute' },
    { action: 'delete' },
    { action: 'chat', scope: 'all' },
    { action: 'chat', instruction: 'x'.repeat(4001) },
  ])
    await expect(handler(request)).rejects.toThrow()
  expect(events).toEqual([])
})

it('rechecks session and active panel after flushing and after binding', async () => {
  for (const phase of ['flush', 'bind']) {
    const { handler, ports } = fixture()
    let available = true
    ports.available = () => available
    if (phase === 'flush')
      ports.flush = vi.fn(async () => {
        available = false
      })
    else
      ports.bind = vi.fn<OfficeActionPorts['bind']>(async () => {
        available = false
        return {
          documentUri: 'file:///old.docx',
          bindingId: 'old',
          kind: 'docx',
          scope: 'document',
          revision: 0,
        }
      })
    await expect(handler({ action: 'chat' })).rejects.toThrow('no longer available')
    expect(ports.chat).not.toHaveBeenCalled()
  }
  const { handler, ports, events } = fixture()
  ports.active = () => false
  await expect(handler({ action: 'save' })).rejects.toThrow('Activate')
  expect(events).toEqual([])
})

it('rejects overlapping menu actions and releases its gate after success or failure', async () => {
  const { handler, ports } = fixture()
  let finish!: () => void
  ports.flush = vi.fn(
    () =>
      new Promise<void>((resolve) => {
        finish = resolve
      }),
  )
  const pending = handler({ action: 'save' })
  await expect(handler({ action: 'saveAs' })).rejects.toThrow('still running')
  finish()
  await pending
  ports.flush = vi.fn(async () => {
    throw new Error('composing')
  })
  await expect(handler({ action: 'save' })).rejects.toThrow('composing')
  ports.flush = vi.fn(async () => {})
  await expect(handler({ action: 'open' })).resolves.toEqual({ completed: true })
})
