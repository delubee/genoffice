import type { EditorHost } from './ports'

export async function requestHostHistory(
  host: EditorHost,
  direction: 'undo' | 'redo',
  flush: () => Promise<boolean>,
) {
  if (!(await flush())) throw new Error('editor_busy')
  await host.history(direction)
}
