import { randomUUID } from 'node:crypto'
import type { SelectionScope } from './read'

export class DocumentBindings {
  private bindings = new Map<
    string,
    { uri: string; sessionId: string; selection?: SelectionScope }
  >()
  bind(uri: string, sessionId: string, selection?: SelectionScope) {
    for (const [bindingId, binding] of this.bindings)
      if (
        binding.uri === uri &&
        binding.sessionId === sessionId &&
        JSON.stringify(binding.selection) === JSON.stringify(selection)
      )
        return bindingId
    const bindingId = randomUUID()
    this.bindings.set(bindingId, {
      uri,
      sessionId,
      ...(selection ? { selection: structuredClone(selection) } : {}),
    })
    if (this.bindings.size > 64) this.bindings.delete(this.bindings.keys().next().value!)
    return bindingId
  }
  resolve(bindingId: string) {
    const value = this.bindings.get(bindingId)
    if (!value) throw new Error('Unknown DOCX binding; request context again')
    return structuredClone(value)
  }
  release(uri: string) {
    for (const [id, binding] of this.bindings) if (binding.uri === uri) this.bindings.delete(id)
  }
}
