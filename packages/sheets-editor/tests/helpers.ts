/** Count candidate resources so disposal tests can detect leaked runtimes. */
export class ResourceCounter {
  private active = new Set<symbol>()
  acquire() {
    const id = Symbol('candidate')
    this.active.add(id)
    let disposed = false
    return () => {
      if (disposed) return
      disposed = true
      this.active.delete(id)
    }
  }
  get size() {
    return this.active.size
  }
}
