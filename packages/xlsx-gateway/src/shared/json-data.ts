/** Reject values JSON.stringify would silently drop, coerce or execute. */
export function assertSerializable(value: unknown, ancestors = new Set<object>()): void {
  if (value === null || typeof value === 'string' || typeof value === 'boolean') return
  if (typeof value === 'number' && Number.isFinite(value)) return
  if (typeof value !== 'object' || value === undefined)
    throw new Error('Changesets must contain finite JSON data only')
  if (ancestors.has(value)) throw new Error('Changesets cannot contain cyclic references')
  if (
    !Array.isArray(value) &&
    Object.getPrototypeOf(value) !== Object.prototype &&
    Object.getPrototypeOf(value) !== null
  )
    throw new Error('Changesets cannot contain runtime instances')
  if (Array.isArray(value))
    for (let index = 0; index < value.length; index++)
      if (!Object.hasOwn(value, index)) throw new Error('Changesets cannot contain sparse arrays')
  ancestors.add(value)
  for (const [key, descriptor] of Object.entries(Object.getOwnPropertyDescriptors(value))) {
    if (Array.isArray(value) && key !== 'length' && !/^(0|[1-9][0-9]*)$/.test(key))
      throw new Error('JSON arrays cannot contain named properties')
    if (!descriptor.enumerable && !(Array.isArray(value) && key === 'length'))
      throw new Error('Changesets cannot contain nonenumerable data')
    if (descriptor.get || descriptor.set) throw new Error('Changesets cannot contain accessors')
    assertSerializable(descriptor.value, ancestors)
  }
  if (Object.getOwnPropertySymbols(value).length)
    throw new Error('Changesets cannot contain symbol keys')
  ancestors.delete(value)
}
