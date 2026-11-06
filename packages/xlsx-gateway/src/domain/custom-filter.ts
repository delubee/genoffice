import { z } from 'zod'

export const customFilterCriteriaSchema = z.strictObject({
  and: z.boolean().optional(),
  filters: z
    .array(
      z.strictObject({
        val: z.union([z.string().max(32_767), z.number().finite()]),
        operator: z
          .enum([
            'equal',
            'notEqual',
            'greaterThan',
            'greaterThanOrEqual',
            'lessThan',
            'lessThanOrEqual',
          ])
          .optional(),
      }),
    )
    .min(1)
    .max(2),
})

/** The desktop dialog and OOXML writer use default equality for text matching. */
export function normalizedCustomFilters(criteria: z.infer<typeof customFilterCriteriaSchema>) {
  return {
    and: criteria.filters.length > 1 && criteria.and === true,
    filters: criteria.filters.map(({ val, operator }) => ({
      val,
      ...(operator && operator !== 'equal' ? { operator } : {}),
    })),
  }
}

function wildcard(pattern: string): RegExp {
  const body = Array.from(pattern, (character) => {
    if (character === '*') return '.*'
    if (character === '?') return '.'
    return character.replace(/[\\^$.*+?()[\]{}|]/g, '\\$&')
  }).join('')
  return new RegExp(`^${body}$`, 'i')
}

/** Shared desktop/session/CLI condition semantics. Numeric comparisons use raw
 * numbers, default equality uses case-insensitive wildcard text matching, and
 * text notEqual retains the desktop's literal/wildcard behavior. */
export function customFilterPredicate(criteria: {
  readonly and?: boolean | undefined
  readonly filters: readonly {
    readonly val: string | number
    readonly operator?: string | undefined
  }[]
}): (value: string | number | boolean | null) => boolean {
  const predicates = criteria.filters.map(({ val, operator }) => {
    const numeric = typeof val === 'number' || /^-?\d+(\.\d+)?$/.test(val)
    if (operator === 'notEqual' && !numeric) {
      const pattern = /[*?]/.test(String(val)) ? wildcard(String(val)) : null
      return (value: string | number | null) => {
        if (val === ' ') return value !== null
        const text = value === null ? null : String(value)
        return pattern && text ? !pattern.test(text) : text !== val
      }
    }
    if (!operator || operator === 'equal') {
      const pattern = wildcard(String(val))
      return (value: string | number | null) =>
        value === null ? String(val) === '' : pattern.test(String(value))
    }
    if (!numeric) return () => false
    const compare = Number(val)
    return (value: string | number | null) => {
      if (operator === 'notEqual') return typeof value !== 'number' || value !== compare
      if (typeof value !== 'number') return false
      switch (operator) {
        case 'greaterThan':
          return value > compare
        case 'greaterThanOrEqual':
          return value >= compare
        case 'lessThan':
          return value < compare
        case 'lessThanOrEqual':
          return value <= compare
        default:
          return false
      }
    }
  })
  if (predicates.length < 1 || predicates.length > 2)
    throw new Error('invalid_custom_filter_conditions')
  return (value) => {
    const input = typeof value === 'boolean' ? (value ? 'TRUE' : 'FALSE') : value
    return criteria.and
      ? predicates.every((predicate) => predicate(input))
      : predicates.some((predicate) => predicate(input))
  }
}
