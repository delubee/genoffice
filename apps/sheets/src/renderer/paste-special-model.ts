/** Host independent Paste Special options and availability. */
import type { ICellData } from '@univerjs/core'

export const PASTE_SPECIAL_OPEN_SHELL_COMMAND = 'paste-special-open'

export const PASTE_SPECIAL_TYPES = [
  'all',
  'formulas',
  'values',
  'formats',
  'comments',
  'validation',
  'all-source-theme',
  'all-except-borders',
  'col-widths',
  'formulas-numfmt',
  'values-numfmt',
  'all-merge-cf',
] as const
export type PasteSpecialType = (typeof PASTE_SPECIAL_TYPES)[number]

export const PASTE_SPECIAL_OPERATIONS = ['none', 'add', 'subtract', 'multiply', 'divide'] as const
export type PasteSpecialOperation = (typeof PASTE_SPECIAL_OPERATIONS)[number]

export interface PasteSpecialOptions {
  readonly type: PasteSpecialType
  readonly operation: PasteSpecialOperation
  readonly skipBlanks: boolean
  readonly transpose: boolean
  readonly link: boolean
}

/// Source cell as the transform sees it: `v`/`t` are the displayed (computed)
/// value, `f` the materialized formula text, `s` the composed style object.
export type SourceCell = Pick<ICellData, 'v' | 't' | 'f' | 'p' | 's'> | null

export type CellGrid<T> = readonly (readonly T[])[]

export type PasteSpecialSource =
  | {
      readonly kind: 'internal'
      readonly unitId: string
      readonly subUnitId: string
      readonly rows: readonly number[]
      readonly cols: readonly number[]
      readonly cut: boolean
    }
  | { readonly kind: 'text'; readonly rows: CellGrid<string> }
  | { readonly kind: 'none' }

/// Paste types the dialog disables (Univer has no note/validation-only paste).
export const UNSUPPORTED_TYPES: ReadonlySet<PasteSpecialType> = new Set(['comments', 'validation'])

export interface PasteSpecialAvailability {
  readonly types: ReadonlySet<PasteSpecialType>
  readonly operation: boolean
  readonly skipBlanks: boolean
  readonly transpose: boolean
  readonly link: boolean
}

/// What the dialog may offer for a given copy source: a cut only pastes
/// whole, outside text has no formulas/styles/source cells to link to.
export function pasteSpecialAvailability(source: PasteSpecialSource): PasteSpecialAvailability {
  if (source.kind === 'none') {
    return { types: new Set(), operation: false, skipBlanks: false, transpose: false, link: false }
  }
  if (source.kind === 'internal' && source.cut) {
    return {
      types: new Set(['all']),
      operation: false,
      skipBlanks: false,
      transpose: false,
      link: false,
    }
  }
  if (source.kind === 'text') {
    return {
      types: new Set(['all', 'values']),
      operation: false,
      skipBlanks: false,
      transpose: true,
      link: false,
    }
  }
  return {
    types: new Set(PASTE_SPECIAL_TYPES.filter((type) => !UNSUPPORTED_TYPES.has(type))),
    operation: true,
    skipBlanks: true,
    transpose: true,
    link: true,
  }
}
