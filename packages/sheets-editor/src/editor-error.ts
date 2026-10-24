import type { editorText } from './i18n'

export function editorError(error: unknown, text: ReturnType<typeof editorText>): string {
  const message = error instanceof Error ? error.message : String(error)
  if (!message) return ''
  if (message.includes('merge_select_range')) return text.mergeSelectRange
  if (message.includes('merge_partial_overlap')) return text.mergePartialOverlap
  if (message.includes('merge_nonempty_cells')) return text.mergeNonemptyCells
  if (message.includes('unsupported_structural_feature')) return text.structureRestricted
  if (message.includes('active_filter_sort')) return text.clearFilterBeforeSort
  if (message.includes('merged_range_operation')) return text.mergedRangeRestricted
  if (message.includes('filter_range_limit')) return text.filterRangeTooLarge
  if (message.includes('summary_target_not_empty')) return text.summaryNeedsEmptyCell
  if (message.includes('unsupported_note_image')) return text.imageNotesRestricted
  return message
}
