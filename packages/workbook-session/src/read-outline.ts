import {
  visibilityEditsSchema,
  axisHiddenOverride,
} from '@genoffice/xlsx-gateway/domain/session-visibility'
import { z } from 'zod'
import {
  workbookRowOutlineResultSchema,
  workbookFilterStateSchema,
} from '@genoffice/xlsx-gateway/shared/workbook-protocol'
import { outlineEditSchema } from '@genoffice/xlsx-gateway/domain/session-outline'
import { currentAxisIndex } from '@genoffice/xlsx-gateway/domain/session-coordinates'
import { WorkbookError, assertNotCancelled, type WorkbookState } from './domain'
import type { WorkbookSource } from './ports'
import { logicalSheet } from './sheet-map'

type Entry = { index: number; level: number; collapsed: boolean; hidden: boolean }
const metadataSchema = z.object({
  outlineSummaryBelow: z.boolean().optional(),
  outlineSummaryRight: z.boolean().optional(),
  columnWidths: z
    .array(
      z.object({
        startColumn: z.number().int().min(0).max(16383),
        endColumn: z.number().int().min(0).max(16383),
        outlineLevel: z.number().int().min(0).max(7).optional(),
        collapsed: z.boolean().optional(),
        hidden: z.boolean().optional(),
      }),
    )
    .default([]),
})

/** Read the complete immutable outline, map source coordinates, then replay the
 * session journal. No viewport cell read is used as a completeness substitute. */
export async function readCurrentOutline(input: {
  state: Readonly<WorkbookState>
  source: WorkbookSource
  sessionId: string
  sheetId: string
  signal?: AbortSignal
}) {
  const sheet = logicalSheet(input.state, input.sheetId)
  const rows = new Map<number, Entry>(),
    cols = new Map<number, Entry>()
  const metadata = metadataSchema.parse(sheet.features?.originalMetadata ?? {})
  let summaryBelow = metadata.outlineSummaryBelow !== false
  let summaryRight = metadata.outlineSummaryRight !== false
  if (sheet.originalSheetId) {
    if (!input.source.readRowOutline)
      throw new WorkbookError(
        'unsupported_operation',
        'The source cannot read complete row outlines.',
      )
    const deadline = Date.now() + 10_000
    for (;;) {
      assertNotCancelled(input.signal)
      const result = workbookRowOutlineResultSchema.parse(
        await input.source.readRowOutline(input.sessionId, sheet.originalSheetId),
      )
      if (result.indexingComplete) {
        for (const row of result.rows) {
          const index = currentAxisIndex(input.state.structuralLog, sheet.id, row.row, 'row')
          if (index !== null)
            rows.set(index, {
              index,
              level: row.level,
              collapsed: row.collapsed ?? false,
              hidden: row.hidden ?? false,
            })
        }
        break
      }
      if (Date.now() >= deadline)
        throw new WorkbookError(
          'index_pending',
          'Wait for complete row outline metadata.',
          {},
          true,
        )
      await new Promise((resolve) => setTimeout(resolve, 25))
    }
    for (const span of metadata.columnWidths) {
      if (!span.outlineLevel && !span.collapsed) continue
      for (let original = span.startColumn; original <= span.endColumn; original++) {
        const index = currentAxisIndex(input.state.structuralLog, sheet.id, original, 'column')
        if (index !== null)
          cols.set(index, {
            index,
            level: span.outlineLevel ?? 0,
            collapsed: span.collapsed ?? false,
            hidden: span.hidden ?? false,
          })
      }
    }
  }
  for (const edit of z.array(outlineEditSchema).parse(sheet.features?.outlineEdits ?? [])) {
    if (edit.sheetId !== sheet.id)
      throw new WorkbookError(
        'invalid_document',
        'Outline journal belongs to a different worksheet.',
      )
    if (edit.op === 'set_outline_settings') {
      summaryBelow = edit.summaryBelow
      summaryRight = edit.summaryRight
      continue
    }
    const entries = edit.axis === 'rows' ? rows : cols
    for (let index = edit.start - 1; index < edit.start + edit.count - 1; index++) {
      const before = entries.get(index)
      entries.set(index, {
        index,
        level: edit.level,
        collapsed: edit.collapsed ?? before?.collapsed ?? false,
        hidden: before?.hidden ?? false,
      })
    }
  }
  if (sheet.features?.filterEdits) {
    const filter = workbookFilterStateSchema.parse(sheet.features.filterEdits)
    const hidden = new Set(filter.hiddenRows)
    for (const entry of rows.values())
      if (
        entry.index > filter.visibilityRange.startRow &&
        entry.index <= filter.visibilityRange.endRow
      )
        entry.hidden = hidden.has(entry.index)
  }
  const visibility = visibilityEditsSchema.parse(sheet.features?.visibilityEdits ?? [])
  for (const [axis, entries] of [
    ['rows', rows],
    ['cols', cols],
  ] as const)
    for (const entry of entries.values())
      entry.hidden = axisHiddenOverride(visibility, axis, entry.index) ?? entry.hidden
  assertNotCancelled(input.signal)
  return {
    stateId: input.state.stateId,
    sheetId: sheet.id,
    summaryBelow,
    summaryRight,
    rows: [...rows.values()].sort((a, b) => a.index - b.index),
    cols: [...cols.values()].sort((a, b) => a.index - b.index),
  }
}
