import { t } from '../../../apps/sheets/src/renderer/i18n/locale'
import { CommandType, ICommandService, type IMutationInfo } from '@univerjs/core'
import {
  ISheetClipboardService,
  SheetClipboardController,
  type ISheetClipboardHook,
} from '@univerjs/preset-sheets-core'
import { formatAddress } from '@genoffice/xlsx-gateway/domain/cell-address'
import {
  workbookOperationSchema,
  type WorkbookOperation,
} from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { toNeutralStyle, plainCellValue } from '../../../apps/sheets/src/renderer/edit-journal'
import type { UniverRuntime } from '../../../apps/sheets/src/renderer/univer-state'
import { manualValueIntents } from './manual-intents'

const planning = new WeakSet<object>()
/** Let the desktop parser and hooks build their complete mutation plan without
 * applying any of it to the projection or its independent undo stack. */
export async function planNativeCellPaste(
  runtime: UniverRuntime,
  html: string,
  text: string,
  expected: { sheetId: string; row: number; column: number },
  pasteType?: string,
): Promise<WorkbookOperation[]> {
  const injector = runtime.univer.__getInjector()
  // Toolbar/context-menu-free embeddings do not eagerly instantiate this controller.
  injector.get(SheetClipboardController)
  const clipboard = injector.get(ISheetClipboardService),
    commands = injector.get(ICommandService)
  if (planning.has(clipboard)) throw new Error(t('appClipboardBusy'))
  planning.add(clipboard)
  const mutations: IMutationInfo[] = [],
    restore: (() => void)[] = []
  const id = `genoffice.operation.plan-paste-${crypto.randomUUID()}`
  let reached = false,
    targetError: Error | undefined
  const block = commands.registerCommand({ id, type: CommandType.OPERATION, handler: () => false })
  const stop = () => {
    reached = true
    return { undos: [], redos: [{ id, params: {} }] }
  }
  const hook: ISheetClipboardHook = {
    id: `genoffice-canonical-paste-${id}`,
    priority: Number.MAX_SAFE_INTEGER,
    onBeforePaste(location) {
      if (
        location.subUnitId !== expected.sheetId ||
        location.range.rows[0] !== expected.row ||
        location.range.cols[0] !== expected.column
      ) {
        targetError = new Error(t('appClipboardSelectionChanged'))
        return false
      }
      return true
    },
    onPasteCells: stop,
    onPastePlainText: stop,
  }
  try {
    for (const native of clipboard.getClipboardHooks()) {
      for (const key of [
        'onPasteRows',
        'onPasteColumns',
        'onPasteCells',
        'onPastePlainText',
      ] as const) {
        const original = native[key]
        if (!original) continue
        // The public hook signatures differ, but each returns a mutation plan.
        const replacement = (...args: unknown[]) => {
          const result = (
            original as (...args: unknown[]) => { redos: IMutationInfo[]; undos: IMutationInfo[] }
          ).apply(native, args)
          mutations.push(...result.redos)
          return { redos: [], undos: [] }
        }
        Object.assign(native, { [key]: replacement })
        restore.push(() => Object.assign(native, { [key]: original }))
      }
    }
    const installed = clipboard.addClipboardHook(hook)
    restore.push(() => installed.dispose())
    if (pasteType) {
      const data: Record<string, Blob> = {}
      if (html) data['text/html'] = new Blob([html], { type: 'text/html' })
      if (text) data['text/plain'] = new Blob([text], { type: 'text/plain' })
      await clipboard.paste(new ClipboardItem(data), pasteType)
    } else await clipboard.legacyPaste(html, text)
    if (targetError) throw targetError
    if (!reached || !mutations.length) throw new Error(t('appClipboardUnsupported'))
    return nativePasteOperations(mutations, expected.sheetId)
  } finally {
    for (const reset of restore.reverse()) reset()
    block.dispose()
    planning.delete(clipboard)
  }
}

/** Preserve the native parser's scalar/formula/style and merge semantics as one
 * authoritative batch. Unsupported mutation families reject before any write. */
export function nativePasteOperations(
  mutations: readonly IMutationInfo[],
  sheetId: string,
): WorkbookOperation[] {
  const cells = new Map<
    string,
    { row: number; column: number; data: Record<string, unknown>; reset: boolean }
  >()
  const before: WorkbookOperation[] = [],
    after: WorkbookOperation[] = []
  for (const mutation of mutations) {
    const params = mutation.params as {
      subUnitId?: string
      cellValue?: Record<string, Record<string, Record<string, unknown> | null>>
      ranges?: { startRow: number; endRow: number; startColumn: number; endColumn: number }[]
    }
    if (params.subUnitId !== sheetId) throw new Error('invalid_manual_target')
    if (mutation.id === 'sheet.mutation.set-range-values') {
      for (const [row, columns] of Object.entries(params.cellValue ?? {}))
        for (const [column, value] of Object.entries(columns)) {
          const r = Number(row),
            c = Number(column)
          if (!Number.isInteger(r) || !Number.isInteger(c) || r < 0 || c < 0)
            throw new Error('invalid_manual_target')
          const address = formatAddress(r, c),
            current = cells.get(address) ?? { row: r, column: c, data: {}, reset: false }
          if (value === null) {
            current.data = { v: null, f: null, p: null }
            current.reset = true
          } else {
            if (Object.hasOwn(value, 's') && value.s === null) current.reset = true
            current.data = { ...current.data, ...value }
          }
          cells.set(address, current)
        }
      if (cells.size > 2000) throw new Error('range_limit_exceeded')
      continue
    }
    if (
      mutation.id === 'sheet.mutation.add-worksheet-merge' ||
      mutation.id === 'sheet.mutation.remove-worksheet-merge'
    ) {
      const adding = mutation.id === 'sheet.mutation.add-worksheet-merge'
      for (const range of params.ranges ?? [])
        (adding ? after : before).push(
          workbookOperationSchema.parse({
            op: adding ? 'merge_cells' : 'unmerge_cells',
            sheetId,
            range: `${formatAddress(range.startRow, range.startColumn)}:${formatAddress(range.endRow, range.endColumn)}`,
          }),
        )
      continue
    }
    throw new Error(t('appClipboardUnsupported'), { cause: new Error(mutation.id) })
  }
  const result = [...before]
  for (const [address, cell] of cells) {
    const raw = cell.data
    if (Object.hasOwn(raw, 'v')) raw.v = plainCellValue(raw.v, raw.t)
    const rich = raw.p as
      { body?: { textRuns?: { st: number; ed: number }[]; dataStream?: string } } | undefined
    if (
      rich?.body?.textRuns?.some(
        (run) => run.st > 0 || run.ed < (rich.body!.dataStream?.replace(/\r?\n$/, '').length ?? 0),
      )
    )
      throw new Error(t('appClipboardUnsupported'))
    if (['v', 'f', 'p'].some((key) => Object.hasOwn(raw, key)))
      result.push(...manualValueIntents(sheetId, { value: { [cell.row]: { [cell.column]: raw } } }))
    if (cell.reset)
      result.push({
        op: 'format_range',
        sheetId,
        range: address,
        reset: true,
        format: { numberFormat: 'General' },
      })
    if (raw.s !== null && raw.s !== undefined) {
      if (typeof raw.s !== 'object') throw new Error(t('appClipboardUnsupported'))
      const style = toNeutralStyle(raw.s as Record<string, unknown>)
      if (style) {
        const { horizontalAlignment, verticalAlignment, textRotation, ...format } = style
        result.push(
          workbookOperationSchema.parse({
            op: 'format_range',
            sheetId,
            range: address,
            format: {
              ...format,
              ...(horizontalAlignment ? { horizontalAlign: horizontalAlignment } : {}),
              ...(verticalAlignment ? { verticalAlign: verticalAlignment } : {}),
              ...(textRotation === undefined
                ? {}
                : {
                    textRotation:
                      textRotation === 255
                        ? 'vertical'
                        : textRotation > 90
                          ? 90 - textRotation
                          : textRotation,
                  }),
            },
          }),
        )
      }
    }
  }
  return [...result, ...after]
}
