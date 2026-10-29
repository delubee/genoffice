import { sessionPivotAdditions } from '@genoffice/xlsx-gateway/domain/session-candidate'
import type { PivotRefreshUpdate } from '@genoffice/xlsx-gateway/gateway/xlsx-pivot-expand'
import type {
  SheetProtectedRangesState,
  SheetProtectionState,
} from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import type { WorkbookPivotAdd } from '@genoffice/xlsx-gateway/shared/workbook-protocol'
import {
  visibilityEditsSchema,
  visibilityStructuralOp,
} from '@genoffice/xlsx-gateway/domain/session-visibility'
import {
  outlineEditSchema,
  outlineStructuralOp,
} from '@genoffice/xlsx-gateway/domain/session-outline'
import type { NameReferenceChange } from '@genoffice/xlsx-gateway/domain/defined-name-references'
import { sessionDefinedNames } from '@genoffice/xlsx-gateway/domain/session-defined-names'
import { copyFile, mkdtemp, rename, rm, unlink } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { z } from 'zod'
import { columnIndex, parseAddress } from '@genoffice/xlsx-gateway/domain/cell-address'
import { sessionMerges } from '@genoffice/xlsx-gateway/domain/session-merges'
import { pixelsToColumnWidth } from '@genoffice/xlsx-gateway/domain/sheet-dimensions'
import { chartKindFromVisual, type SheetVisual } from '@genoffice/xlsx-gateway/domain/chart-visual'
import {
  saveWorkbookViaSidecar,
  type StreamingSaveRequest,
  type ArchiveClient,
} from '@genoffice/xlsx-gateway/gateway/xlsx-package-io'
import type { SheetEditPlan } from '@genoffice/xlsx-gateway/gateway/xlsx-sheets'
import { workbookVisualAddSchema } from '@genoffice/xlsx-gateway/shared/workbook-protocol'
import {
  workbookChartEditSchema,
  workbookVisualEditSchema,
} from '@genoffice/xlsx-gateway/shared/edit-schemas'
import { workbookFilterStateSchema } from '@genoffice/xlsx-gateway/shared/workbook-protocol'
import type { SheetFilterState } from '@genoffice/xlsx-gateway/gateway/xlsx-filter'
import { WorkbookError, assertNotCancelled, type WorkbookState } from './domain'
import type { PackageWriter } from './ports'
import { hashFile } from './resources'
import { workbookStateSchema } from './schemas'
import { workbookOperationSchema } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import type { SheetPageSetupState } from '@genoffice/xlsx-gateway/gateway/xlsx-page-setup'
import type { SheetDvState, SheetNoteState } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import type { DvCellArea, DvWireRule } from '@genoffice/xlsx-gateway/gateway/xlsx-dv'
import { captureFileBaseline, publishStagedWorkbook, type FileBaseline } from './publish'

const metadataSchema = z.object({ sheets: z.array(z.object({ id: z.string(), name: z.string() })) })

/** Every save starts from original bytes and replays the complete current root. */
export class SessionPackageWriter implements PackageWriter {
  publishedBaseline: FileBaseline | undefined
  private metadata: z.infer<typeof metadataSchema>
  constructor(
    private readonly client: ArchiveClient,
    metadata: unknown,
  ) {
    this.metadata = metadataSchema.parse(metadata)
  }
  async stage(
    originalPath: string,
    rawState: Readonly<WorkbookState>,
    targetPath: string,
    signal?: AbortSignal,
  ) {
    assertNotCancelled(signal)
    const state = workbookStateSchema.parse(rawState),
      originalHash = await hashFile(originalPath)
    const names = new Map<string, string>()
    for (const [logicalId, entry] of Object.entries(state.sheetMap))
      if (entry.originalSheetId) {
        const original = this.metadata.sheets.find((sheet) => sheet.id === entry.originalSheetId)
        if (!original)
          throw new WorkbookError(
            'invalid_document',
            'A logical worksheet has no original package mapping.',
          )
        names.set(logicalId, original.name)
      }
    // Compatibility for a root without structural tombstones is limited to
    // nonstructural saves; deleted logical IDs are never guessed by name.
    if (!names.size && !state.structuralLog.length)
      for (const sheet of state.sheets)
        if (sheet.originalSheetId)
          names.set(
            sheet.id,
            this.metadata.sheets.find((item) => item.id === sheet.originalSheetId)!.name,
          )
    const directory = await mkdtemp(join(dirname(targetPath), '.genoffice-stage-'))
    let current = join(directory, '0.xlsx'),
      sequence = 0
    try {
      await copyFile(originalPath, current)
      const writeStage = async (
        request: Omit<StreamingSaveRequest, 'client' | 'sourcePath' | 'targetPath'>,
      ) => {
        assertNotCancelled(signal)
        const next = join(directory, `${++sequence}.xlsx`)
        await saveWorkbookViaSidecar({
          ...request,
          client: this.client,
          sourcePath: current,
          targetPath: next,
        })
        assertNotCancelled(signal)
        await unlink(current)
        current = next
      }
      for (const [index, op] of state.structuralLog.entries()) {
        if (op.op === 'add_sheet') {
          const entry = Object.entries(state.sheetMap).find(
            ([, value]) => value.createdAtLogIndex === index,
          )
          if (!entry)
            throw new WorkbookError(
              'invalid_document',
              'An added sheet is missing its stable allocation.',
            )
          names.set(entry[0], op.name)
          await writeStage({
            edits: [],
            sheetPlan: {
              renames: [],
              additions: [{ name: op.name }],
              removals: [],
              order: [...names.values()],
            },
          })
        } else if (op.op === 'rename_sheet' || op.op === 'delete_sheet') {
          const name = names.get(op.sheetId)
          if (!name)
            throw new WorkbookError(
              'invalid_document',
              'A structural operation references an unknown worksheet.',
            )
          const plan: SheetEditPlan = {
            renames: op.op === 'rename_sheet' ? [{ sheetName: name, newName: op.name }] : [],
            additions: [],
            removals: op.op === 'delete_sheet' ? [name] : [],
            order: [],
          }
          if (op.op === 'delete_sheet') names.delete(op.sheetId)
          else names.set(op.sheetId, op.name)
          await writeStage({ edits: [], sheetPlan: { ...plan, order: [...names.values()] } })
        } else if (
          op.op === 'insert_rows' ||
          op.op === 'delete_rows' ||
          op.op === 'insert_cols' ||
          op.op === 'delete_cols'
        ) {
          const name = names.get(op.sheetId)
          if (!name)
            throw new WorkbookError(
              'invalid_document',
              'A structural operation references an unknown worksheet.',
            )
          await writeStage({
            edits: [],
            structuralOps: [
              {
                sheetName: name,
                ops: [
                  {
                    kind:
                      op.op === 'insert_rows'
                        ? 'insert-rows'
                        : op.op === 'delete_rows'
                          ? 'remove-rows'
                          : op.op === 'insert_cols'
                            ? 'insert-cols'
                            : 'remove-cols',
                    index: 'row' in op ? op.row - 1 : columnIndex(op.column),
                    count: op.count,
                  },
                ],
              },
            ],
          })
        } else
          throw new WorkbookError(
            'unsupported_operation',
            `The package writer cannot replay ${op.op}.`,
          )
      }
      const nameJournal = state.features?.definedNamesEdits ? sessionDefinedNames(state) : null
      const definedNamesState = nameJournal
        ? {
            preserveNames: nameJournal.preserveNames,
            names: nameJournal.names.map(({ name, formula, scopeSheetId }) => {
              const sheetIndex =
                scopeSheetId === undefined
                  ? undefined
                  : state.sheets.findIndex((sheet) => sheet.id === scopeSheetId)
              if (sheetIndex === -1)
                throw new WorkbookError('invalid_document', 'Unknown defined-name scope.')
              return { name, formula, ...(sheetIndex === undefined ? {} : { sheetIndex }) }
            }),
          }
        : null
      const chartEdits: NonNullable<StreamingSaveRequest['chartEdits']>[number][] = []
      const visualEdits: NonNullable<StreamingSaveRequest['visualEdits']>[number][] = []
      const hyperlinkEdits: NonNullable<StreamingSaveRequest['hyperlinkEdits']>[number][] = []
      const edits: StreamingSaveRequest['edits'][number][] = [],
        visualAdditions: NonNullable<StreamingSaveRequest['visualAdditions']>[number][] = [],
        formulaValues: NonNullable<StreamingSaveRequest['formulaValues']>[number][] = []
      const pageSetupStates: SheetPageSetupState[] = [],
        noteStates: SheetNoteState[] = []
      const filterStates: SheetFilterState[] = []
      const sheetProtections: SheetProtectionState[] = []
      const protectedRangeStates: SheetProtectedRangesState[] = []
      const workbookLock = state.features?.workbookProtectionEdits as
        { lockStructure: boolean } | undefined
      const dvStates: SheetDvState[] = []
      for (const sheet of state.sheets) {
        if (names.get(sheet.id) !== sheet.name)
          throw new WorkbookError(
            'invalid_document',
            'The structural log does not produce the current sheet map.',
          )
        for (const [chartPath, patch] of Object.entries(
          (sheet.features?.chartEdits ?? {}) as Record<string, object>,
        )) {
          const metadata = state.features?.originalMetadata as
            { visuals?: { id: string; sheetId: string; chartPath?: string }[] } | undefined
          const matches =
            metadata?.visuals?.filter((visual) => visual.chartPath === chartPath) ?? []
          const native = matches[0]
          if (matches.length !== 1 || !native || native.sheetId !== sheet.originalSheetId)
            throw new WorkbookError('invalid_document', 'Unknown original chart owner.')
          if (
            (sheet.features?.visualEdits as Record<string, { remove?: true }> | undefined)?.[
              native.id
            ]?.remove
          )
            continue
          chartEdits.push(workbookChartEditSchema.parse({ ...patch, chartPath }))
        }
        for (const [visualId, patch] of Object.entries(
          (sheet.features?.visualEdits ?? {}) as Record<string, unknown>,
        )) {
          const original = (
            state.features?.originalMetadata as
              | {
                  visuals?: {
                    id: string
                    sheetId: string
                    drawingPath?: string
                    drawingIndex?: number
                  }[]
                }
              | undefined
          )?.visuals?.find(
            (visual) => visual.id === visualId && visual.sheetId === sheet.originalSheetId,
          )
          if (!original)
            throw new WorkbookError(
              'invalid_document',
              'A visual edit has no original drawing mapping.',
            )
          visualEdits.push(
            workbookVisualEditSchema.parse({
              ...(patch as object),
              drawingPath: original.drawingPath,
              drawingIndex: original.drawingIndex,
            }),
          )
        }
        const heights = z
          .record(z.string().regex(/^[1-9][0-9]*$/), z.number().finite().min(0).max(409.5))
          .parse(sheet.features?.rowHeights ?? {})
        const widths = z
          .record(z.string().regex(/^[A-Z]{1,3}$/), z.number().finite().min(0).max(8000))
          .parse(sheet.features?.colWidths ?? {})
        const sizeOps = [
          ...Object.entries(heights).map(([row, height]) => ({
            kind: 'set-row-size' as const,
            start: Number(row) - 1,
            end: Number(row) - 1,
            size: height,
          })),
          ...Object.entries(widths).map(([column, width]) => ({
            kind: 'set-col-size' as const,
            start: columnIndex(column),
            end: columnIndex(column),
            size:
              ((sheet.features?.columnWidthsCharacters ?? {}) as Record<string, number>)[column] ??
              pixelsToColumnWidth(width),
          })),
        ]
        const outlineOps = z
          .array(outlineEditSchema)
          .parse(sheet.features?.outlineEdits ?? [])
          .map(outlineStructuralOp)
        if (outlineOps.length)
          await writeStage({
            edits: [],
            structuralOps: [{ sheetName: sheet.name, ops: outlineOps }],
          })
        if (sizeOps.length)
          await writeStage({ edits: [], structuralOps: [{ sheetName: sheet.name, ops: sizeOps }] })
        if (sheet.features?.mergeEdits) {
          const original = sessionMerges({ merges: sheet.features.merges ?? [] })
          const current = sessionMerges(sheet.features)
          const ops = [
            ...original.map((range) => ({ kind: 'unmerge-cells' as const, range })),
            ...current.map((range) => ({ kind: 'merge-cells' as const, range })),
          ]
          if (ops.length)
            await writeStage({ edits: [], structuralOps: [{ sheetName: sheet.name, ops }] })
        }
        if (sheet.features?.filterEdits) {
          const { sheetId: _id, ...filter } = workbookFilterStateSchema.parse(
            sheet.features.filterEdits,
          )
          filterStates.push({ sheetName: sheet.name, ...filter })
        }
        const validationEdits = sheet.features?.dataValidationEdits as
          { rules?: DvWireRule[]; remove?: DvCellArea[] } | undefined
        if (
          validationEdits &&
          ((validationEdits.rules?.length ?? 0) || validationEdits.remove?.length)
        )
          dvStates.push({
            sheetName: sheet.name,
            rules: validationEdits.rules ?? [],
            append: true,
            remove: validationEdits.remove ?? [],
          })
        if (sheet.features?.pageSetupEdits || sheet.features?.freezeEdits) {
          const page = sheet.features.pageSetupEdits
            ? workbookOperationSchema.parse({
                ...(sheet.features.pageSetupEdits as object),
                op: 'set_page_setup',
                sheetId: sheet.id,
              })
            : undefined
          const freeze = sheet.features.freezeEdits
            ? z
                .object({
                  frozenRows: z.number().int().min(0).max(100),
                  frozenColumns: z.number().int().min(0).max(100),
                })
                .parse(sheet.features.freezeEdits)
            : undefined
          if (page && page.op !== 'set_page_setup')
            throw new WorkbookError('invalid_document', 'Invalid page settings.')
          const {
            op: _op,
            sheetId: _id,
            ...settings
          } = page ?? { op: undefined, sheetId: undefined }
          pageSetupStates.push({ sheetName: sheet.name, ...settings, ...freeze })
        }
        const protection = sheet.features?.protectionEdits as
          (Omit<SheetProtectionState, 'sheetName'> & { hasPassword: boolean }) | undefined
        const allowRanges = sheet.features?.protectedRangeEdits as
          { name: string; sqref: string }[] | undefined
        if (allowRanges) protectedRangeStates.push({ sheetName: sheet.name, ranges: allowRanges })
        if (protection) {
          const { hasPassword: _hasPassword, ...write } = protection
          sheetProtections.push({ sheetName: sheet.name, ...write })
        }
        if (sheet.features?.hyperlinkEdits) {
          hyperlinkEdits.push({
            sheetName: sheet.name,
            edits: Object.entries(sheet.features.hyperlinkEdits).map(([address, target]) => {
              if (target !== null && typeof target !== 'string')
                throw new WorkbookError('invalid_document', 'Invalid hyperlink target.')
              const cell = parseAddress(address)
              return { row: cell.row, column: cell.column, target }
            }),
          })
        }
        if (sheet.features?.noteEdits || sheet.features?.threadEdits) {
          const original =
            (sheet.features.originalMetadata as { comments?: SheetNoteState['notes'] } | undefined)
              ?.comments ?? []
          const notes = new Map(
            original.map((note) => [String(note.row) + ':' + String(note.column), note]),
          )
          // Threads first: a thread converted to a note keeps its author.
          const threadAuthors = new Map<string, string>()
          for (const [address, thread] of Object.entries(
            (sheet.features.threadEdits ?? {}) as Record<
              string,
              (NonNullable<SheetNoteState['notes'][number]['thread']> & { text: string }) | null
            >,
          )) {
            const point = parseAddress(address),
              key = `${point.row}:${point.column}`
            const existing = notes.get(key)
            if (thread === null) {
              if (existing?.thread) threadAuthors.set(key, existing.thread.author)
              notes.delete(key)
              continue
            }
            const { text, ...meta } = thread
            notes.set(key, { ...point, author: thread.author, text, thread: meta })
          }
          const changes = z
            .record(z.string(), z.string().min(1).max(32000).nullable())
            .parse(sheet.features.noteEdits ?? {})
          for (const [address, text] of Object.entries(changes)) {
            const point = parseAddress(address),
              key = `${point.row}:${point.column}`
            if (text === null) notes.delete(key)
            else {
              const existing = notes.get(key)
              notes.set(key, {
                ...existing,
                ...point,
                author: existing?.author ?? threadAuthors.get(key) ?? 'GenOffice',
                text,
                ...(existing?.thread || threadAuthors.has(key) ? { thread: null } : {}),
              })
            }
          }
          noteStates.push({ sheetName: sheet.name, notes: [...notes.values()] })
        }
        const formulaCells: NonNullable<
          StreamingSaveRequest['formulaValues']
        >[number]['cells'][number][] = []
        for (const [address, cell] of Object.entries(sheet.cells)) {
          const point = parseAddress(address)
          edits.push({
            sheetName: sheet.name,
            row: point.row,
            column: point.column,
            writeValue: true,
            ...(cell.styleReset ? { styleReset: true } : {}),
            cell: {
              value: cell.formula && cell.calculationState !== 'ready' ? null : cell.rawValue,
              ...(cell.kind === 'error' ? { error: true } : {}),
              ...(cell.formula ? { formula: cell.formula } : {}),
            },
            ...(cell.style || cell.numberFormat
              ? {
                  style: {
                    ...cell.style,
                    ...(cell.numberFormat ? { numberFormat: cell.numberFormat } : {}),
                  },
                }
              : {}),
          })
          if (cell.formula && cell.calculationState === 'ready' && cell.cachedValue !== undefined)
            formulaCells.push({
              row: point.row,
              column: point.column,
              value: cell.kind === 'error' ? { error: String(cell.cachedValue) } : cell.cachedValue,
            })
        }
        if (formulaCells.length) formulaValues.push({ sheetName: sheet.name, cells: formulaCells })
        for (const visual of (sheet.features?.visuals ??
          sheet.features?.charts ??
          []) as unknown as SheetVisual[]) {
          if (visual.kind === 'image') {
            const addition = workbookVisualAddSchema.parse({
              sheetId: sheet.id,
              anchor: visual.anchor,
              image: {
                mediaType: visual.mediaType,
                base64: visual.mediaDataUrl.split(',')[1],
                altText: visual.altText,
                editAs: visual.editAs,
                rotation: visual.rotation,
                flipH: visual.flipH,
                flipV: visual.flipV,
                ...(visual.frameWidth && visual.frameHeight
                  ? { frameSize: { width: visual.frameWidth, height: visual.frameHeight } }
                  : {}),
              },
            })
            visualAdditions.push({
              sheetName: sheet.name,
              anchor: addition.anchor,
              image: addition.image,
            })
            continue
          }
          if (visual.kind === 'shape') {
            const addition = workbookVisualAddSchema.parse({
              sheetId: sheet.id,
              anchor: visual.anchor,
              shape: {
                shapeType: visual.shapeType,
                text: visual.text,
                isTextBox: visual.isTextBox,
                fillColor: visual.fillColor,
                lineColor: visual.lineColor,
                altText: visual.altText,
                editAs: visual.editAs,
                rotation: visual.rotation,
                flipH: visual.flipH,
                flipV: visual.flipV,
                ...(visual.frameWidth && visual.frameHeight
                  ? { frameSize: { width: visual.frameWidth, height: visual.frameHeight } }
                  : {}),
              },
            })
            visualAdditions.push({
              sheetName: sheet.name,
              anchor: addition.anchor,
              shape: addition.shape,
            })
            continue
          }
          const chart = visual.chart
          const chartType = chartKindFromVisual(chart)
          if (!chartType) throw new WorkbookError('unsupported_operation', 'Unknown chart type.')
          const addition = workbookVisualAddSchema.parse({
            sheetId: sheet.id,
            anchor: visual.anchor,
            chart: {
              altText: visual.altText,
              editAs: visual.editAs,
              chartType,
              title: chart.title,
              series: chart.series.map((series) => ({
                name: series.name,
                categories: series.categories,
                values: series.values,
                ...(series.explosionPct === undefined ? {} : { explosionPct: series.explosionPct }),
                ...(series.pointExplosions?.length
                  ? {
                      pointExplosions: Object.fromEntries(
                        series.pointExplosions.map((point) => [String(point.index), point.pct]),
                      ),
                    }
                  : {}),
                ...(series.color ? { color: series.color } : {}),
                ...(series.pointColors?.length
                  ? {
                      pointColors: Object.fromEntries(
                        series.pointColors.map((point) => [String(point.index), point.color]),
                      ),
                    }
                  : {}),
                ...(series.valuesRef ? { valuesRef: series.valuesRef } : {}),
                ...(series.categoriesRef ? { categoriesRef: series.categoriesRef } : {}),
              })),
              ...(chart.dataLabelPosition === undefined
                ? {}
                : { dataLabelPosition: chart.dataLabelPosition }),
              ...(chart.dataLabelFormat === undefined
                ? {}
                : { dataLabelFormat: chart.dataLabelFormat }),
              ...(chart.grouping === undefined ? {} : { grouping: chart.grouping }),
              ...(chart.gridlines === undefined ? {} : { gridlines: chart.gridlines }),
              ...(chart.valueAxis === undefined ? {} : { valueAxis: chart.valueAxis }),
              ...(chart.gapWidthPct === undefined ? {} : { gapWidthPct: chart.gapWidthPct }),
              ...(chart.holeSizePct === undefined ? {} : { holeSizePct: chart.holeSizePct }),
              ...(chart.axisTitles
                ? {
                    axisTitles: Object.fromEntries(
                      Object.entries(chart.axisTitles).filter(
                        ([, value]) => typeof value === 'string',
                      ),
                    ),
                  }
                : {}),
              ...(chart.legend ? { legend: chart.legend } : {}),
              ...(chart.dataLabels ? { dataLabels: chart.dataLabels } : {}),
            },
          })
          visualAdditions.push({
            sheetName: sheet.name,
            anchor: addition.anchor,
            chart: addition.chart,
          })
        }
      }
      // Refreshed file pivots: refreshOnLoad for their caches and widened locations. The
      // writer checks growth areas against its input, so they share the cell-edit stage.
      const pivotCacheRefreshPaths: string[] = [],
        pivotRefreshUpdates: PivotRefreshUpdate[] = []
      for (const sheet of state.sheets) {
        const refresh = (sheet.features?.pivotRefresh ?? {}) as {
          cachePaths?: string[]
          outputRefs?: Record<string, string>
        }
        pivotCacheRefreshPaths.push(...(refresh.cachePaths ?? []))
        const relayouts = Object.values(
          (sheet.features?.pivotRelayouts ?? {}) as Record<
            string,
            WorkbookPivotAdd & { cachePath: string }
          >,
        )
        for (const [cachePath, newOutputRef] of Object.entries(refresh.outputRefs ?? {})) {
          const edited = relayouts.find((item) => item.cachePath === cachePath)
          let relayout: NonNullable<PivotRefreshUpdate['relayout']> | undefined
          if (edited) {
            const { sheetId: _sheet, sourceSheetId, cachePath: _cache, ...layout } = edited
            const origin = state.sheets.find((item) => item.id === sourceSheetId)
            if (!origin)
              throw new WorkbookError('invalid_document', 'An edited PivotTable lost its source.')
            relayout = { ...layout, sourceSheetName: origin.name }
          }
          pivotRefreshUpdates.push({
            cachePath,
            sheetName: sheet.name,
            newOutputRef,
            ...(relayout ? { relayout } : {}),
          })
        }
      }
      if (
        edits.length ||
        sheetProtections.length ||
        protectedRangeStates.length ||
        workbookLock !== undefined ||
        pivotCacheRefreshPaths.length ||
        pivotRefreshUpdates.length ||
        hyperlinkEdits.length ||
        visualAdditions.length ||
        visualEdits.length ||
        chartEdits.length ||
        state.structuralLog.length ||
        pageSetupStates.length ||
        noteStates.length ||
        filterStates.length ||
        dvStates.length ||
        definedNamesState !== null
      )
        await writeStage({
          edits,
          sheetProtections,
          protectedRangeStates,
          workbookProtectionState: workbookLock ?? null,
          pivotCacheRefreshPaths,
          pivotRefreshUpdates,
          hyperlinkEdits,
          visualAdditions: visualAdditions.slice(0, 50),
          visualEdits,
          chartEdits,
          formulaValues,
          pageSetupStates,
          noteStates,
          filterStates,
          dvStates,
          definedNamesState,
          nameReferencePlan: state.features?.definedNameReferences
            ? {
                changes: state.features.definedNameReferences as unknown as NameReferenceChange[],
                sheets: state.sheets.map(({ id, name }) => ({ id, name })),
              }
            : null,
          invalidateFormulaCaches: true,
        })
      for (const sheet of state.sheets) {
        const ops = visibilityEditsSchema
          .parse(sheet.features?.visibilityEdits ?? [])
          .map(visibilityStructuralOp)
        if (ops.length)
          await writeStage({ edits: [], structuralOps: [{ sheetName: sheet.name, ops }] })
      }
      for (let start = 50; start < visualAdditions.length; start += 50)
        await writeStage({ edits: [], visualAdditions: visualAdditions.slice(start, start + 50) })
      // New PivotTables pin final coordinates and names, so they follow every
      // structural and sheet stage in a stage of their own.
      const pivots = sessionPivotAdditions(state).map(({ sheetId, sourceSheetId, ...pivot }) => {
        const owner = state.sheets.find((sheet) => sheet.id === sheetId),
          origin = state.sheets.find((sheet) => sheet.id === sourceSheetId)
        if (!owner || !origin)
          throw new WorkbookError('invalid_document', 'A new PivotTable lost its worksheet.')
        return { ...pivot, sheetName: owner.name, sourceSheetName: origin.name }
      })
      for (let start = 0; start < pivots.length; start += 20)
        await writeStage({ edits: [], pivotAdditions: pivots.slice(start, start + 20) })
      assertNotCancelled(signal)
      if ((await hashFile(originalPath)) !== originalHash)
        throw new WorkbookError(
          'invalid_document',
          'The immutable source package changed while saving.',
        )
      const stagedPath = join(directory, 'ready.xlsx')
      await rename(current, stagedPath)
      return { stagedPath, hash: await hashFile(stagedPath) }
    } catch (error) {
      await rm(directory, { recursive: true, force: true })
      throw error
    }
  }
  async publish(
    stagedPath: string,
    targetPath: string,
    expectedHash: string | null,
    baseline?: FileBaseline,
  ) {
    const expected = baseline ?? (await captureFileBaseline(targetPath))
    if (expected.sha256 !== expectedHash)
      throw new WorkbookError(
        'external_change_conflict',
        'The save target changed; the staged workbook has been preserved.',
        { stagedPath, targetPath },
      )
    this.publishedBaseline = await publishStagedWorkbook({
      stagedPath,
      targetPath,
      baseline: expected,
    })
    await rm(dirname(stagedPath), { recursive: true, force: true }).catch(() => {})
  }
}
