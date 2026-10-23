import { copyFile, mkdtemp, rename, rm, unlink } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { z } from 'zod'
import { columnIndex, parseAddress } from '@genoffice/xlsx-gateway/domain/cell-address'
import type { SheetVisual } from '@genoffice/xlsx-gateway/domain/chart-visual'
import {
  saveWorkbookViaSidecar,
  type StreamingSaveRequest,
  type ArchiveClient,
} from '@genoffice/xlsx-gateway/gateway/xlsx-package-io'
import type { SheetEditPlan } from '@genoffice/xlsx-gateway/gateway/xlsx-sheets'
import { workbookVisualAddSchema } from '@genoffice/xlsx-gateway/shared/workbook-protocol'
import { WorkbookError, assertNotCancelled, type WorkbookState } from './domain'
import type { PackageWriter } from './ports'
import { hashFile } from './resources'
import { workbookStateSchema } from './schemas'
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
      const edits: StreamingSaveRequest['edits'][number][] = [],
        visualAdditions: NonNullable<StreamingSaveRequest['visualAdditions']>[number][] = [],
        formulaValues: NonNullable<StreamingSaveRequest['formulaValues']>[number][] = []
      for (const sheet of state.sheets) {
        if (names.get(sheet.id) !== sheet.name)
          throw new WorkbookError(
            'invalid_document',
            'The structural log does not produce the current sheet map.',
          )
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
        for (const visual of (sheet.features?.charts ?? []) as unknown as SheetVisual[]) {
          const chart = visual.chart
          if (
            chart.chartTypes.length !== 1 ||
            !['barChart', 'lineChart', 'pieChart'].includes(chart.chartTypes[0]!)
          )
            throw new WorkbookError(
              'unsupported_operation',
              'The chart type cannot be serialized by this editor yet.',
            )
          const chartType = chart.chartTypes.includes('pieChart')
            ? 'pie'
            : chart.chartTypes.includes('lineChart')
              ? 'line'
              : chart.barDirection === 'bar'
                ? 'bar'
                : 'column'
          const addition = workbookVisualAddSchema.parse({
            sheetId: sheet.id,
            anchor: visual.anchor,
            chart: {
              chartType,
              title: chart.title,
              series: chart.series.map((series) => ({
                name: series.name,
                categories: series.categories,
                values: series.values,
                ...(series.valuesRef ? { valuesRef: series.valuesRef } : {}),
                ...(series.categoriesRef ? { categoriesRef: series.categoriesRef } : {}),
              })),
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
      if (edits.length || visualAdditions.length || state.structuralLog.length)
        await writeStage({ edits, visualAdditions, formulaValues, invalidateFormulaCaches: true })
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
