import { z } from 'zod'
import { LIMITS } from '../domain'
import { addressSchema } from '../schemas'
import { parseRange, rangeCellCount } from '@genoffice/xlsx-gateway/domain/cell-address'

const format = { numberFormat: z.string().max(512).optional() }
const scalarSchemas = [
  z.strictObject({ kind: z.literal('blank'), value: z.null(), ...format }),
  z.strictObject({ kind: z.literal('string'), value: z.string().max(32767), ...format }),
  z.strictObject({ kind: z.literal('number'), value: z.number().finite(), ...format }),
  z.strictObject({ kind: z.literal('boolean'), value: z.boolean(), ...format }),
  z.strictObject({
    kind: z.literal('date'),
    serial: z.number().finite(),
    dateSystem: z.enum(['1900', '1904']),
    numberFormat: z.string().min(1).max(512),
  }),
  z.strictObject({
    kind: z.literal('error'),
    excelCode: z.enum([
      '#NULL!',
      '#DIV/0!',
      '#VALUE!',
      '#REF!',
      '#NAME?',
      '#NUM!',
      '#N/A',
      '#GETTING_DATA',
      '#SPILL!',
      '#CALC!',
      '#FIELD!',
      '#BLOCKED!',
      '#CONNECT!',
      '#UNKNOWN!',
      '#BUSY!',
    ]),
    ...format,
  }),
] as const
export const scalarCellSchema = z.discriminatedUnion('kind', scalarSchemas)
export const typedCellSchema = z.discriminatedUnion('kind', [
  ...scalarSchemas,
  z.strictObject({
    kind: z.literal('formula'),
    formula: z.string().startsWith('=').max(32767),
    cached: scalarCellSchema.nullable(),
    calculationState: z.enum(['ready', 'pending', 'error']),
    ...format,
  }),
])
export type TypedCell = z.infer<typeof typedCellSchema>
export const exchangeRecordSchema = z.strictObject({
  rowOffset: z.number().int().nonnegative(),
  columnOffset: z.number().int().nonnegative(),
  cells: z.array(typedCellSchema).min(1).max(16384),
})
export type ExchangeRecord = z.infer<typeof exchangeRecordSchema>
export const shapeSchema = z
  .strictObject({
    rows: z.number().int().min(1).max(1048576),
    columns: z.number().int().min(1).max(16384),
    cellCount: z.number().int().min(1).max(LIMITS.artifactCells),
  })
  .refine(
    (shape) => shape.rows * shape.columns === shape.cellCount,
    'rows * columns must equal cellCount.',
  )
export type ExchangeShape = z.infer<typeof shapeSchema>
const sha = z.string().regex(/^[a-f0-9]{64}$/)
export const dataFileSchema = z.strictObject({
  fileName: z
    .string()
    .min(1)
    .max(128)
    .regex(/^[a-zA-Z0-9_-]+\.ndjson$/),
  sha256: sha,
  bytes: z.number().int().min(1).max(LIMITS.artifactBytes),
  recordCount: z.number().int().min(1).max(LIMITS.artifactCells),
  encoding: z.literal('utf-8'),
  format: z.literal('typed-ndjson-v1'),
})
export type DataFile = z.infer<typeof dataFileSchema>
export const snapshotManifestSchema = z
  .strictObject({
    schemaVersion: z.literal(1),
    kind: z.literal('workbook-snapshot'),
    snapshotId: z.string().min(1),
    sessionId: z.string().min(1),
    documentUri: z.string().min(1),
    sourceRevision: z.number().int().nonnegative(),
    sourceStateId: z.string().min(1),
    sheet: z.strictObject({ logicalSheetId: z.string().min(1), name: z.string().min(1) }),
    range: z.string(),
    shape: shapeSchema,
    dateSystem: z.enum(['1900', '1904']),
    calculationState: z.enum(['ready', 'pending', 'error']),
    data: dataFileSchema,
    complete: z.literal(true),
  })
  .refine((manifest) => {
    try {
      const bounds = parseRange(manifest.range)
      return (
        rangeCellCount(bounds) === manifest.shape.cellCount &&
        bounds.endRow - bounds.startRow + 1 === manifest.shape.rows &&
        bounds.endColumn - bounds.startColumn + 1 === manifest.shape.columns &&
        bounds.endRow < 1048576 &&
        bounds.endColumn < 16384
      )
    } catch {
      return false
    }
  }, 'Range and shape must describe the same Excel rectangle.')
export type SnapshotManifest = z.infer<typeof snapshotManifestSchema>
export const resultManifestSchema = z.strictObject({
  schemaVersion: z.literal(1),
  kind: z.literal('workbook-result'),
  source: z.strictObject({
    snapshotId: z.string().min(1),
    sourceManifestSha256: sha,
    sourceDataSha256: sha,
    sourceRevision: z.number().int().nonnegative(),
  }),
  output: dataFileSchema,
  shape: shapeSchema,
  target: z.union([
    z.strictObject({ sheetId: z.string().min(1), startCell: addressSchema }),
    z.strictObject({
      newSheetName: z
        .string()
        .min(1)
        .max(31)
        .refine(
          (name) =>
            !/[\\/?*[\]:]/.test(name) &&
            name.trim() === name &&
            !name.startsWith("'") &&
            !name.endsWith("'"),
          'Invalid worksheet name.',
        ),
      startCell: addressSchema,
    }),
  ]),
  mode: z.enum(['values', 'formulas']).default('values'),
  complete: z.literal(true),
})
export type ResultManifest = z.infer<typeof resultManifestSchema>

/** Private host input, never a model-facing tool schema. */
export const importBatchSchema = z.strictObject({
  planId: z.string().min(1),
  resultHash: sha,
  mode: z.enum(['values', 'formulas']),
  target: z.union([
    z.strictObject({ sheetId: z.string().min(1) }),
    z.strictObject({ newSheetName: z.string().min(1).max(31) }),
  ]),
  cells: z
    .array(z.strictObject({ address: addressSchema, cell: typedCellSchema }))
    .min(1)
    .max(LIMITS.importBatchCells),
})
export type ImportBatch = z.infer<typeof importBatchSchema>

export const preparedImportSchema = z
  .strictObject({
    planId: z.string().uuid(),
    snapshotId: z.string(),
    bindingId: z.string(),
    sessionId: z.string(),
    initialExpectedRevision: z.number().int().nonnegative(),
    chainExpectedRevision: z.number().int().nonnegative(),
    expiresAt: z.number().finite(),
    status: z.enum(['prepared', 'applying', 'completed', 'failed', 'cancelled', 'stale']),
    nextCursor: z.number().int().min(0).max(1000),
    completedBatchCount: z.number().int().min(0).max(1000),
    completedCount: z.number().int().min(0).max(LIMITS.artifactCells),
    remainingCount: z.number().int().min(0).max(LIMITS.artifactCells),
    batchCount: z.number().int().min(1).max(1000),
    target: resultManifestSchema.shape.target,
    shape: shapeSchema,
    mode: z.enum(['values', 'formulas']),
    resultHash: sha,
    overwriteCellCount: z.number().int().nonnegative(),
    maxBatchCells: z.literal(1000),
    historyPreflight: z.string(),
    maintenanceError: z.string().nullable(),
  })
  .refine(
    (plan) =>
      plan.completedCount + plan.remainingCount === plan.shape.cellCount &&
      plan.completedBatchCount === plan.nextCursor &&
      plan.nextCursor <= plan.batchCount,
  )
export type PreparedImport = z.infer<typeof preparedImportSchema>
