import { z } from 'zod'
import { workbookOperationSchema } from '../domain/workbook-dsl'
import { workbookStyleEditSchema } from './edit-schemas'
import { parseAddress } from '../domain/cell-address'
import { MAX_GRID_COLUMNS, MAX_GRID_ROWS } from './grid-bounds'

const id = z.string().min(1).max(256)
const revision = z.number().int().nonnegative().max(Number.MAX_SAFE_INTEGER)
const scalar = z.union([z.string(), z.number().finite(), z.boolean(), z.null()])
const calculation = z.enum(['ready', 'pending', 'error'])

export const addressSchema = z
  .string()
  .regex(/^[A-Z]{1,3}[1-9][0-9]{0,6}$/)
  .refine((value) => {
    try {
      const parsed = parseAddress(value)
      return parsed.row < MAX_GRID_ROWS && parsed.column < MAX_GRID_COLUMNS
    } catch {
      return false
    }
  }, 'Address exceeds the Excel grid')

export const cellRecordSchema = z
  .strictObject({
    kind: z.enum(['blank', 'string', 'number', 'boolean', 'date', 'error']),
    rawValue: scalar,
    formula: z.string().startsWith('=').optional(),
    calculationState: calculation.optional(),
    cachedValue: scalar.optional(),
    numberFormat: z.string().optional(),
    dateSystem: z.enum(['1900', '1904']).optional(),
    style: workbookStyleEditSchema.optional(),
  })
  .superRefine((cell, ctx) => {
    const valid =
      cell.kind === 'blank'
        ? cell.rawValue === null
        : cell.kind === 'number' || cell.kind === 'date'
          ? typeof cell.rawValue === 'number'
          : cell.kind === 'boolean'
            ? typeof cell.rawValue === 'boolean'
            : typeof cell.rawValue === 'string'
    if (!valid) ctx.addIssue({ code: 'custom', message: 'Cell kind does not match its raw value' })
    if (cell.kind === 'date' && !cell.dateSystem)
      ctx.addIssue({ code: 'custom', message: 'Dates require an explicit date system' })
  })

export const sheetRecordSchema = z.strictObject({
  id,
  name: z.string().min(1).max(31),
  rows: z.number().int().min(0).max(MAX_GRID_ROWS),
  columns: z.number().int().min(0).max(MAX_GRID_COLUMNS),
  cells: z.record(addressSchema, cellRecordSchema),
  originalSheetId: id.optional(),
  partPath: z.string().optional(),
  readOnly: z.boolean().optional(),
  protected: z.boolean().optional(),
  features: z.record(z.string(), z.json()).optional(),
})

export const workbookStateSchema = z
  .strictObject({
    schemaVersion: z.literal(1),
    stateId: id,
    dateSystem: z.enum(['1900', '1904']).default('1900'),
    sheets: z.array(sheetRecordSchema).min(1),
    features: z.record(z.string(), z.json()).optional(),
    structuralLog: z.array(workbookOperationSchema),
    sheetMap: z
      .record(
        z.string(),
        z.strictObject({
          originalSheetId: z.string().optional(),
          originalName: z.string().optional(),
          createdName: z.string().optional(),
          createdAtLogIndex: z.number().int().nonnegative().optional(),
        }),
      )
      .default({}),
  })
  .superRefine((state, ctx) => {
    if (new Set(state.sheets.map((sheet) => sheet.id)).size !== state.sheets.length)
      ctx.addIssue({ code: 'custom', message: 'Duplicate logical sheet ids' })
    if (new Set(state.sheets.map((sheet) => sheet.name.toLowerCase())).size !== state.sheets.length)
      ctx.addIssue({ code: 'custom', message: 'Duplicate worksheet names' })
  })

export const transactionInputSchema = z.strictObject({
  schemaVersion: z.literal(1),
  sessionId: id,
  transactionId: id,
  bindingId: id.optional(),
  expectedRevision: revision,
  origin: z.enum(['manual', 'agent', 'import', 'history']),
  summary: z.string().min(1).max(2048),
  operations: z.array(workbookOperationSchema).min(1),
})

export const receiptSchema = z.strictObject({
  sessionId: id,
  transactionId: id,
  status: z.literal('committed'),
  revision,
  stateId: id,
  previousRevision: revision,
  affectedTargets: z.array(z.string()),
  summary: z.string(),
  calculationState: calculation,
})

export const bindingSchema = z.strictObject({
  bindingId: id,
  sessionId: id,
  documentUri: z.string().min(1),
  logicalSheetId: id,
  frozenRange: z.string().nullable(),
  boundRevision: revision,
  createdAt: z.number().finite(),
  status: z.enum(['active', 'stale', 'invalidated']),
  selectionStatus: z.enum(['current', 'stale', 'selection_stale_structure']).default('current'),
})

export const protocolMessageSchema = z.strictObject({
  protocolVersion: z.literal(1),
  requestId: id,
  sessionId: id,
  method: z.enum([
    'view.ready',
    'view.flush',
    'view.selection',
    'view.presentation',
    'view.intent',
    'candidate.prepare',
    'candidate.result',
    'state.committed',
    'view.resync',
    'history.request',
    'source.read',
    'view.dispose',
  ]),
  payload: z.json(),
})
