import { z } from 'zod'
export const operationSchema = z.discriminatedUnion('type', [
  z
    .object({
      type: z.literal('replaceText'),
      blockId: z.string().min(1),
      start: z.number().int().nonnegative(),
      end: z.number().int().nonnegative(),
      text: z.string().max(100000),
    })
    .strict(),
  z
    .object({
      type: z.literal('formatParagraph'),
      blockId: z.string().min(1),
      align: z.enum(['left', 'center', 'right', 'justify']),
    })
    .strict(),
  z
    .object({
      type: z.literal('setHeading'),
      blockId: z.string().min(1),
      level: z.number().int().min(0).max(9),
    })
    .strict(),
  z
    .object({
      type: z.literal('setList'),
      blockId: z.string().min(1),
      kind: z.enum(['bullet', 'ordered']),
    })
    .strict(),
  z
    .object({
      type: z.literal('insertParagraph'),
      blockId: z.string().min(1),
      text: z.string().max(100000),
    })
    .strict(),
  z.object({ type: z.literal('deleteParagraph'), blockId: z.string().min(1) }).strict(),
  z
    .object({
      type: z.literal('formatText'),
      blockId: z.string().min(1),
      start: z.number().int().nonnegative(),
      end: z.number().int().nonnegative(),
      format: z
        .object({
          bold: z.boolean().optional(),
          italic: z.boolean().optional(),
          underline: z.boolean().optional(),
          strike: z.boolean().optional(),
          font: z.string().min(1).max(200).optional(),
          color: z
            .string()
            .regex(/^[a-f0-9]{6}$/i)
            .optional(),
          sizeHalfPoints: z.number().min(2).max(800).optional(),
        })
        .strict()
        .refine((value) => Object.keys(value).length > 0),
    })
    .strict(),
])
export const operationsSchema = z.array(operationSchema).min(1).max(100)
export type DocxOperation = z.infer<typeof operationSchema>

// Marker text and fonts are written into numbering.xml attributes.
const safeText = (max: number) =>
  z
    .string()
    .max(max)
    // eslint-disable-next-line no-control-regex -- C0 controls never serialize.
    .refine((value) => !/[<>&\x00-\x1f]/.test(value), 'Invalid list text')
const levelSchema = z
  .object({
    numFmt: z.string().regex(/^[A-Za-z]{1,40}$/),
    customFormat: safeText(80).optional(),
    lvlText: safeText(40),
    indentLeft: z.number().int().min(-31680).max(31680),
    hanging: z.number().int().min(-31680).max(31680).optional(),
    start: z.number().int().min(0).max(32767).optional(),
    suff: z.enum(['tab', 'space', 'nothing']).optional(),
    lvlJc: z.enum(['left', 'center', 'right']).optional(),
    tabStop: z.number().int().min(0).max(31680).optional(),
    font: safeText(200).optional(),
    szHalfPoints: z.number().int().min(2).max(3276).optional(),
    color: z
      .string()
      .regex(/^[0-9a-f]{6}$/i)
      .optional(),
    bold: z.boolean().optional(),
    italic: z.boolean().optional(),
    pStyle: z
      .string()
      .regex(/^[A-Za-z0-9_.-]{1,253}$/)
      .optional(),
    isLgl: z.boolean().optional(),
    lvlRestart: z.number().int().min(0).max(9).optional(),
  })
  .strict()
const numIdSchema = z.string().regex(/^[1-9][0-9]{0,8}$/)
export const pendingListDefinitionsSchema = z
  .object({
    newDefs: z
      .array(
        z
          .object({
            numId: numIdSchema,
            kind: z.enum(['bullet', 'ordered']),
            levels: z.array(levelSchema).min(1).max(9).optional(),
          })
          .strict(),
      )
      .max(200),
    restartNums: z
      .array(
        z
          .object({
            numId: numIdSchema,
            abstractNumId: z.string().regex(/^[0-9]{1,9}$/),
            startOverrides: z.record(
              z.string().regex(/^[0-8]$/),
              z.number().int().min(0).max(32767),
            ),
          })
          .strict(),
      )
      .max(200),
  })
  .strict()

const SECTION_KEYS = new Set([
  'pageWidth',
  'pageHeight',
  'orientation',
  'marginTop',
  'marginRight',
  'marginBottom',
  'marginLeft',
  'marginTopFixed',
  'marginBottomFixed',
  'gutter',
  'gutterAtTop',
  'pageBorder',
  'pageBorderProps',
  'columns',
  'colSpace',
  'colWidths',
  'lineNumbers',
  'bidi',
  'headerDist',
  'footerDist',
  'vAlign',
  'docGrid',
  'textDirection',
  'footnotePr',
  'endnotePr',
])
/** sectPr values are numbers, flags and identifier-like words; nothing can carry markup. */
const settingValue = (value: unknown, depth: number): boolean => {
  if (value == null || typeof value === 'boolean') return true
  if (typeof value === 'number') return Number.isFinite(value) && Math.abs(value) <= 10_000_000
  if (typeof value === 'string') return /^[A-Za-z0-9 #._-]{0,64}$/.test(value)
  if (depth >= 4) return false
  if (Array.isArray(value))
    return value.length <= 64 && value.every((item) => settingValue(item, depth + 1))
  if (typeof value === 'object')
    return Object.entries(value).every(
      ([key, item]) => /^[A-Za-z]{1,40}$/.test(key) && settingValue(item, depth + 1),
    )
  return false
}
const twips = z.number().int().min(-31680).max(31680)
export const documentSettingsSchema = z
  .object({
    section: z
      .object({
        pageWidth: z.number().int().min(1).max(63360),
        pageHeight: z.number().int().min(1).max(63360),
        orientation: z.enum(['portrait', 'landscape']),
        marginTop: twips,
        marginRight: twips,
        marginBottom: twips,
        marginLeft: twips,
        pageBorder: z.boolean(),
      })
      .catchall(z.unknown())
      .refine(
        (section) =>
          Object.entries(section).every(
            ([key, value]) => SECTION_KEYS.has(key) && settingValue(value, 0),
          ),
        'Invalid section settings',
      )
      .optional(),
    pageColor: z
      .string()
      .regex(/^[0-9a-f]{6}$/i)
      .nullable()
      .optional(),
    mirrorMargins: z.boolean().optional(),
    watermark: z
      .string()
      .min(1)
      .max(255)
      // eslint-disable-next-line no-control-regex -- C0 controls never serialize.
      .refine((value) => !/[<>&"\x00-\x1f]/.test(value), 'Invalid watermark')
      .nullable()
      .optional(),
  })
  .strict()
