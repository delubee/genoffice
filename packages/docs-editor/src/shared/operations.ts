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
// C0 controls other than tab/newline never serialize.
// eslint-disable-next-line no-control-regex -- the pattern names exactly those controls.
const C0_CONTROLS = /[\x00-\x08\x0b\x0c\x0e-\x1f]/
const commentText = (max: number) =>
  z
    .string()
    .max(max)
    .refine((value) => !C0_CONTROLS.test(value))
const commentId = z.string().regex(/^[0-9]{1,9}$/)
export const commentsSchema = z
  .array(
    z
      .object({
        id: commentId,
        author: commentText(255),
        initials: commentText(20).optional(),
        date: z
          .string()
          .regex(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(:\d{2}(\.\d{1,3})?)?Z?$/)
          .optional(),
        text: commentText(65536),
        parentId: commentId.optional(),
        done: z.boolean().optional(),
        paraId: z
          .string()
          .regex(/^[0-9A-Fa-f]{8}$/)
          .optional(),
      })
      .strict(),
  )
  .max(5000)
  .refine(
    (comments) => new Set(comments.map((comment) => comment.id)).size === comments.length,
    'Duplicate comment ids',
  )

/**
 * Header/footer content regenerates its part from paragraphs and runs: values
 * are plain JSON whose text is escaped by the writer; nothing may carry raw XML.
 */
/** Run fields the writer emits verbatim; the session admits only values taken from the file. */
export const HF_RAW_KEYS = new Set(['rawRPr', 'fldBeginXml', 'sdtCheckboxXml', 'xml'])
const hfValue = (key: string, value: unknown, depth: number): boolean => {
  if (HF_RAW_KEYS.has(key)) return value == null || typeof value === 'string'
  if (value == null || typeof value === 'boolean') return true
  if (typeof value === 'number') return Number.isFinite(value) && Math.abs(value) <= 10_000_000
  if (typeof value === 'string')
    return (
      value.length <= 65536 &&
      // eslint-disable-next-line no-control-regex -- C0 controls other than tab/newline never serialize.
      !/[\x00-\x08\x0b\x0c\x0e-\x1f]/.test(value) &&
      (key === 'text' || !/[<>]/.test(value))
    )
  if (depth >= 8) return false
  if (Array.isArray(value))
    return value.length <= 2000 && value.every((item) => hfValue(key, item, depth + 1))
  if (typeof value === 'object')
    return Object.entries(value).every(
      ([name, item]) => /^[A-Za-z0-9_]{1,60}$/.test(name) && hfValue(name, item, depth + 1),
    )
  return false
}
/** Word's iterated SHA-512 password record (hash only; the password never crosses). */
const base64 = z.string().regex(/^[A-Za-z0-9+/=]{4,200}$/)
const passwordHash = {
  hash: base64.optional(),
  salt: base64.optional(),
  spinCount: z.number().int().min(1).max(10_000_000).optional(),
  algorithmSid: z.literal(14).optional(),
}
/** Theme typefaces and names: any script, never markup. */
const themeText = z
  .string()
  .max(100)
  .refine((value) => !/[<>&"]/.test(value) && !C0_CONTROLS.test(value))
/** Styles pane definitions: identifiers plus plain run/paragraph properties. */
const styleIdPattern = /^[A-Za-z0-9_.-]{1,253}$/
export const styleUpsertsSchema = z
  .array(
    z
      .object({
        styleId: z.string().regex(styleIdPattern),
        type: z.enum(['paragraph', 'character', 'table']).optional(),
        name: z
          .string()
          .max(253)
          .refine((value) => !/[<>&"]/.test(value))
          .optional(),
        basedOn: z.string().regex(styleIdPattern).nullable().optional(),
        next: z.string().regex(styleIdPattern).nullable().optional(),
        quickFormat: z.boolean().optional(),
        rPr: z.record(z.string(), z.unknown()).optional(),
        pPr: z.record(z.string(), z.unknown()).optional(),
      })
      .strict()
      .refine(
        (style) => settingValue(style.rPr ?? null, 0) && settingValue(style.pPr ?? null, 0),
        'Invalid style properties',
      ),
  )
  .max(500)

/** Footnote / endnote bodies: plain JSON runs whose text the writer escapes. */
export const notesSchema = z
  .array(
    z
      .object({ id: z.string().regex(/^-?[0-9]{1,9}$/), text: z.string().max(65536) })
      .catchall(z.unknown())
      .refine((note) => hfValue('note', note, 0), 'Invalid note'),
  )
  .max(5000)
  .refine(
    (notes) => new Set(notes.map((note) => note.id)).size === notes.length,
    'Duplicate note ids',
  )
const headerFooterSchema = z
  .object({
    text: z.string().max(65536),
    pageNumber: z.boolean().optional(),
    paras: z.array(z.record(z.string(), z.unknown())).max(200).optional(),
    newPictures: z
      .record(
        z.string().regex(/^[A-Za-z0-9_-]{1,80}$/),
        z
          .object({
            base64: z.string().regex(/^[A-Za-z0-9+/=]+$/),
            mime: z.enum(['image/png', 'image/jpeg', 'image/gif']),
          })
          .strict(),
      )
      .optional(),
  })
  .strict()
  .refine(
    (value) => hfValue('header', { ...value, newPictures: undefined }, 0),
    'Invalid header or footer',
  )
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
    header: headerFooterSchema.optional(),
    comments: commentsSchema.optional(),
    footnotes: notesSchema.optional(),
    styles: styleUpsertsSchema.optional(),
    protection: z
      .object({
        edit: z.enum(['readOnly', 'comments', 'trackedChanges', 'forms', 'none']),
        enforced: z.boolean(),
        ...passwordHash,
      })
      .strict()
      .nullable()
      .optional(),
    writeProtection: z
      .object({ recommended: z.boolean().optional(), ...passwordHash })
      .strict()
      .nullable()
      .optional(),
    removePersonalInfo: z.boolean().optional(),
    themeFonts: z
      .record(z.string().regex(/^[A-Za-z]{1,20}$/), themeText)
      .refine((fonts) => typeof fonts.major === 'string' && typeof fonts.minor === 'string')
      .optional(),
    themeColors: z
      .record(
        z.string().regex(/^[A-Za-z0-9]{1,20}$/),
        z.union([z.string().regex(/^[0-9A-Fa-f]{6}$/), themeText]),
      )
      .optional(),
    endnotes: notesSchema.optional(),
    footer: headerFooterSchema.optional(),
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
