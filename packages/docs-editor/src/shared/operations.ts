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
