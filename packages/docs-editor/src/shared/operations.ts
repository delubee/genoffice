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
