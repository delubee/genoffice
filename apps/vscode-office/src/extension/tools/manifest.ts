import { z } from 'zod'
import { workbookToolDefinitions, toolSchemas } from './schemas'
import { docxToolDefinitions, docxToolSchemas } from '../docx-schemas'

export const toolContributions = [
  ...workbookToolDefinitions.map((tool) => ({
    name: tool.name,
    displayName: tool.displayName,
    modelDescription: tool.description,
    canBeReferencedInPrompt: true,
    toolReferenceName: tool.name,
    inputSchema: z.toJSONSchema(toolSchemas[tool.kind], {
      target: 'draft-7',
      unrepresentable: 'any',
    }),
  })),
  ...docxToolDefinitions.map((tool) => ({
    name: tool.name,
    displayName: tool.displayName,
    modelDescription: tool.description,
    canBeReferencedInPrompt: true,
    toolReferenceName: tool.name,
    inputSchema: z.toJSONSchema(docxToolSchemas[tool.kind], {
      target: 'draft-7',
      unrepresentable: 'any',
    }),
  })),
]
