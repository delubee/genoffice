import { readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { WorkbookError } from '@genoffice/workbook-session/domain'
import { workbookToolDefinitions } from './tools/schemas'
import { inspectSkillInstallation } from './skill-installation'

export function workbookToolAvailability(registeredNames: readonly string[]) {
  const missing = workbookToolDefinitions
    .map((tool) => tool.name)
    .filter((name) => !registeredNames.includes(name))
  return {
    registered: missing.length === 0,
    missing,
    currentChatEnablement: 'not-exposed-by-stable-api',
    hint: 'Enable GenOffice tools in the native agent tool picker. Use an agent that supports these tools; manual workbook editing is independent.',
  }
}
export async function workbookSkillStatus(input: {
  workspaceRoot: string
  bundleDirectory: string
  registeredToolNames: readonly string[]
}) {
  const source = await readFile(join(input.bundleDirectory, 'SKILL.md'), 'utf8')
  const version =
    /^ {2}version:\s*(['"]?)([0-9]+\.[0-9]+\.[0-9]+(?:[-+][A-Za-z0-9.-]+)?)\1\s*$/m.exec(
      source,
    )?.[2]
  if (!version)
    throw new WorkbookError(
      'skill_bundle_invalid',
      'The packaged skill has no valid version metadata.',
    )
  const installation = await inspectSkillInstallation({ ...input, version })
  return { ...installation, tools: workbookToolAvailability(input.registeredToolNames) }
}
