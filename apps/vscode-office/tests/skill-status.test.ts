import { expect, it } from 'vitest'
import { mkdtemp, mkdir, writeFile, rm } from 'node:fs/promises'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { workbookSkillStatus } from '../src/extension/skill-status'

it.each(['"0.2.0"', "'0.2.0'", '0.2.0'])(
  'reads valid YAML version %s when inspecting a shipped skill',
  async (version) => {
    const root = await mkdtemp(join(tmpdir(), 'genoffice-skill-version-'))
    try {
      const bundleDirectory = join(root, 'bundle')
      await mkdir(bundleDirectory)
      await writeFile(
        join(bundleDirectory, 'SKILL.md'),
        `---\nname: genoffice-workbook\nmetadata:\n  version: ${version}\n---\n`,
      )
      const status = await workbookSkillStatus({
        workspaceRoot: root,
        bundleDirectory,
        registeredToolNames: [],
      })
      expect(status.version).toBe('0.2.0')
      expect(status.exists).toBe(false)
    } finally {
      await rm(root, { recursive: true, force: true })
    }
  },
)
