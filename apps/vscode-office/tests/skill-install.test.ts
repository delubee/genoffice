import { expect, it } from 'vitest'
import { mkdir, readFile, writeFile, symlink, stat } from 'node:fs/promises'
import { join } from 'node:path'
import { inspectSkillInstallation, installWorkbookSkill } from '../src/extension/skill-installation'
import { withTestWorkspace } from './helpers'
import { workbookToolAvailability } from '../src/extension/skill-status'

it('reports missing registered native tools without claiming current-chat enablement or changing global model settings', () => {
  const status = workbookToolAvailability([])
  expect(status.registered).toBe(false)
  expect(status.missing).toContain('genoffice_workbook_applyImport')
  expect(status.hint).toContain('tool picker')
  expect(status.currentChatEnablement).toBe('not-exposed-by-stable-api')
})

it('rejects a workspace skill path redirected outside its selected root without creating outside directories', async () => {
  await withTestWorkspace(async (directory) => {
    const root = join(directory, 'root'),
      outside = join(directory, 'outside'),
      bundle = join(directory, 'bundle')
    await Promise.all([mkdir(root), mkdir(outside), mkdir(bundle)])
    await writeFile(join(bundle, 'SKILL.md'), 'Source instructions')
    await symlink(outside, join(root, '.github'), process.platform === 'win32' ? 'junction' : 'dir')
    await expect(
      inspectSkillInstallation({ workspaceRoot: root, bundleDirectory: bundle, version: '1' }),
    ).rejects.toMatchObject({ code: 'skill_path_not_allowed' })
    await expect(stat(join(outside, 'skills'))).rejects.toMatchObject({ code: 'ENOENT' })
  })
})

it('installs only into the selected workspace and detects user changes without modifying files during inspection or cancellation', async () => {
  await withTestWorkspace(async (directory) => {
    const bundle = join(directory, 'bundle'),
      a = join(directory, 'a'),
      b = join(directory, 'b')
    await Promise.all([mkdir(bundle), mkdir(a), mkdir(b)])
    await writeFile(
      join(bundle, 'SKILL.md'),
      '---\nname: genoffice-workbook\ndescription: Test skill\n---\nOriginal instructions\n',
    )
    const initial = await inspectSkillInstallation({
      workspaceRoot: a,
      bundleDirectory: bundle,
      version: '1.0.0',
    })
    expect(initial).toMatchObject({ exists: false, userModified: false })
    const installed = await installWorkbookSkill({ ...initial, decision: 'install' })
    expect(await readFile(join(installed.targetDirectory, 'SKILL.md'), 'utf8')).toContain(
      'Original instructions',
    )
    expect(
      (
        await inspectSkillInstallation({
          workspaceRoot: b,
          bundleDirectory: bundle,
          version: '1.0.0',
        })
      ).exists,
    ).toBe(false)
    await writeFile(join(installed.targetDirectory, 'SKILL.md'), 'User instructions\n')
    const modified = await inspectSkillInstallation({
      workspaceRoot: a,
      bundleDirectory: bundle,
      version: '1.0.0',
    })
    expect(modified).toMatchObject({ exists: true, userModified: true })
    expect(await installWorkbookSkill({ ...modified, decision: 'keep' })).toMatchObject({
      status: 'kept',
    })
    expect(await readFile(join(installed.targetDirectory, 'SKILL.md'), 'utf8')).toBe(
      'User instructions\n',
    )
    expect(
      (
        await inspectSkillInstallation({
          workspaceRoot: a,
          bundleDirectory: bundle,
          version: '1.0.0',
        })
      ).userModified,
    ).toBe(true)
  })
})

it('backs up an approved update, preserves extra user files and refuses a target changed since the preview', async () => {
  await withTestWorkspace(async (directory) => {
    const bundle = join(directory, 'bundle'),
      root = join(directory, 'workspace')
    await Promise.all([mkdir(bundle), mkdir(root)])
    await writeFile(join(bundle, 'SKILL.md'), 'Version one\n')
    const first = await installWorkbookSkill({
      ...(await inspectSkillInstallation({
        workspaceRoot: root,
        bundleDirectory: bundle,
        version: '1',
      })),
      decision: 'install',
    })
    await writeFile(join(first.targetDirectory, 'notes.txt'), 'Keep my notes\n')
    await writeFile(join(first.targetDirectory, 'SKILL.md'), 'My edited instructions\n')
    await writeFile(join(bundle, 'SKILL.md'), 'Version two\n')
    const preview = await inspectSkillInstallation({
      workspaceRoot: root,
      bundleDirectory: bundle,
      version: '2',
    })
    expect(preview.userModified).toBe(true)
    expect(preview.differences).toContainEqual(
      expect.objectContaining({ path: 'notes.txt', kind: 'preserved' }),
    )
    await writeFile(join(first.targetDirectory, 'SKILL.md'), 'Changed after preview\n')
    await expect(installWorkbookSkill({ ...preview, decision: 'update' })).rejects.toMatchObject({
      code: 'skill_changed',
    })
    expect(await readFile(join(first.targetDirectory, 'SKILL.md'), 'utf8')).toBe(
      'Changed after preview\n',
    )
    const result = await installWorkbookSkill({
      ...(await inspectSkillInstallation({
        workspaceRoot: root,
        bundleDirectory: bundle,
        version: '2',
      })),
      decision: 'update',
    })
    expect(await readFile(join(result.targetDirectory, 'SKILL.md'), 'utf8')).toBe('Version two\n')
    expect(await readFile(join(result.targetDirectory, 'notes.txt'), 'utf8')).toBe(
      'Keep my notes\n',
    )
    expect(await readFile(join(result.backupDirectory!, 'SKILL.md'), 'utf8')).toBe(
      'Changed after preview\n',
    )
    expect(
      (
        await inspectSkillInstallation({
          workspaceRoot: root,
          bundleDirectory: bundle,
          version: '2',
        })
      ).userModified,
    ).toBe(false)
  })
})
