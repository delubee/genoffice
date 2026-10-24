import * as vscode from 'vscode'
import { join } from 'node:path'
import { WorkbookError } from '@genoffice/workbook-session/domain'
import { installWorkbookSkill, type SkillPreview } from '../skill-installation'
import { workbookSkillStatus } from '../skill-status'

export function registerWorkbookSkillCommands(context: vscode.ExtensionContext) {
  const bundleDirectory = context.asAbsolutePath('skills/genoffice-workbook')
  const root = async (requested?: vscode.Uri) => {
    if (!vscode.workspace.isTrusted)
      throw new WorkbookError(
        'untrusted_workspace',
        'Trust this workspace before installing a skill.',
      )
    const selected = requested
      ? vscode.workspace.workspaceFolders?.find(
          (folder) => folder.uri.toString() === requested.toString(),
        )
      : await vscode.window.showWorkspaceFolderPick({
          placeHolder: 'Select a workspace for .github/skills/genoffice-workbook',
        })
    if (!selected || selected.uri.scheme !== 'file') {
      if (requested)
        throw new WorkbookError(
          'path_not_allowed',
          'Select a current workspace folder on the extension host.',
        )
      return undefined
    }
    return selected
  }
  const inspect = (workspaceRoot: string) =>
    workbookSkillStatus({
      workspaceRoot,
      bundleDirectory,
      registeredToolNames: vscode.lm.tools.map((tool) => tool.name),
    })
  const showDifferences = async (preview: SkillPreview) => {
    const selected = await vscode.window.showQuickPick(
      preview.differences
        .filter((file) => file.kind !== 'same')
        .map((file) => ({ label: file.path, description: file.kind, file })),
      { title: 'GenOffice skill differences; custom files are preserved' },
    )
    if (!selected) return
    const source = vscode.Uri.file(join(preview.bundleDirectory, ...selected.file.path.split('/'))),
      target = vscode.Uri.file(join(preview.targetDirectory, ...selected.file.path.split('/')))
    if (selected.file.kind === 'changed')
      await vscode.commands.executeCommand(
        'vscode.diff',
        target,
        source,
        `Installed ↔ GenOffice ${preview.version}: ${selected.file.path}`,
      )
    else
      await vscode.commands.executeCommand(
        'vscode.open',
        selected.file.kind === 'added' ? source : target,
      )
  }
  return vscode.Disposable.from(
    vscode.commands.registerCommand(
      'genoffice.installWorkbookSkill',
      async (requested?: vscode.Uri) => {
        const selected = await root(requested)
        if (!selected) return
        let preview = await inspect(selected.uri.fsPath)
        let decision: 'install' | 'update' = 'install'
        if (preview.exists) {
          if (!preview.updateAvailable && !preview.userModified) {
            void vscode.window.showInformationMessage(
              `GenOffice skill ${preview.version} is current: ${preview.targetDirectory}`,
            )
            return { status: 'current', targetDirectory: preview.targetDirectory }
          }
          for (;;) {
            const choice = await vscode.window.showWarningMessage(
              `GenOffice skill ${preview.installedVersion ?? 'untracked'} → ${preview.version}${preview.userModified ? '; local changes detected' : ''}. Target: ${preview.targetDirectory}. Updating retains a previous copy under .genoffice/skill-backups and preserves extra files.`,
              { modal: true },
              'Keep Existing',
              'View Differences',
              'Update with Backup',
            )
            if (!choice || choice === 'Keep Existing')
              return installWorkbookSkill({ ...preview, decision: 'keep' })
            if (choice === 'View Differences') {
              await showDifferences(preview)
              preview = await inspect(selected.uri.fsPath)
              continue
            }
            decision = 'update'
            break
          }
        }
        const result = await installWorkbookSkill({ ...preview, decision })
        const enabled = vscode.workspace
          .getConfiguration('chat', selected.uri)
          .get<boolean>('useAgentSkills')
        void vscode.window
          .showInformationMessage(
            `GenOffice skill ${preview.version} installed: ${result.targetDirectory}${result.backupDirectory ? `; previous copy: ${result.backupDirectory}` : ''}.${!preview.tools.registered ? ` ${preview.tools.hint}` : ''}${enabled === false ? ' Agent skills are disabled in VS Code settings.' : ' Enable GenOffice tools in your native agent tool picker.'}`,
            'Open Skill',
            ...(enabled === false ? ['Open Skill Setting'] : []),
          )
          .then(async (action) => {
            if (action === 'Open Skill')
              await vscode.commands.executeCommand(
                'vscode.open',
                vscode.Uri.file(join(result.targetDirectory, 'SKILL.md')),
              )
            else if (action === 'Open Skill Setting')
              await vscode.commands.executeCommand(
                'workbench.action.openSettings',
                'chat.useAgentSkills',
              )
          })
        return result
      },
    ),
    vscode.commands.registerCommand(
      'genoffice.workbookSkillStatus',
      async (requested?: vscode.Uri) => {
        const selected = await root(requested)
        if (!selected) return
        const status = await inspect(selected.uri.fsPath)
        return {
          ...status,
          agentSkillsSetting:
            vscode.workspace
              .getConfiguration('chat', selected.uri)
              .get<boolean>('useAgentSkills') ?? null,
        }
      },
    ),
  )
}
