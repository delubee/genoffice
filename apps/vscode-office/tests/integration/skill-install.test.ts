import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import { suite, test } from 'mocha'
import { mkdir, readFile, writeFile } from 'node:fs/promises'
import { basename, join } from 'node:path'

suite('Product workspace skill commands', () => {
  test('installs the packaged skill in the explicit workspace and reports local modifications', async function () {
    const extension = vscode.extensions.getExtension('genoffice.genoffice-workbook')!
    await extension.activate()
    const root = vscode.workspace.workspaceFolders![0]!.uri
    assert.equal(
      basename(root.fsPath),
      'test-workspace',
      'this test only owns the isolated fixture workspace',
    )
    const inspect = () => vscode.commands.executeCommand<any>('genoffice.workbookSkillStatus', root)
    const before = await inspect()
    if (before.exists) {
      this.skip()
      return
    }
    let controlledDigest: string | undefined
    try {
      const result = await vscode.commands.executeCommand<{
        status: string
        targetDirectory: string
      }>('genoffice.installWorkbookSkill', root)
      assert.equal(result!.status, 'installed')
      assert.equal(
        result!.targetDirectory,
        join(root.fsPath, '.github', 'skills', 'genoffice-workbook'),
      )
      const status = await inspect()
      controlledDigest = status.currentDigest
      assert.equal(status.version, '0.1.0')
      assert.equal(status.userModified, false)
      assert.equal(status.tools.registered, true)
      assert.ok(
        (await readFile(join(result!.targetDirectory, 'SKILL.md'), 'utf8')).includes(
          'genoffice_workbook_',
        ),
      )
      assert.ok(
        (await readFile(join(result!.targetDirectory, 'scripts', 'analyze.py'), 'utf8')).includes(
          'Decimal',
        ),
      )
      await writeFile(
        join(result!.targetDirectory, 'SKILL.md'),
        'User changes made in the isolated fixture workspace\n',
      )
      const modified = await inspect()
      controlledDigest = modified.currentDigest
      assert.equal(modified.userModified, true)
      await assert.rejects(
        Promise.resolve(
          vscode.commands.executeCommand(
            'genoffice.installWorkbookSkill',
            vscode.Uri.joinPath(root, 'not-a-workspace'),
          ),
        ),
      )
      await mkdir(join(extension.extensionPath, 'test-results'), { recursive: true })
      await writeFile(
        join(extension.extensionPath, 'test-results/skill-native-proof.json'),
        JSON.stringify(
          {
            passed: true,
            version: vscode.version,
            skillVersion: status.version,
            workspacePath: root.fsPath,
            installedPath: result!.targetDirectory,
            toolsRegistered: true,
            userChangesDetected: true,
            outsideWorkspaceRejected: true,
            modelDiscovery: 'not verified',
            confirmationRejectionInRealChat: 'not verified',
            checkedAt: new Date().toISOString(),
          },
          null,
          2,
        ),
      )
    } finally {
      const current = await inspect()
      if (controlledDigest && current.currentDigest === controlledDigest)
        await vscode.workspace.fs.delete(vscode.Uri.file(current.targetDirectory), {
          recursive: true,
          useTrash: false,
        })
    }
  })
})
