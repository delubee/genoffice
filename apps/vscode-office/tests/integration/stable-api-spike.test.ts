import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import { suite, test } from 'mocha'

suite('Stable API proof', () => {
  test('real custom document edit/save/undo/redo and native tool registration', async () => {
    const extension = vscode.extensions.getExtension('genoffice-test.workbook-api-probe')!
    const probe = await extension.activate()
    const uri = vscode.Uri.joinPath(
      vscode.workspace.workspaceFolders![0]!.uri,
      'stable-api.testbook',
    )
    const disk = async () => new TextDecoder().decode(await vscode.workspace.fs.readFile(uri))
    try {
      await vscode.workspace.fs.writeFile(uri, new TextEncoder().encode('1'))
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice-test.probe')
      assert.equal(probe.inspect().value, 1)
      assert.ok(vscode.lm.tools.some((tool) => tool.name === 'genoffice_test_probe'))
      await probe.prepare({ expected: 1, value: 2 })
      assert.equal(probe.inspect().value, 1, 'prepare must not edit')
      await vscode.lm.invokeTool('genoffice_test_probe', {
        input: { expected: 1, value: 2 },
        toolInvocationToken: undefined,
      })
      assert.equal(probe.inspect().value, 2)
      await vscode.commands.executeCommand('workbench.action.files.save')
      assert.equal(await disk(), '2')
      await vscode.commands.executeCommand('undo')
      assert.equal(probe.inspect().value, 1, 'undo survives save')
      await vscode.commands.executeCommand('redo')
      assert.equal(probe.inspect().value, 2)
      await probe.prepare({ expected: 2, value: 3 })
      probe.edit(4)
      await assert.rejects(
        Promise.resolve(
          vscode.lm.invokeTool('genoffice_test_probe', {
            input: { expected: 2, value: 3 },
            toolInvocationToken: undefined,
          }),
        ),
      )
      assert.equal(probe.inspect().value, 4)
      await vscode.commands.executeCommand('workbench.action.files.save')
      assert.equal(await disk(), '4')
      assert.equal(probe.inspect().undos, 1)
      assert.equal(probe.inspect().redos, 1)
      assert.equal(probe.inspect().edits, 2, 'history movement must not fire new edit events')
    } finally {
      await vscode.commands.executeCommand('workbench.action.closeAllEditors')
      await vscode.workspace.fs.delete(uri, { useTrash: false })
    }
  })
})
