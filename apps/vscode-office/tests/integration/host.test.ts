import { strict as assert } from 'node:assert'
import { suite, test } from 'mocha'
import * as vscode from 'vscode'

suite('Extension host harness', () => {
  test('activates the actual extension in a local test workspace', async () => {
    const extension = vscode.extensions.getExtension('genoffice.genoffice-workbook')
    assert.ok(extension, 'The development extension must be installed in the test host')
    await extension.activate()
    assert.equal(extension.isActive, true)
    assert.ok(vscode.workspace.workspaceFolders?.length)
    assert.equal(vscode.workspace.workspaceFolders?.[0]?.uri.scheme, 'file')
  })
})
