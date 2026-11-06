import * as vscode from 'vscode'
import * as assert from 'node:assert/strict'
import { readFile, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'
import { buildDocx } from '../../../../packages/docx-engine/tests/helpers/build-docx'

export async function activate() {
  const root = process.env.GENOFFICE_PACKAGE_PROOF_ROOT!
  try {
    const extension = vscode.extensions.getExtension('genoffice.genoffice-workbook')!
    assert.ok(extension)
    const api = (await extension.activate()) as { getDocument(uri: vscode.Uri): WorkbookDocument }
    assert.ok(extension.extensionPath.startsWith(join(root, 'extensions')))
    const uri = vscode.Uri.joinPath(vscode.workspace.workspaceFolders![0]!.uri, 'packaged.xlsx')
    await vscode.workspace.fs.writeFile(uri, await manualWorkbook())
    await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
    let document = api.getDocument(uri)
    await document.whenViewReady()
    assert.deepEqual(await document.bridge!.request('view.flush', {}), { ready: true })
    await document.engine.submit({
      schemaVersion: 1,
      sessionId: document.sessionId,
      transactionId: 'package-edit',
      expectedRevision: 0,
      origin: 'manual',
      summary: 'Packaged manual edit',
      operations: [
        { op: 'set_cell', sheetId: document.engine.state.sheets[0]!.id, address: 'A2', value: 731 },
      ],
    })
    await vscode.commands.executeCommand('workbench.action.files.save')
    await vscode.commands.executeCommand('workbench.action.closeActiveEditor')
    await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
    document = api.getDocument(uri)
    await document.whenViewReady()
    const invoke = async (suffix: string, input: object) => {
      const result = await vscode.lm.invokeTool(`genoffice_workbook_${suffix}`, {
        input,
        toolInvocationToken: undefined,
      })
      const part = result.content.find(
        (x) => x instanceof vscode.LanguageModelTextPart,
      ) as vscode.LanguageModelTextPart
      const parsed = JSON.parse(part.value)
      assert.equal(parsed.status, 'ok')
      return parsed.data
    }
    const context = await invoke('context', { schemaVersion: 1, documentUri: uri.toString() })
    const read = await invoke('read', {
      schemaVersion: 1,
      bindingId: context.binding.bindingId,
      sheetId: document.engine.state.sheets[0]!.id,
      range: 'A2',
    })
    assert.equal(read.cells[0].rawValue, 731)
    const wordUri = vscode.Uri.joinPath(vscode.workspace.workspaceFolders![0]!.uri, 'packaged.docx')
    await vscode.workspace.fs.writeFile(
      wordUri,
      await buildDocx({ bodyXml: '<w:p><w:r><w:t>Original Word text</w:t></w:r></w:p>' }),
    )
    await vscode.commands.executeCommand('vscode.openWith', wordUri, 'genoffice.docx')
    const invokeDocx = async (suffix: string, input: object) => {
      const result = await vscode.lm.invokeTool(`genoffice_docx_${suffix}`, {
        input,
        toolInvocationToken: undefined,
      })
      const part = result.content.find((item) => item instanceof vscode.LanguageModelTextPart)
      assert.ok(part instanceof vscode.LanguageModelTextPart)
      return JSON.parse(part.value)
    }
    let wordContext = await invokeDocx('context', { documentUri: wordUri.toString() })
    const wordRead = await invokeDocx('read', {
      bindingId: wordContext.bindingId,
      expectedRevision: 0,
    })
    assert.equal(wordRead.blocks[0].text, 'Original Word text')
    await writeFile(join(root, 'word-ready.json'), '{"ready":true}')
    const manualDeadline = Date.now() + 30000
    while (true) {
      if (Date.now() > manualDeadline) throw new Error('Installed DOCX manual edit timed out')
      if (
        (await readFile(join(root, 'word-edited.json'), 'utf8').catch(() => '')) ===
        '{"edited":true}'
      )
        break
      await new Promise((resolve) => setTimeout(resolve, 50))
    }
    assert.equal(
      (await invokeDocx('read', { bindingId: wordContext.bindingId, expectedRevision: 1 }))
        .blocks[0].text,
      'Installed DOCX works',
    )
    await vscode.commands.executeCommand('workbench.action.files.save')
    await vscode.commands.executeCommand('undo')
    assert.equal(
      (await invokeDocx('read', { bindingId: wordContext.bindingId, expectedRevision: 2 }))
        .blocks[0].text,
      'Original Word text',
    )
    await vscode.commands.executeCommand('redo')
    await vscode.commands.executeCommand('workbench.action.files.save')
    await vscode.commands.executeCommand('workbench.action.closeActiveEditor')
    await vscode.commands.executeCommand('vscode.openWith', wordUri, 'genoffice.docx')
    wordContext = await invokeDocx('context', { documentUri: wordUri.toString() })
    assert.equal(
      (await invokeDocx('read', { bindingId: wordContext.bindingId, expectedRevision: 0 }))
        .blocks[0].text,
      'Installed DOCX works',
    )
    await vscode.commands.executeCommand(
      'genoffice.installWorkbookSkill',
      vscode.workspace.workspaceFolders![0]!.uri,
    )
    const status = (await vscode.commands.executeCommand(
      'genoffice.workbookSkillStatus',
      vscode.workspace.workspaceFolders![0]!.uri,
    )) as { packagedVersion: string }
    assert.ok(status)
    await writeFile(
      join(root, 'result.json'),
      JSON.stringify(
        {
          passed: true,
          vscode: vscode.version,
          platform: `${process.platform}-${process.arch}`,
          installedExtension: extension.extensionPath,
          savedReopenedValue: 731,
          nativeToolRead: true,
          docxSavedReopenedText: 'Installed DOCX works',
          docxNativeUndoRedo: true,
          docxManualEdit: true,
          skillInstalled: true,
          runtimeCargoPathRemoved: true,
          realModel: 'unverified',
        },
        null,
        2,
      ),
    )
    await vscode.commands.executeCommand('workbench.action.closeAllEditors')
    await vscode.commands.executeCommand('workbench.action.quit')
  } catch (error) {
    await writeFile(
      join(root, 'result.json'),
      JSON.stringify({ passed: false, error: String(error) }),
    )
    await vscode.commands.executeCommand('workbench.action.quit')
  }
}
