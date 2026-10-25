import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import { suite, test } from 'mocha'
import JSZip from 'jszip'
import { buildDocx, TABLE_XML } from '../../../../packages/docx-engine/tests/helpers/build-docx'
import { execFile } from 'node:child_process'
import { promisify } from 'node:util'
import { join } from 'node:path'
import { randomBytes } from 'node:crypto'
import { mkdir, writeFile } from 'node:fs/promises'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'

suite('Product DOCX editor', () => {
  test('opens in the existing plugin, edits through native tools and saves with shared native undo', async () => {
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as {
      docxBackupForTests(uri: vscode.Uri): Promise<vscode.CustomDocumentBackup>
      docxRestoreForTests(uri: vscode.Uri, backupId: string): Promise<void>
    }
    const root = vscode.workspace.workspaceFolders![0]!.uri
    const uri = vscode.Uri.joinPath(root, 'native-docx.docx')
    const inputPath = vscode.Uri.joinPath(root, 'native-docx-input.json'),
      outputPath = vscode.Uri.joinPath(root, 'native-docx-result.json')
    const editedText = 'Edited through native chat tools'.padEnd(100, 'A')
    const invoke = async (kind: string, input: object) => {
      const result = await vscode.lm.invokeTool(`genoffice_docx_${kind}`, {
        input,
        toolInvocationToken: undefined,
      })
      const part = result.content.find((item) => item instanceof vscode.LanguageModelTextPart)
      assert.ok(part instanceof vscode.LanguageModelTextPart)
      return JSON.parse(part.value)
    }
    try {
      const body =
        '<w:p><w:r><w:t>Original</w:t></w:r></w:p>' +
        Array.from(
          { length: 199 },
          (_, index) => `<w:p><w:r><w:t>PARAGRAPH ${index}</w:t></w:r></w:p>`,
        ).join('') +
        TABLE_XML.repeat(10)
      const fixtureZip = await JSZip.loadAsync(await buildDocx({ bodyXml: body }))
      fixtureZip.file(
        'customXml/performance.xml',
        `<payload>${randomBytes(1100000).toString('base64')}</payload>`,
      )
      const fixture = await fixtureZip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
      assert.ok(fixture.length >= 1024 ** 2)
      await vscode.workspace.fs.writeFile(uri, fixture)
      const openedAt = Date.now()
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.docx')
      let context:
        { bindingId: string; descriptor: { revision: number; sessionId: string } } | undefined
      const deadline = Date.now() + 15000
      while (!context && Date.now() < deadline) {
        try {
          context = await invoke('context', { documentUri: uri.toString() })
        } catch (error) {
          if (!/busy/.test(String(error))) throw error
          await new Promise((resolve) => setTimeout(resolve, 100))
        }
      }
      assert.ok(context, 'DOCX webview becomes ready')
      const coldOpenMs = Date.now() - openedAt
      assert.ok(coldOpenMs <= 10000, `1 MiB / 200 paragraphs / 10 tables open in ${coldOpenMs} ms`)
      const results = join(
        vscode.extensions.getExtension('genoffice.genoffice-workbook')!.extensionPath,
        'test-results',
      )
      await mkdir(results, { recursive: true })
      const before = await invoke('read', { bindingId: context.bindingId, expectedRevision: 0 })
      assert.equal(before.blocks[0].text, 'Original')
      const input = {
        bindingId: context.bindingId,
        expectedRevision: 0,
        transactionId: 'native-docx-edit',
        summary: 'Replace fixture paragraph',
        operations: [
          {
            type: 'replaceText',
            blockId: before.blocks[0].blockId,
            start: 0,
            end: 8,
            text: editedText,
          },
        ],
      }
      const appliedAt = Date.now()
      const edited = await invoke('apply', input)
      const toolEdit100Ms = Date.now() - appliedAt
      assert.ok(toolEdit100Ms <= 1000, `100-character tool edit projected in ${toolEdit100Ms} ms`)
      await writeFile(
        join(results, `docx-${process.platform}-${process.arch}-performance.json`),
        JSON.stringify(
          {
            fixtureBytes: fixture.length,
            paragraphs: 200,
            tables: 10,
            coldOpenMs,
            toolEdit100Ms,
            vscode: vscode.version,
          },
          null,
          2,
        ),
      )
      assert.equal(edited.status, 'applied')
      assert.equal(
        (
          await invoke('operation_status', {
            bindingId: context.bindingId,
            transactionId: input.transactionId,
          })
        ).status,
        'applied',
      )
      await vscode.commands.executeCommand('workbench.action.files.save')
      const saved = await JSZip.loadAsync(await vscode.workspace.fs.readFile(uri))
      assert.match(
        await saved.file('word/document.xml')!.async('string'),
        /Edited through native chat tools/,
      )
      await vscode.commands.executeCommand('undo')
      const reverted = await invoke('read', { bindingId: context.bindingId, expectedRevision: 2 })
      assert.equal(reverted.blocks[0].text, 'Original')
      await vscode.commands.executeCommand('redo')
      assert.equal(
        (await invoke('read', { bindingId: context.bindingId, expectedRevision: 3 })).blocks[0]
          .text,
        editedText,
      )
      const exported = await invoke('export', {
        bindingId: context.bindingId,
        expectedRevision: 3,
        path: inputPath.fsPath,
      })
      await promisify(execFile)(
        process.platform === 'win32' ? 'python' : 'python3',
        [
          join(
            vscode.extensions.getExtension('genoffice.genoffice-workbook')!.extensionPath,
            'skills/genoffice-workbook/scripts/docx_transform.py',
          ),
          inputPath.fsPath,
          outputPath.fsPath,
          '--uppercase',
        ],
        { timeout: 10000 },
      )
      const plan = await invoke('prepare_import', {
        bindingId: context.bindingId,
        snapshotId: exported.snapshotId,
        expectedRevision: 3,
        path: outputPath.fsPath,
      })
      assert.equal(plan.changedBlocks, 1)
      await invoke('apply_import', {
        bindingId: context.bindingId,
        expectedRevision: 3,
        importId: plan.importId,
        sha256: plan.sha256,
        transactionId: 'native-docx-python',
      })
      assert.equal(
        (await invoke('read', { bindingId: context.bindingId, expectedRevision: 4 })).blocks[0]
          .text,
        editedText.toUpperCase(),
      )
      await vscode.commands.executeCommand('undo')
      assert.equal(
        (await invoke('read', { bindingId: context.bindingId, expectedRevision: 5 })).blocks[0]
          .text,
        editedText,
      )
      const unsaved = 'Unsaved recovery'
      await invoke('apply', {
        bindingId: context.bindingId,
        expectedRevision: 5,
        transactionId: 'before-backup',
        summary: 'Create an unsaved recovery fixture',
        operations: [
          {
            type: 'replaceText',
            blockId: before.blocks[0].blockId,
            start: 0,
            end: editedText.length,
            text: unsaved,
          },
        ],
      })
      const backup = await api.docxBackupForTests(uri)
      try {
        await invoke('apply', {
          bindingId: context.bindingId,
          expectedRevision: 6,
          transactionId: 'after-backup',
          summary: 'Change the live fixture after backup',
          operations: [
            {
              type: 'replaceText',
              blockId: before.blocks[0].blockId,
              start: 0,
              end: unsaved.length,
              text: 'Later edits',
            },
          ],
        })
        await api.docxRestoreForTests(uri, backup.id)
        const restoredContext = await invoke('context', { documentUri: uri.toString() })
        assert.notEqual(restoredContext.descriptor.sessionId, context.descriptor.sessionId)
        assert.equal(
          (await invoke('read', { bindingId: restoredContext.bindingId, expectedRevision: 0 }))
            .blocks[0].text,
          unsaved,
        )
      } finally {
        await backup.delete()
      }
    } finally {
      await vscode.commands.executeCommand('workbench.action.revertAndCloseActiveEditor')
      await vscode.workspace.fs.delete(uri).then(
        () => {},
        () => {},
      )
      for (const path of [inputPath, outputPath])
        await vscode.workspace.fs.delete(path).then(
          () => {},
          () => {},
        )
    }
  })
  test('keeps XLSX and DOCX independent while editing a hidden, explicitly bound Word document', async () => {
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as { getDocument(uri: vscode.Uri): WorkbookDocument }
    const root = vscode.workspace.workspaceFolders![0]!.uri
    const word = vscode.Uri.joinPath(root, 'coexist.docx'),
      sheet = vscode.Uri.joinPath(root, 'coexist.xlsx')
    const invoke = async (kind: string, input: object) => {
      const result = await vscode.lm.invokeTool(`genoffice_docx_${kind}`, {
        input,
        toolInvocationToken: undefined,
      })
      const part = result.content.find((value) => value instanceof vscode.LanguageModelTextPart)
      assert.ok(part instanceof vscode.LanguageModelTextPart)
      return JSON.parse(part.value)
    }
    try {
      await vscode.workspace.fs.writeFile(
        word,
        await buildDocx({ bodyXml: '<w:p><w:r><w:t>Word</w:t></w:r></w:p>' }),
      )
      await vscode.workspace.fs.writeFile(sheet, await manualWorkbook())
      await vscode.commands.executeCommand('vscode.openWith', word, 'genoffice.docx')
      const context = await invoke('context', { documentUri: word.toString() })
      const read = await invoke('read', { bindingId: context.bindingId, expectedRevision: 0 })
      await vscode.commands.executeCommand('vscode.openWith', sheet, 'genoffice.workbook')
      const workbook = api.getDocument(sheet)
      await workbook.whenViewReady()
      const workbookState = workbook.engine.state.stateId
      await invoke('apply', {
        bindingId: context.bindingId,
        expectedRevision: 0,
        transactionId: 'hidden-word',
        summary: 'Edit the explicitly bound hidden Word document',
        operations: [
          {
            type: 'replaceText',
            blockId: read.blocks[0].blockId,
            start: 0,
            end: 4,
            text: 'Hidden Word edit',
          },
        ],
      })
      assert.equal(workbook.engine.state.stateId, workbookState)
      assert.equal(
        (await invoke('read', { bindingId: context.bindingId, expectedRevision: 1 })).blocks[0]
          .text,
        'Hidden Word edit',
      )
      await vscode.commands.executeCommand('vscode.openWith', word, 'genoffice.docx')
      await vscode.commands.executeCommand('undo')
      assert.equal(
        (await invoke('read', { bindingId: context.bindingId, expectedRevision: 2 })).blocks[0]
          .text,
        'Word',
      )
      assert.equal(workbook.engine.state.stateId, workbookState)
    } finally {
      for (const [uri, editor] of [
        [word, 'genoffice.docx'],
        [sheet, 'genoffice.workbook'],
      ] as const) {
        await vscode.commands.executeCommand('vscode.openWith', uri, editor)
        await vscode.commands.executeCommand('workbench.action.revertAndCloseActiveEditor')
        await vscode.workspace.fs.delete(uri).then(
          () => {},
          () => {},
        )
      }
    }
  })
})
