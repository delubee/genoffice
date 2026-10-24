import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import { readFile, writeFile, readdir, stat, readlink } from 'node:fs/promises'
import { basename, dirname, join } from 'node:path'
import { createHash } from 'node:crypto'
import { manualWorkbook } from '../fixtures/manual-workbook'
import type { WorkbookDocument } from '../../src/extension/workbook-document'

async function nativeBackupFiles(directory: string): Promise<string[]> {
  const entries = await readdir(directory, { withFileTypes: true }).catch(() => [])
  const result: string[] = []
  for (const entry of entries) {
    const path = join(directory, entry.name)
    if (entry.isDirectory()) result.push(...(await nativeBackupFiles(path)))
    else if (
      (await stat(path)).size < 1024 * 1024 &&
      (await readFile(path, 'utf8')).includes('backupId')
    )
      result.push(path)
  }
  return result
}

// Do not hold VS Code's activation barrier while waiting for its restoration phase.
export function activate() {
  void runProof()
}

async function runProof() {
  const root = process.env.GENOFFICE_RECOVERY_PROOF_ROOT
  if (!root || !basename(root).startsWith('normal-recovery-'))
    throw new Error('An isolated recovery proof root is required.')
  const workspace = vscode.workspace.workspaceFolders?.[0]?.uri
  assert.equal(workspace?.fsPath, join(root, 'workspace'))
  const uri = vscode.Uri.joinPath(workspace!, 'crash.xlsx'),
    marker = join(root, 'marker.json'),
    result = join(root, 'result.json'),
    phase = process.env.GENOFFICE_RECOVERY_PROOF_PHASE
  const killTestApplication = async () => {
    assert.equal(process.platform, 'linux')
    const executable = await readlink(`/proc/${process.ppid}/exe`)
    assert.ok(executable.startsWith(`${dirname(dirname(vscode.env.appRoot))}/`))
    assert.ok(vscode.env.appRoot.includes('/apps/vscode-office/.vscode-test/'))
    process.kill(process.ppid, 'SIGKILL')
  }
  try {
    const api = (await vscode.extensions
      .getExtension('genoffice.genoffice-workbook')!
      .activate()) as { getDocument(uri: vscode.Uri): WorkbookDocument | undefined }
    if (phase === 'seed') {
      await vscode.workspace.fs.writeFile(uri, await manualWorkbook())
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
    } else {
      const deadline = Date.now() + (phase === 'recover' ? 2_000 : 30_000)
      while (!api.getDocument(uri) && Date.now() < deadline)
        await new Promise((resolve) => setTimeout(resolve, 50))
      if (!api.getDocument(uri))
        if (phase === 'recover')
          await vscode.commands.executeCommand('genoffice.recoverWorkbook', uri)
        else await vscode.commands.executeCommand('workbench.action.logWorkingCopies')
    }
    const document = api.getDocument(uri)!
    assert.ok(document, 'VS Code must automatically reopen its native backup')
    await document.whenViewReady()
    await document.bridge!.request('view.flush', {})
    const hash = async () =>
      createHash('sha256')
        .update(await readFile(uri.fsPath))
        .digest('hex')
    if (phase === 'seed' || phase === 'crash') {
      const previous = phase === 'crash' ? JSON.parse(await readFile(marker, 'utf8')) : undefined
      if (previous) {
        assert.equal(document.engine.state.stateId, previous.stateId)
        assert.equal(document.engine.state.sheets[0]!.cells.A2!.rawValue, 99)
      }
      await document.engine.submit({
        schemaVersion: 1,
        sessionId: document.sessionId,
        transactionId: 'normal-crash-proof',
        expectedRevision: 0,
        origin: 'manual',
        summary: 'Unsaved crash fixture',
        operations: [
          {
            op: 'set_cell',
            sheetId: document.engine.state.sheets[0]!.id,
            address: 'A2',
            value: phase === 'crash' ? 101 : 99,
          },
        ],
      })
      assert.equal(vscode.window.tabGroups.activeTabGroup.activeTab?.isDirty, true)
      let files: string[] = [],
        backupIds: string[] = []
      const deadline = Date.now() + 20_000
      while (!files.length && Date.now() < deadline) {
        files = await nativeBackupFiles(join(root, 'profile', 'Backups'))
        backupIds = await Promise.all(
          files.map(async (path) => {
            const text = await readFile(path, 'utf8')
            return JSON.parse(text.slice(text.indexOf(' ') + 1)).backupId as string
          }),
        )
        const states = await Promise.all(
          backupIds.map(async (id) => {
            const pointer = JSON.parse(await readFile(vscode.Uri.parse(id).fsPath, 'utf8'))
            const manifest =
              pointer.payload.kind === 'recovery-slot'
                ? JSON.parse(await readFile(pointer.payload.manifestPath, 'utf8'))
                : pointer
            return manifest.payload.stateId as string
          }),
        )
        if (states.some((id) => id !== document.engine.state.stateId)) files = []
        if (!files.length) await new Promise((resolve) => setTimeout(resolve, 50))
      }
      assert.ok(files.length, 'the normal window must persist its native backup index')
      await writeFile(
        marker,
        JSON.stringify({
          sessionId: document.sessionId,
          stateId: document.engine.state.stateId,
          sourceHash: await hash(),
          processId: process.pid,
          nativeBackupFiles: files,
          backupIds,
          version: vscode.version,
        }),
      )
      await writeFile(result, JSON.stringify({ phase, seeded: true }))
      if (phase === 'crash' || process.env.GENOFFICE_RECOVERY_PROOF_FRESH === '1')
        await killTestApplication()
      else await vscode.commands.executeCommand('workbench.action.closeWindow')
      return
    }
    assert.ok(phase === 'restore' || phase === 'recover')
    const prior = JSON.parse(await readFile(marker, 'utf8')) as {
      sessionId: string
      stateId: string
      sourceHash: string
      processId: number
      version: string
    }
    assert.notEqual(process.pid, prior.processId)
    assert.notEqual(document.sessionId, prior.sessionId)
    assert.equal(document.engine.state.stateId, prior.stateId)
    assert.equal(
      document.engine.state.sheets[0]!.cells.A2!.rawValue,
      phase === 'recover' ? 99 : 101,
    )
    assert.equal(document.engine.descriptor.dirty, true)
    assert.equal(document.engine.history.length, 0)
    assert.equal(
      await hash(),
      prior.sourceHash,
      'the original must remain unchanged until explicit save',
    )
    assert.equal(vscode.window.tabGroups.activeTabGroup.activeTab?.isDirty, true)
    await vscode.commands.executeCommand('workbench.action.files.save')
    assert.equal(document.engine.descriptor.dirty, false)
    assert.equal(vscode.window.tabGroups.activeTabGroup.activeTab?.isDirty, false)
    await vscode.commands.executeCommand('workbench.action.closeAllEditors')
    await document.whenDisposed()
    await writeFile(
      result,
      JSON.stringify({
        phase,
        passed: true,
        version: vscode.version,
        oldSessionId: prior.sessionId,
        newSessionId: document.sessionId,
        recoveredStateId: prior.stateId,
        originalHash: prior.sourceHash,
      }),
    )
    await vscode.commands.executeCommand('workbench.action.closeWindow')
  } catch (error) {
    await writeFile(
      result,
      JSON.stringify({
        phase,
        passed: false,
        error: error instanceof Error ? error.message : String(error),
      }),
    )
    await killTestApplication()
  }
}
