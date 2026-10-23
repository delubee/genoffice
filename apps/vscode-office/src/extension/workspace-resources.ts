import * as vscode from 'vscode'
import { realpath } from 'node:fs/promises'
import { WorkbookError } from '@genoffice/workbook-session/domain'
import { selectWorkspace, assertHostFile } from './resource-policy'

export function chooseWorkspaceRoot(documentUri: vscode.Uri, requestedUri?: string) {
  return selectWorkspace({
    trusted: vscode.workspace.isTrusted,
    roots: (vscode.workspace.workspaceFolders ?? []).map((root) => root.uri),
    requested: requestedUri,
    containing: vscode.workspace.getWorkspaceFolder(documentUri)?.uri,
  })
}

export async function assertCurrentWorkspaceRoot(path: string) {
  if (!vscode.workspace.isTrusted)
    throw new WorkbookError('untrusted_workspace', 'Trust this workspace before exchanging files.')
  const canonical = await realpath(path)
  const roots = await Promise.all(
    (vscode.workspace.workspaceFolders ?? [])
      .filter((root) => root.uri.scheme === 'file')
      .map((root) => realpath(root.uri.fsPath)),
  )
  if (!roots.includes(canonical))
    throw new WorkbookError(
      'path_not_allowed',
      'The snapshot root is no longer a workspace in this window.',
    )
}

export function workspaceResultPath(value: string, root: string) {
  const uri = vscode.Uri.parse(value),
    authority = vscode.Uri.file(root).authority
  assertHostFile(uri, authority)
  return uri.fsPath
}
