import { execFile } from 'node:child_process'
import { userInfo } from 'node:os'
import * as vscode from 'vscode'

let cached: Promise<string> | undefined

/**
 * The name new comments, notes and tracked changes carry: the
 * `genoffice.authorName` setting, else git's user.name, else the OS account.
 */
export function authorName(): Promise<string> {
  const configured = vscode.workspace
    .getConfiguration('genoffice')
    .get<string>('authorName')
    ?.trim()
  if (configured) return Promise.resolve(configured)
  cached ??= new Promise<string>((resolve) => {
    execFile('git', ['config', '--get', 'user.name'], { timeout: 2000 }, (error, stdout) => {
      const name = error ? '' : stdout.trim()
      if (name) return resolve(name)
      try {
        resolve(userInfo().username)
      } catch {
        resolve('')
      }
    })
  })
  return cached
}
