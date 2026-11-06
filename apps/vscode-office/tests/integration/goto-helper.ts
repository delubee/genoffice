import * as vscode from 'vscode'
import type { Frame } from '@playwright/test'

/**
 * Go To a reference through the shared dialog. Under a full-suite run a toast or a busy view can
 * leave the dialog open after Go, so a stuck dialog is dismissed and the sequence retried.
 */
export async function gotoReference(frame: Frame, reference: string): Promise<void> {
  const attempts = 4
  for (let attempt = 1; ; attempt++) {
    await vscode.commands.executeCommand('notifications.clearAll')
    await vscode.commands.executeCommand('genoffice.workbook.goto')
    const go = frame.getByRole('dialog', { name: 'Go To', exact: true })
    await go.waitFor({ state: 'visible', timeout: 10_000 })
    await go.getByRole('textbox', { name: 'Reference', exact: true }).fill(reference)
    await go.getByRole('button', { name: 'Go', exact: true }).click()
    try {
      await go.waitFor({ state: 'hidden', timeout: 8_000 })
      return
    } catch (error) {
      if (attempt >= attempts) throw error
      await frame.page().keyboard.press('Escape')
      await go.waitFor({ state: 'hidden', timeout: 5_000 }).catch(() => undefined)
    }
  }
}
