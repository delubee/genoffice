import type { Frame } from '@playwright/test'

export async function pickOutlineCommand(frame: Frame, name: string) {
  const menu = /^(Ungroup|Show Detail|Clear Outline)/.test(name) ? 'Ungroup' : 'Group'
  await frame.getByRole('button', { name: menu, exact: true }).click()
  await frame.getByRole('option', { name, exact: true }).click()
}
