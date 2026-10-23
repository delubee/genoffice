import * as vscode from 'vscode'
import { z } from 'zod'
import { WorkbookError } from '@genoffice/workbook-session/domain'
import { parseRange } from '@genoffice/xlsx-gateway/domain/cell-address'
import type { WorkbookEditorProvider } from '../provider'
import { flushCurrentEditor } from '../tools/current-state'

const selectionSchema = z.strictObject({
  stateId: z.string(),
  revision: z.number().int().nonnegative(),
  selection: z.object({ sheetId: z.string(), address: z.string() }).nullable(),
})

export function registerBindSelection(provider: WorkbookEditorProvider) {
  return vscode.commands.registerCommand('genoffice.bindSelection', async (uri?: vscode.Uri) => {
    const tab = vscode.window.tabGroups.activeTabGroup.activeTab?.input
    const target =
      uri ??
      (tab instanceof vscode.TabInputCustom && tab.viewType === 'genoffice.workbook'
        ? tab.uri
        : undefined)
    const document = target ? provider.getDocument(target) : undefined
    if (!document)
      throw new WorkbookError(
        'requires_target',
        'Focus an open GenOffice workbook before binding its selection.',
      )
    await flushCurrentEditor(document)
    if (!document.bridge)
      throw new WorkbookError('view_unavailable', 'Open the workbook view to bind its selection.')
    const view = selectionSchema.parse(await document.bridge.request('view.selection', {}))
    if (!view.selection)
      throw new WorkbookError('requires_target', 'Select cells in the workbook before binding.')
    if (
      view.stateId !== document.engine.state.stateId ||
      view.revision !== document.engine.descriptor.revision
    )
      throw new WorkbookError(
        'revision_conflict',
        'The selection projection is stale; wait for the view and bind again.',
        {},
        true,
      )
    parseRange(view.selection.address)
    const binding = provider.bindings.bind(
      document.sessionId,
      view.selection.sheetId,
      view.selection.address,
    )
    await vscode.env.clipboard.writeText(
      JSON.stringify({
        schemaVersion: 1,
        bindingId: binding.bindingId,
        documentUri: binding.documentUri,
        sheetId: binding.logicalSheetId,
        range: binding.frozenRange,
        revision: binding.boundRevision,
      }),
    )
    void vscode.window.showInformationMessage(
      `GenOffice: copied workbook binding ${binding.bindingId}. Paste it into VS Code Chat.`,
    )
    return binding
  })
}
