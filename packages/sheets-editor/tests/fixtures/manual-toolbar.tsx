import { createRoot } from 'react-dom/client'
import { useState } from 'react'
import type { WorkbookAppearance } from '../../src/appearance'
import { WorkbookEditor } from '../../src/WorkbookEditor'
import type { WorkbookRuntime } from '../../src/runtime'
import {
  workbookStateSchema,
  transactionInputSchema,
} from '@genoffice/xlsx-gateway/shared/session-state'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import type { EditorHost } from '../../src/ports'
import '../../../../apps/vscode-office/src/webview/theme.css'
import { ThemeService } from '@univerjs/core'

/** In-memory authority for UI tests; native sidecar/save are tested separately. */
export async function mountManualToolbar(appearanceControls = false, merged = false) {
  let state = workbookStateSchema.parse({
      schemaVersion: 1,
      stateId: 'toolbar-before',
      structuralLog: [],
      sheets: [
        {
          id: 'toolbar-sheet',
          name: 'Data',
          rows: 3,
          columns: 2,
          cells: {
            A1: { kind: 'string', rawValue: 'Region' },
            B1: { kind: 'string', rawValue: 'Sales' },
            A2: { kind: 'string', rawValue: 'North' },
            B2: { kind: 'number', rawValue: 120 },
            A3: { kind: 'string', rawValue: 'South' },
            B3: { kind: 'number', rawValue: -30 },
          },
          features: merged
            ? { merges: [{ startRow: 0, startColumn: 0, endRow: 0, endColumn: 1 }] }
            : {},
        },
      ],
    }),
    revision = 0
  let projection: WorkbookRuntime
  const host: EditorHost = {
    async read(sheetId) {
      return {
        merges:
          state.sheets.find((sheet) => sheet.id === sheetId)!.features?.mergeEdits ??
          state.sheets.find((sheet) => sheet.id === sheetId)!.features?.merges ??
          [],
        cells: Object.entries(state.sheets.find((sheet) => sheet.id === sheetId)!.cells).map(
          ([address, cell]) => ({ address, ...cell }),
        ),
      }
    },
    async submit(raw) {
      const intent = transactionInputSchema.parse({
        schemaVersion: 1,
        sessionId: 'toolbar',
        origin: 'manual',
        summary: 'UI edit',
        ...raw,
      })
      if (intent.expectedRevision !== revision) throw new Error('stale_revision')
      const result = await computeCandidateState(state, intent, {
        state,
        completeSheetIds: state.sheets.map((sheet) => sheet.id),
        writableSheetIds: state.sheets.map((sheet) => sheet.id),
        allocation: {
          stateId: crypto.randomUUID(),
          sheetIds: Object.fromEntries(
            intent.operations.flatMap((op, index) =>
              op.op === 'add_sheet' ? [[String(index), crypto.randomUUID()]] : [],
            ),
          ),
        },
      })
      state = result.nextState
      revision++
      await projection.install(state.stateId, revision, state)
      return { revision }
    },
    selection() {},
    async history() {
      throw new Error('History is outside this fixture')
    },
  }
  const root = createRoot(document.body)
  if (appearanceControls) document.documentElement.dataset.workbookAppearance = 'excel'
  await new Promise<void>((resolve, reject) => {
    const onReady = (runtime: WorkbookRuntime) => {
      projection = runtime
      void runtime.install(state.stateId, revision, state).then(resolve, reject)
    }
    function Fixture() {
      const [appearance, setAppearance] = useState<WorkbookAppearance>('excel')
      return (
        <WorkbookEditor
          host={host}
          locale="en"
          onReady={onReady}
          appearance={appearanceControls ? appearance : 'vscode'}
          onAppearanceChange={
            appearanceControls
              ? (value) => {
                  document.documentElement.dataset.workbookAppearance = value
                  setAppearance(value)
                }
              : undefined
          }
        />
      )
    }
    root.render(<Fixture />)
  })
  const facade = (projection! as unknown as { runtime: { univerAPI: any } }).runtime.univerAPI
  return {
    beginMergedEdit() {
      const book = facade.getActiveWorkbook()
      const sheet = book.getActiveSheet()
      sheet.setActiveRange(sheet.getRange('A1:B1'))
      return book.startEditing()
    },
    select(range: string) {
      const sheet = facade.getActiveWorkbook().getActiveSheet()
      sheet.setActiveRange(sheet.getRange(range))
    },
    selectSheet(name: string) {
      const sheet = state.sheets.find((sheet) => sheet.name === name)!
      facade.getActiveWorkbook().setActiveSheet(sheet.id)
    },
    value(address: string) {
      return facade.getActiveWorkbook().getActiveSheet().getRange(address).getValue()
    },
    sizes: () => projection.projectedDimensions('A1:C3'),
    flush: () => projection.flushEditor(),
    cellRect: (address: string) => {
      const cell = facade.getActiveWorkbook().getActiveSheet().getRange(address).getCellRect()
      return { left: cell.left, right: cell.right, top: cell.top, bottom: cell.bottom }
    },
    view: () => projection.projectedView(),
    merges: () => projection.projectedMerges(),
    rowVisible: (row: number) =>
      facade.getActiveWorkbook().getActiveSheet().getSheet().getRowVisible(row),
    formula() {
      try {
        facade.getActiveWorkbook().getActiveSheet().getRange('C2').setFormula('=B2*2')
      } catch (error) {
        if (!String(error).includes('manual_intent_forwarded')) throw error
      }
    },
    snapshot: () => structuredClone(state),
    revision: () => revision,
    darkTheme: () =>
      (
        projection as unknown as {
          runtime: {
            univer: { __getInjector(): { get(token: typeof ThemeService): ThemeService } }
          }
        }
      ).runtime.univer
        .__getInjector()
        .get(ThemeService).darkMode,
    editNote(value: string) {
      return projection.enqueue([
        { op: 'set_note', sheetId: 'toolbar-sheet', address: 'B2', text: value },
      ])
    },
    edit() {
      return projection.enqueue([
        { op: 'set_cell', sheetId: 'toolbar-sheet', address: 'B2', value: 200 },
      ])
    },
    dispose() {
      root.unmount()
    },
  }
}
