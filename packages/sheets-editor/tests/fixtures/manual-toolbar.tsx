import { createRoot } from 'react-dom/client'
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
export async function mountManualToolbar() {
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
        },
      ],
    }),
    revision = 0
  let projection: WorkbookRuntime
  const host: EditorHost = {
    async read(sheetId) {
      return {
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
  await new Promise<void>((resolve, reject) =>
    root.render(
      <WorkbookEditor
        host={host}
        locale="en"
        onReady={(runtime) => {
          projection = runtime
          void runtime.install(state.stateId, revision, state).then(resolve, reject)
        }}
      />,
    ),
  )
  const facade = (projection! as unknown as { runtime: { univerAPI: any } }).runtime.univerAPI
  return {
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
    formula() {
      try {
        facade.getActiveWorkbook().getActiveSheet().getRange('C2').setFormula('=B2*2')
      } catch (error) {
        if (!String(error).includes('manual_intent_forwarded')) throw error
      }
    },
    snapshot: () => structuredClone(state),
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
