import { WorkbookRuntime } from '../../src/runtime'
import { workbookStateSchema } from '@genoffice/xlsx-gateway/shared/session-state'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import { ICommandService } from '@univerjs/core'
export async function mountManualRuntime() {
  const container = document.createElement('div')
  container.style.width = '100vw'
  container.style.height = '600px'
  document.body.append(container)
  let state = workbookStateSchema.parse({
    schemaVersion: 1,
    stateId: 'manual-before',
    structuralLog: [],
    sheets: [
      {
        id: 'manual-sheet',
        name: 'Data',
        rows: 10,
        columns: 5,
        cells: { A1: { kind: 'number', rawValue: 123 } },
      },
    ],
  })
  let revision = 0,
    requested: any,
    resolveEdit: ((value: unknown) => void) | undefined
  const runtime = new WorkbookRuntime(
    container,
    {
      async read() {
        return {
          cells: Object.entries(state.sheets[0]!.cells).map(([address, cell]) => ({
            address,
            ...cell,
          })),
        }
      },
      submit(intent) {
        requested = intent
        return new Promise((resolve) => {
          resolveEdit = resolve
        })
      },
      selection() {},
      async history() {
        throw new Error('History not part of this fixture')
      },
    },
    (error) => {
      throw error
    },
  )
  await runtime.install(state.stateId, revision, state)
  const facade = (runtime as unknown as { runtime: { univerAPI: any } }).runtime.univerAPI
  return {
    value: () => facade.getActiveWorkbook().getActiveSheet().getRange('A1').getValue(),
    mutateDirectly() {
      const service = (runtime as any).runtime.univer.__getInjector().get(ICommandService)
      return service.syncExecuteCommand('sheet.mutation.set-range-values', {
        unitId: 'projection',
        subUnitId: 'manual-sheet',
        cellValue: { 0: { 0: { v: 456 } } },
      })
    },
    type(value: number) {
      try {
        facade.getActiveWorkbook().getActiveSheet().getRange('A1').setValue(value)
      } catch (error) {
        if (!String(error).includes('manual_intent_forwarded')) throw error
      }
    },
    requested: () => requested,
    style: () => facade.getActiveWorkbook().getActiveSheet().getRange('A1').getCellStyleData(),
    format() {
      try {
        facade.getActiveWorkbook().getActiveSheet().getRange('A1').setFontWeight('bold')
      } catch (error) {
        if (!String(error).includes('manual_intent_forwarded')) throw error
      }
    },
    async commitFormat() {
      const result = await computeCandidateState(
        state,
        {
          schemaVersion: 1,
          sessionId: 'manual',
          origin: 'manual',
          summary: 'Format',
          ...requested,
        },
        {
          state,
          completeSheetIds: ['manual-sheet'],
          writableSheetIds: ['manual-sheet'],
          allocation: { stateId: 'format-after', sheetIds: {} },
        },
      )
      state = result.nextState
      revision++
      await runtime.install(state.stateId, revision, state)
      resolveEdit!({ revision })
    },
    async commit() {
      state = structuredClone(state)
      state.stateId = 'manual-after'
      state.sheets[0]!.cells.A1 = { kind: 'number', rawValue: requested.operations[0].value }
      revision++
      await runtime.install(state.stateId, revision, state)
      resolveEdit!({ revision })
      await Promise.resolve()
    },
    composing(value: boolean) {
      container.dispatchEvent(
        new CompositionEvent(value ? 'compositionstart' : 'compositionend', { bubbles: true }),
      )
    },
    flush: () => runtime.flushEditor(),
    dispose() {
      runtime.dispose()
      container.remove()
    },
  }
}
