import { expect, it } from 'vitest'
import { WorkbookTools } from '../src/extension/tools/service'
import { WorkbookBindings } from '@genoffice/workbook-session/bindings'
import { initialState } from '@genoffice/workbook-session/state'
import type { DocumentDescriptor } from '@genoffice/workbook-session/domain'

it('does not use focus or basename to replace a frozen task binding', async () => {
  const bindings = new WorkbookBindings()
  const documents = ['file:///a/report.xlsx', 'file:///b/report.xlsx'].map((uri, index) => {
    const state = initialState({
      sheets: [{ id: 'base', name: 'Data', rowCount: 3, columnCount: 1 }],
    })
    const descriptor: DocumentDescriptor = {
      documentUri: uri,
      sessionId: String(index),
      revision: 0,
      stateId: state.stateId,
      savedStateId: state.stateId,
      dirty: false,
      lifecycle: 'ready',
    }
    bindings.observe(descriptor, state)
    return {
      descriptor: () => descriptor,
      state: () => state,
      flush: async () => {},
      snapshot() {
        throw new Error('context must not materialize cells')
      },
      submit: async () => {
        throw new Error('read-only context')
      },
      operationStatus: async () => ({ status: 'unknown' as const }),
    }
  })
  const tools = new WorkbookTools({
    bindings,
    documents: () => documents,
    operationStatus: async () => ({ status: 'unknown' }),
  })
  const a = bindings.bind('0', documents[0]!.state().sheets[0]!.id, 'A1:A3')
  documents.reverse()
  expect(await tools.invoke('context', { schemaVersion: 1, bindingId: a.bindingId })).toMatchObject(
    {
      status: 'ok',
      sessionId: '0',
      data: { binding: { documentUri: 'file:///a/report.xlsx', frozenRange: 'A1:A3' } },
    },
  )
  expect(await tools.invoke('context', { schemaVersion: 1 })).toMatchObject({
    data: { requires_target: true },
  })
  expect(bindings.get(a.bindingId).frozenRange).toBe('A1:A3')
})
