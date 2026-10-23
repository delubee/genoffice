import { expect, it } from 'vitest'
import { WorkbookTools } from '../src/extension/tools/service'
import { WorkbookBindings } from '@genoffice/workbook-session/bindings'
import { initialState } from '@genoffice/workbook-session/state'
import type { DocumentDescriptor } from '@genoffice/workbook-session/domain'

it('requires an explicit target, rejects unknown versions and does not register exchange tools prematurely', async () => {
  const tools = new WorkbookTools({
    bindings: new WorkbookBindings(),
    documents: () => [],
    operationStatus: async () => ({ status: 'unknown' }),
  })
  expect(await tools.invoke('context', { schemaVersion: 1 })).toMatchObject({
    status: 'ok',
    data: { requires_target: true, documents: [] },
  })
  expect(await tools.invoke('read', { schemaVersion: 2 })).toMatchObject({
    status: 'error',
    code: 'invalid_input',
  })
  expect(await tools.invoke('export', { schemaVersion: 1 })).toMatchObject({
    status: 'error',
    code: 'unsupported_operation',
  })
  expect(
    await tools.invoke('context', { schemaVersion: 1, documentUri: 'file:///not-open.xlsx' }),
  ).toMatchObject({ status: 'error', code: 'session_closed' })
})

it('prepare only describes effects, while invocation rechecks the explicit target and cancellation', async () => {
  const state = initialState({
    sheets: [{ id: 'base', name: 'Data', rowCount: 2, columnCount: 1 }],
  })
  let descriptor: DocumentDescriptor = {
    documentUri: 'file:///a/report.xlsx',
    sessionId: 'a',
    revision: 0,
    stateId: state.stateId,
    savedStateId: state.stateId,
    dirty: false,
    lifecycle: 'ready',
  }
  const bindings = new WorkbookBindings()
  bindings.observe(descriptor, state)
  const binding = bindings.bind('a', state.sheets[0]!.id, 'A1:A2')
  let flushes = 0,
    reads = 0
  const document = {
    descriptor: () => descriptor,
    state: () => state,
    flush: async () => {
      flushes++
    },
    snapshot() {
      reads++
      throw new Error('prepare must not read')
    },
    submit: async () => {
      throw new Error('prepare must not write')
    },
    operationStatus: async () => ({ status: 'unknown' as const }),
  }
  const tools = new WorkbookTools({
    bindings,
    documents: () => [document],
    operationStatus: async () => ({ status: 'unknown' }),
  })
  const input = {
    schemaVersion: 1,
    bindingId: binding.bindingId,
    expectedRevision: 0,
    transactionId: 'test',
    summary: 'Set a value',
    operations: [{ op: 'set_cell', sheetId: state.sheets[0]!.id, address: 'A1', value: 2 }],
  }
  expect(tools.prepare('apply', input)).toMatchObject({ confirmationRequired: true })
  expect(tools.prepare('apply', input).message).toContain('A1')
  expect(
    await tools.invoke('apply', {
      ...input,
      operations: [
        ...input.operations,
        { op: 'insert_rows', sheetId: state.sheets[0]!.id, row: 1, count: 1 },
      ],
    }),
  ).toMatchObject({ status: 'error', code: 'invalid_input' })
  expect(
    await tools.invoke('apply', {
      ...input,
      operations: [{ ...input.operations[0], unrecognized: true }],
    }),
  ).toMatchObject({ status: 'error', code: 'invalid_input' })
  expect(flushes).toBe(0)
  expect(reads).toBe(0)
  const controller = new AbortController()
  controller.abort()
  expect(await tools.invoke('apply', input, controller.signal)).toMatchObject({
    status: 'error',
    code: 'operation_cancelled',
  })
  expect(flushes).toBe(0)
  descriptor = { ...descriptor, revision: descriptor.revision + 1 }
  expect(await tools.invoke('apply', input)).toMatchObject({
    status: 'error',
    code: 'revision_conflict',
  })
})
