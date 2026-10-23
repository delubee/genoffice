import { describe, expect, it } from 'vitest'
import {
  cellRecordSchema,
  workbookStateSchema,
  transactionInputSchema,
  protocolMessageSchema,
} from '../src/schemas'

describe('untrusted workbook boundaries', () => {
  it('preserves empty text separately from blank cells and rejects non-finite numbers', () => {
    expect(cellRecordSchema.parse({ kind: 'string', rawValue: '' }).rawValue).toBe('')
    expect(cellRecordSchema.parse({ kind: 'blank', rawValue: null }).rawValue).toBeNull()
    expect(cellRecordSchema.safeParse({ kind: 'number', rawValue: Infinity }).success).toBe(false)
    expect(cellRecordSchema.safeParse({ kind: 'string', rawValue: 123 }).success).toBe(false)
    expect(cellRecordSchema.parse({ kind: 'string', rawValue: '=1+1' }).formula).toBeUndefined()
  })

  it('rejects unknown wire versions, executable fields and invalid sheet addresses', () => {
    const message = {
      protocolVersion: 1,
      requestId: 'r',
      sessionId: 's',
      method: 'view.ready',
      payload: {},
    }
    expect(protocolMessageSchema.safeParse(message).success).toBe(true)
    expect(protocolMessageSchema.safeParse({ ...message, protocolVersion: 2 }).success).toBe(false)
    expect(protocolMessageSchema.safeParse({ ...message, method: 'run.python' }).success).toBe(
      false,
    )
    expect(protocolMessageSchema.safeParse({ ...message, shell: 'anything' }).success).toBe(false)
    const state = {
      schemaVersion: 1,
      stateId: 'state-0',
      sheets: [
        {
          id: 'sheet-0',
          name: 'Data',
          rows: 10,
          columns: 10,
          cells: { A0: { kind: 'blank', rawValue: null } },
        },
      ],
      structuralLog: [],
    }
    expect(workbookStateSchema.safeParse(state).success).toBe(false)
    expect(
      workbookStateSchema.safeParse({
        ...state,
        sheets: [{ ...state.sheets[0], cells: { A1: { kind: 'string', rawValue: '00123' } } }],
      }).success,
    ).toBe(true)
  })

  it('requires explicit transaction identity, revision and typed DSL operations', () => {
    const input = {
      schemaVersion: 1,
      sessionId: 's',
      transactionId: 't',
      expectedRevision: 0,
      origin: 'agent',
      summary: 'Update A1',
      operations: [{ op: 'set_cell', sheetId: 'sheet-0', address: 'A1', value: 1 }],
    }
    expect(transactionInputSchema.safeParse(input).success).toBe(true)
    expect(transactionInputSchema.safeParse({ ...input, expectedRevision: -1 }).success).toBe(false)
    expect(
      transactionInputSchema.safeParse({ ...input, operations: [{ op: 'exec', code: 'anything' }] })
        .success,
    ).toBe(false)
    expect(transactionInputSchema.safeParse({ ...input, transactionId: '' }).success).toBe(false)
  })
})
