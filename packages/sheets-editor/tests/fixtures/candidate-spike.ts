import { ICommandService, LocaleType, Univer, UniverInstanceType } from '@univerjs/core'
import type { IWorkbookData, Workbook } from '@univerjs/core'
import { FUniver } from '@univerjs/core/lib/facade'
import { UniverSheetsPlugin } from '@univerjs/sheets'
import '@univerjs/sheets/facade'
import { buildChartVisual } from '@genoffice/xlsx-gateway/domain/chart-visual'

export async function runSpike() {
  let activeRuntimes = 0
  const create = (snapshot: Partial<IWorkbookData>) => {
    const univer = new Univer({
      locale: LocaleType.EN_US,
      locales: { [LocaleType.EN_US]: { sheets: {} } },
    })
    activeRuntimes++
    univer.registerPlugin(UniverSheetsPlugin)
    const model = univer.createUnit<IWorkbookData, Workbook>(
      UniverInstanceType.UNIVER_SHEET,
      structuredClone(snapshot),
    )
    const api = FUniver.newAPI(univer)
    return {
      univer,
      model,
      api,
      dispose() {
        univer.dispose()
        activeRuntimes--
      },
    }
  }
  const main = create({
    id: 'spike-main',
    sheetOrder: ['logical-sheet'],
    sheets: {
      'logical-sheet': {
        id: 'logical-sheet',
        name: 'Data',
        rowCount: 20,
        columnCount: 10,
        cellData: {
          0: { 0: { v: 'Label' }, 1: { v: 'Value' } },
          1: { 0: { v: 'A' }, 1: { v: 3 } },
        },
      },
    },
  })
  const before = main.model.getSnapshot()
  const canonicalBefore = JSON.stringify(before)
  let failedCandidateDiscarded = false
  const rejected = create(before)
  try {
    rejected.api.getActiveWorkbook()!.getActiveSheet()!.getRange('B2').setValue(999)
    throw new Error('Injected failure at operation 2')
  } catch {
    failedCandidateDiscarded = JSON.stringify(main.model.getSnapshot()) === canonicalBefore
  } finally {
    rejected.dispose()
  }
  const commands = main.univer.__getInjector().get(ICommandService)
  let intercepted = false
  const guard = commands.beforeCommandExecuted((command) => {
    if (command.id.includes('set-range-values')) {
      intercepted = true
      throw new Error('Forward the manual intent to the host before applying it')
    }
  })
  try {
    main.api.getActiveWorkbook()!.getActiveSheet()!.getRange('B2').setValue(777)
  } catch {
    /* Expected veto. */
  }
  guard.dispose()
  const blockedBeforeMutation =
    intercepted && JSON.stringify(main.model.getSnapshot()) === canonicalBefore
  const candidate = create(before)
  let structuralChange = false
  let serializedChart = false
  let restored = false
  try {
    const sheet = candidate.api.getActiveWorkbook()!.getActiveSheet()!
    sheet.insertRows(1, 1)
    const after = candidate.model.getSnapshot()
    structuralChange =
      after.sheets['logical-sheet']!.rowCount ===
        (before.sheets['logical-sheet']!.rowCount ?? 0) + 1 &&
      after.sheets['logical-sheet']!.cellData![2]![1]!.v === 3
    const chart = buildChartVisual({
      id: 'chart-1',
      sheetId: 'logical-sheet',
      sheetName: 'Data',
      chartType: 'column',
      dataRange: 'A1:B3',
      values: [
        ['Label', 'Value'],
        [null, null],
        ['A', 3],
      ],
    })
    const roundtrip = JSON.parse(
      JSON.stringify({
        forward: { workbook: after, charts: [chart] },
        inverse: { workbook: before, charts: [] },
      }),
    )
    serializedChart = roundtrip.forward.charts[0].id === 'chart-1'
    const inverse = create(roundtrip.inverse.workbook)
    try {
      restored =
        JSON.stringify(inverse.model.getSnapshot()) === canonicalBefore &&
        roundtrip.inverse.charts.length === 0
    } finally {
      inverse.dispose()
    }
  } finally {
    candidate.dispose()
    main.dispose()
  }
  return {
    failedCandidateDiscarded,
    blockedBeforeMutation,
    structuralChange,
    serializedChart,
    restored,
    activeRuntimes,
  }
}

export async function runProductionCandidate() {
  const { IsolatedCandidateInterpreter } = await import('../../src/candidate')
  const state = {
    schemaVersion: 1 as const,
    stateId: 'root-before',
    dateSystem: '1900' as const,
    structuralLog: [],
    sheetMap: {},
    sheets: [
      {
        id: 'logical-sheet',
        name: 'Data',
        rows: 20,
        columns: 10,
        cells: {
          A1: { kind: 'string' as const, rawValue: 'Label' },
          B1: { kind: 'string' as const, rawValue: 'Value' },
          A2: { kind: 'string' as const, rawValue: 'A' },
          B2: { kind: 'number' as const, rawValue: 3 },
        },
      },
    ],
  }
  const { workbookStateSchema } = await import('@genoffice/xlsx-gateway/shared/session-state')
  const before = JSON.stringify(workbookStateSchema.parse(state))
  const intent = {
    schemaVersion: 1,
    sessionId: 'session',
    transactionId: 'candidate-transaction',
    expectedRevision: 0,
    origin: 'manual',
    summary: 'Insert rows and columns',
    operations: [
      { op: 'insert_rows', sheetId: 'logical-sheet', row: 2, count: 1 },
      { op: 'insert_cols', sheetId: 'logical-sheet', column: 'C', count: 1 },
    ],
  }
  const context: import('../../src/candidate').CandidateContext = {
    async hydrate(current) {
      return {
        state: structuredClone(current),
        completeSheetIds: ['logical-sheet'],
        writableSheetIds: ['logical-sheet'],
        allocation: { stateId: `production-after-${current.stateId}`, sheetIds: {} },
      }
    },
  }
  const failed = new IsolatedCandidateInterpreter({
    ...context,
    fault(point) {
      if (point === 'afterOperation:1') throw new Error('Injected candidate failure')
    },
  })
  let rejected = false
  try {
    await failed.prepare(state, intent)
  } catch {
    rejected = true
  }
  const intact =
    JSON.stringify(workbookStateSchema.parse(state)) === before && failed.activeRuntimeCount === 0
  failed.dispose()
  const successful = new IsolatedCandidateInterpreter(context)
  const changes = await successful.prepare(state, intent)
  const restored = JSON.stringify(JSON.parse(JSON.stringify(changes.inverse)).state) === before
  const withChart = await successful.prepare(changes.nextState, {
    ...intent,
    expectedRevision: 1,
    transactionId: 'chart-after-structure',
    operations: [
      { op: 'add_chart', sheetId: 'logical-sheet', chartType: 'column', dataRange: 'A1:B3' },
    ],
  })
  const chart = (withChart.nextState.sheets[0]!.features!.charts as unknown[]).length === 1
  const moved = changes.nextState.sheets[0]!.cells.B3?.rawValue === 3
  const disposed = successful.activeRuntimeCount === 0
  successful.dispose()
  return { rejected, intact, restored, chart, moved, disposed }
}

export async function runCandidateGuards() {
  const { IsolatedCandidateInterpreter } = await import('../../src/candidate')
  const { workbookStateSchema } = await import('@genoffice/xlsx-gateway/shared/session-state')
  const state = workbookStateSchema.parse({
    schemaVersion: 1,
    stateId: 'guards',
    structuralLog: [],
    sheets: [
      {
        id: 'one',
        name: 'Data',
        rows: 20,
        columns: 10,
        cells: {
          A1: { kind: 'number', rawValue: 2 },
          A2: { kind: 'number', rawValue: 3 },
          B2: { kind: 'error', rawValue: '#REF!' },
        },
      },
      {
        id: 'two',
        name: 'Other',
        rows: 20,
        columns: 10,
        cells: { A1: { kind: 'number', rawValue: 3, formula: '=Data!A2' } },
      },
    ],
  })
  const interpreter = new IsolatedCandidateInterpreter({
    async hydrate(current) {
      return {
        state: structuredClone(current),
        completeSheetIds: ['one', 'two'],
        writableSheetIds: ['one', 'two'],
        allocation: { stateId: `guards-after-${current.stateId}`, sheetIds: {} },
      }
    },
  })
  const intent = (operations: unknown[]) => ({
    schemaVersion: 1,
    sessionId: 'session',
    transactionId: crypto.randomUUID(),
    expectedRevision: 0,
    origin: 'agent',
    summary: 'Guard test',
    operations,
  })
  try {
    const literal = await interpreter.prepare(
      state,
      intent([{ op: 'set_cell', sheetId: 'one', address: 'B2', value: '#REF!' }]),
    )
    const result = await interpreter.prepare(
      literal.nextState,
      intent([
        { op: 'insert_rows', sheetId: 'one', row: 2, count: 1 },
        { op: 'rename_sheet', sheetId: 'one', name: 'Renamed' },
      ]),
    )
    const formula = result.nextState.sheets[1]!.cells.A1!.formula
    const literalKind = result.nextState.sheets[0]!.cells.B3!.kind
    state.sheets[0]!.protected = true
    let protectedRejected = false
    try {
      await interpreter.prepare(
        state,
        intent([{ op: 'set_cell', sheetId: 'one', address: 'A1', value: 8 }]),
      )
    } catch (error) {
      protectedRejected = String(error).includes('protected_target')
    }
    return {
      formula,
      literalKind,
      protectedRejected,
      noRuntimes: interpreter.activeRuntimeCount === 0,
    }
  } finally {
    interpreter.dispose()
  }
}
