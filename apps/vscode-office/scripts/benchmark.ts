import * as assert from 'node:assert/strict'
import * as vscode from 'vscode'
import { cpus, totalmem, platform, arch } from 'node:os'
import { createHash } from 'node:crypto'
import { writeFile, mkdir } from 'node:fs/promises'
import { join } from 'node:path'
import { performance } from 'node:perf_hooks'
import { salesWorkbook } from '../tests/fixtures/sales-workbook'
import type { WorkbookDocument } from '../src/extension/workbook-document'
import { transactionInputSchema } from '@genoffice/workbook-session/schemas'

export async function runBenchmark() {
  const extension = vscode.extensions.getExtension('genoffice.genoffice-workbook')!
  const api = (await extension.activate()) as { getDocument(uri: vscode.Uri): WorkbookDocument }
  const fixture = await salesWorkbook()
  const opening: number[] = [],
    editing: number[] = []
  const output = join(extension.extensionPath, 'test-results')
  const metadata = {
    measuredAt: new Date().toISOString(),
    vscode: vscode.version,
    node: process.versions.node,
    platform: platform(),
    arch: arch(),
    cpu: cpus()[0]?.model,
    logicalCpus: cpus().length,
    totalMemoryBytes: totalmem(),
    fixtureSha256: createHash('sha256').update(fixture).digest('hex'),
    referenceEnvironment: false,
    environmentNote:
      'Actual host recorded; 4 vCPU / 8 GiB / SSD reference environment must be independently provisioned and verified. Storage medium is not inferred.',
    startup:
      'First sample: cold workbook in an already activated extension host; remaining samples: warm host, distinct document sessions. Dependency preparation excluded.',
    measurement:
      'SC-001: openWith start through installed first viewport and two animation frames. SC-005: submit already validated ordinary 1000-cell batch through installed grid values and two animation frames. No model or external program; numeric cells contain no formulas.',
  }
  await mkdir(output, { recursive: true })
  const publish = async (complete: boolean) => {
    const gates = {
      sc001: {
        samples: opening.length,
        underLimit: opening.filter((x) => x <= 10000).length,
        limitMs: 10000,
      },
      sc005: {
        samples: editing.length,
        underLimit: editing.filter((x) => x <= 2000).length,
        limitMs: 2000,
      },
    }
    await writeFile(
      join(output, 'benchmark-raw.json'),
      JSON.stringify(
        {
          ...metadata,
          complete,
          openingMs: opening,
          editingMs: editing,
          gates,
          localThresholdsPassed:
            complete && gates.sc001.underLimit >= 19 && gates.sc005.underLimit >= 19,
          referenceGatePassed: false,
        },
        null,
        2,
      ),
    )
  }
  for (let i = 0; i < 20; i++) {
    const uri = vscode.Uri.joinPath(
      vscode.workspace.workspaceFolders![0]!.uri,
      `benchmark-${i}.xlsx`,
    )
    try {
      await vscode.workspace.fs.writeFile(uri, fixture)
      const started = performance.now()
      await vscode.commands.executeCommand('vscode.openWith', uri, 'genoffice.workbook')
      const document = api.getDocument(uri)
      await document.whenViewReady()
      await document.bridge!.request('view.flush', {})
      const presented = (await document.bridge!.request('view.presentation', {
        stateId: document.engine.state.stateId,
        range: 'D2',
      })) as { visible: boolean; values: unknown[][] }
      assert.equal(presented.visible, true)
      assert.equal(presented.values[0]![0], 100)
      opening.push(performance.now() - started)
      if (i === 0) {
        const sheetId = document.engine.state.sheets[0]!.id
        for (let batch = 0; batch < 20; batch++) {
          const values = Array.from({ length: 50 }, () =>
            Array.from({ length: 20 }, () => batch + 1),
          )
          const input = {
            schemaVersion: 1 as const,
            sessionId: document.sessionId,
            expectedRevision: document.engine.descriptor.revision,
            transactionId: `benchmark-edit-${batch}`,
            origin: 'agent' as const,
            summary: 'Benchmark numeric batch',
            operations: [{ op: 'set_range' as const, sheetId, range: 'A1:T50', values }],
          }
          const validated = transactionInputSchema.parse(input)
          const editStarted = performance.now()
          await document.engine.submit(validated)
          const shown = (await document.bridge!.request('view.presentation', {
            stateId: document.engine.state.stateId,
            range: 'A1:T50',
          })) as { visible: boolean; values: unknown[][] }
          assert.equal(shown.visible, true)
          assert.deepEqual(shown.values, values)
          editing.push(performance.now() - editStarted)
          await publish(false)
        }
      }
      await publish(false)
    } finally {
      await vscode.commands.executeCommand('workbench.action.revertAndCloseActiveEditor')
      await vscode.workspace.fs.delete(uri)
    }
  }
  await publish(true)
}
