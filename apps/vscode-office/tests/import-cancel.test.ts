import { expect, it } from 'vitest'
import { readFile, readdir, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { withTestWorkspace } from './helpers'
import { importFixture } from '../../../packages/workbook-session/tests/import-fixture'

it('cancels before the first commit without a new sheet and retains the user workspace script/result', async () => {
  await withTestWorkspace(async (directory) => {
    const fixture = await importFixture(directory)
    try {
      const script = join(fixture.snapshot.directory, 'analysis.py')
      await writeFile(script, '# User workspace program\n')
      const dataPath = join(fixture.snapshot.directory, 'result.ndjson'),
        before = await readFile(dataPath),
        state = fixture.engine.state
      const plan = await fixture.prepare(),
        controller = new AbortController()
      controller.abort()
      await expect(
        fixture.imports.apply({
          planId: plan.planId,
          bindingId: 'task-binding',
          cursor: 0,
          expectedRevision: 0,
          transactionId: 'cancel-first',
          engine: fixture.engine,
          signal: controller.signal,
        }),
      ).rejects.toMatchObject({
        code: 'operation_cancelled',
        details: { progress: { completedCount: 0, status: 'cancelled' } },
      })
      expect(fixture.engine.state).toBe(state)
      expect(fixture.engine.state.sheets).toHaveLength(1)
      expect(await readFile(script, 'utf8')).toBe('# User workspace program\n')
      expect(await readFile(dataPath)).toEqual(before)
      expect(
        (await readdir(join(directory, 'private'))).filter((name) =>
          name.startsWith('prepared-import-'),
        ),
      ).toHaveLength(0)
    } finally {
      await fixture.dispose()
    }
  })
})

it('invalidates closed-session plans while committed history and user files remain independent of private staging cleanup', async () => {
  await withTestWorkspace(async (directory) => {
    const fixture = await importFixture(directory)
    try {
      const plan = await fixture.prepare(),
        originalId = fixture.engine.state.stateId
      const result = await fixture.imports.apply({
        planId: plan.planId,
        bindingId: 'task-binding',
        cursor: 0,
        expectedRevision: 0,
        transactionId: 'committed-before-close',
        engine: fixture.engine,
      })
      const committedId = fixture.engine.state.stateId
      await fixture.imports.invalidateSession(fixture.engine.descriptor.sessionId)
      expect(() => fixture.imports.status(plan.planId)).toThrow()
      expect(() => fixture.snapshots.get(fixture.snapshot.manifest.snapshotId)).toThrow()
      expect(fixture.engine.state.stateId).toBe(committedId)
      expect((await fixture.engine.operationStatus(result.receipt.transactionId)).status).toBe(
        'committed',
      )
      await fixture.engine.undo()
      expect(fixture.engine.state.stateId).toBe(originalId)
      await fixture.engine.redo()
      expect(fixture.engine.state.stateId).toBe(committedId)
      expect(await readFile(fixture.resultManifestPath, 'utf8')).toContain('workbook-result')
    } finally {
      await fixture.dispose()
    }
  })
})
