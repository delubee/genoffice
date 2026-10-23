import { expect, it } from 'vitest'
import { writeFile, readFile, unlink, mkdir } from 'node:fs/promises'
import { join } from 'node:path'
import { captureFileBaseline, publishStagedWorkbook } from '@genoffice/workbook-session/publish'
import { withTestWorkspace } from './helpers'

async function publicationFixture(run: (targetPath: string, stagedPath: string) => Promise<void>) {
  await withTestWorkspace(async (directory) => {
    const targetPath = join(directory, 'book.xlsx'),
      stage = join(directory, 'private-stage'),
      stagedPath = join(stage, 'ready.xlsx')
    await mkdir(stage, { mode: 0o700 })
    await writeFile(targetPath, 'original package')
    await writeFile(stagedPath, 'edited package')
    await run(targetPath, stagedPath)
  })
}

it('rejects an external replacement even when it has the same bytes as the baseline', async () => {
  await publicationFixture(async (targetPath, stagedPath) => {
    const baseline = await captureFileBaseline(targetPath)
    await unlink(targetPath)
    await writeFile(targetPath, 'original package')
    await expect(publishStagedWorkbook({ targetPath, stagedPath, baseline })).rejects.toMatchObject(
      { code: 'external_change_conflict' },
    )
    expect(await readFile(stagedPath, 'utf8')).toBe('edited package')
    expect(await readFile(targetPath, 'utf8')).toBe('original package')
  })
})

it('rechecks the actual destination after staging and before promotion', async () => {
  await publicationFixture(async (targetPath, stagedPath) => {
    const baseline = await captureFileBaseline(targetPath)
    await expect(
      publishStagedWorkbook({
        targetPath,
        stagedPath,
        baseline,
        async fault(point) {
          if (point === 'beforePublish') await writeFile(targetPath, 'external content')
        },
      }),
    ).rejects.toMatchObject({ code: 'external_change_conflict' })
    expect(await readFile(targetPath, 'utf8')).toBe('external content')
    expect(await readFile(stagedPath, 'utf8')).toBe('edited package')
  })
})

it('detects post-publication interference and retains old and candidate copies without overwriting the external file', async () => {
  await publicationFixture(async (targetPath, stagedPath) => {
    const baseline = await captureFileBaseline(targetPath)
    let failure: any
    try {
      await publishStagedWorkbook({
        targetPath,
        stagedPath,
        baseline,
        async fault(point) {
          if (point === 'afterPublish') await writeFile(targetPath, 'external after publish')
        },
      })
    } catch (error) {
      failure = error
    }
    expect(failure).toMatchObject({
      code: 'external_change_conflict',
      details: { published: true },
    })
    expect(await readFile(targetPath, 'utf8')).toBe('external after publish')
    expect(await readFile(failure.details.previousPath, 'utf8')).toBe('original package')
    expect(await readFile(failure.details.stagedPath, 'utf8')).toBe('edited package')
  })
})
