import { expect, it } from 'vitest'
import { readFile, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { computeCandidateState } from '@genoffice/xlsx-gateway/domain/session-candidate'
import type { WorkbookOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import { manualWorkbook } from '../../../apps/vscode-office/tests/fixtures/manual-workbook'
import { XlsxSidecarClient } from '../src/sidecar/client'
import { SidecarWorkbookSource } from '../src/sidecar/source'
import { openWorkbookSnapshot } from '../src/open'
import { initialState } from '../src/state'
import { prepareCandidateMaterial } from '../src/candidate-material'
import { SessionPackageWriter } from '../src/package-writer'
import type { WorkbookState } from '../src/domain'
import { useSessionTestResources } from './helpers'
const resources = useSessionTestResources()

type Comment = {
  row: number
  column: number
  author: string
  text: string
  thread?: { done: boolean; replies: unknown[]; text: string }
}

it('saves threaded comments, replies, resolution and conversion to notes through the session', async () => {
  const directory = await resources.directory()
  let path = join(directory, 'threads.xlsx')
  await writeFile(path, await manualWorkbook())
  const client = new XlsxSidecarClient(
    fileURLToPath(
      new URL(
        '../../../apps/sheets/native/xlsx-engine/target/release/xlsx-sidecar' +
          (process.platform === 'win32' ? '.exe' : ''),
        import.meta.url,
      ),
    ),
  )
  try {
    const source = new SidecarWorkbookSource(client)
    const cycle = async (operations: (sheetId: string) => WorkbookOperation[]) => {
      const opened = await openWorkbookSnapshot({
        uri: pathToFileURL(path).href,
        storageDirectory: directory,
        source,
      })
      try {
        const state: WorkbookState = initialState(opened.metadata)
        const ops = operations(state.sheets[0]!.id)
        const result = await computeCandidateState(
          state,
          {
            schemaVersion: 1,
            sessionId: 'threads',
            transactionId: crypto.randomUUID(),
            expectedRevision: 0,
            origin: 'manual',
            summary: 'Comments',
            operations: ops,
          },
          await prepareCandidateMaterial({
            state,
            operations: ops,
            source,
            sourceSessionId: opened.original.sidecarSessionId,
          }),
        )
        const staged = await new SessionPackageWriter(client, opened.metadata).stage(
          opened.original.path,
          result.nextState,
          join(directory, `stage-${crypto.randomUUID()}.xlsx`),
        )
        path = join(directory, `saved-${crypto.randomUUID()}.xlsx`)
        await writeFile(path, await readFile(staged.stagedPath))
        const reopened = await source.openSnapshot(path)
        try {
          return (reopened.metadata as { sheets: { comments: Comment[] }[] }).sheets[0]!.comments
        } finally {
          await source.close(reopened.sessionId)
        }
      } finally {
        await opened.dispose()
      }
    }
    const thread = {
      id: '{11111111-2222-3333-4444-555555555555}',
      personId: '{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}',
      author: 'Reviewer',
      dT: '2026-10-29T10:00:00.00',
      done: false,
      text: 'Is this total right?',
      replies: [] as {
        id: string
        personId: string
        author: string
        dT: string
        text: string
      }[],
    }
    let comments = await cycle((sheetId) => [{ op: 'set_thread', sheetId, address: 'B2', thread }])
    expect(comments.find((c) => c.row === 1 && c.column === 1)).toMatchObject({
      author: 'Reviewer',
      text: 'Is this total right?',
      thread: { done: false, text: 'Is this total right?' },
    })
    comments = await cycle((sheetId) => [
      {
        op: 'set_thread',
        sheetId,
        address: 'B2',
        thread: {
          ...thread,
          done: true,
          replies: [
            {
              id: '{99999999-2222-3333-4444-555555555555}',
              personId: thread.personId,
              author: 'Reviewer',
              dT: '2026-10-29T11:00:00.00',
              text: 'Yes, checked.',
            },
          ],
        },
      },
    ])
    const replied = comments.find((c) => c.row === 1 && c.column === 1)!
    expect(replied.thread?.done).toBe(true)
    expect(replied.thread?.replies).toHaveLength(1)
    comments = await cycle((sheetId) => [
      { op: 'set_thread', sheetId, address: 'B2', thread: null },
      { op: 'set_note', sheetId, address: 'B2', text: 'Is this total right?\n\nYes, checked.' },
    ])
    const converted = comments.find((c) => c.row === 1 && c.column === 1)!
    expect(converted.thread).toBeUndefined()
    expect(converted.author).toBe('Reviewer')
    expect(converted.text).toContain('Yes, checked.')
  } finally {
    client.stop()
  }
})
