import { expect, it } from 'vitest'
import { spawn } from 'node:child_process'

function run(executable: string, args: string[], signal?: AbortSignal) {
  return new Promise<{ code: number | null; output: string }>((resolve, reject) => {
    const child = spawn(executable, args, { signal, stdio: ['ignore', 'pipe', 'ignore'] })
    let output = ''
    child.stdout.on('data', (chunk: Buffer) => {
      output += chunk.toString()
    })
    child.once('error', reject)
    child.once('exit', (code) => resolve({ code, output }))
  })
}

it('uses an existing standard-library Python interpreter and reports missing executables, script/dependency errors and cancellation without installing anything', async () => {
  const analysis = await run('python3', [
    '-c',
    'from decimal import Decimal; import json; print(json.dumps({"total":str(Decimal("100")+Decimal("20"))}))',
  ])
  expect(analysis.code).toBe(0)
  expect(JSON.parse(analysis.output)).toEqual({ total: '120' })
  await expect(run('/no-such-genoffice-test-interpreter', [])).rejects.toMatchObject({
    code: 'ENOENT',
  })
  expect((await run('python3', ['-c', 'raise RuntimeError("test script failure")'])).code).not.toBe(
    0,
  )
  expect(
    (await run('python3', ['-c', 'import genoffice_nonexistent_test_dependency'])).code,
  ).not.toBe(0)
  const controller = new AbortController()
  controller.abort()
  await expect(
    run('python3', ['-c', 'print("cancelled")'], controller.signal),
  ).rejects.toMatchObject({ name: 'AbortError' })
})
