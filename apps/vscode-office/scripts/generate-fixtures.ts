import { mkdir, writeFile } from 'node:fs/promises'
import { join, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'
import { createHash } from 'node:crypto'
import {
  salesWorkbook,
  typesWorkbook,
  preservationWorkbook,
  SALES_EXPECTED,
} from '../tests/fixtures/sales-workbook'

const directory =
  process.argv[2] ?? fileURLToPath(new URL('../test-fixtures/generated/', import.meta.url))
await mkdir(directory, { recursive: true })
const files = [
  ['sales-100k.xlsx', await salesWorkbook()],
  ['types-1900.xlsx', await typesWorkbook()],
  ['types-1904.xlsx', await typesWorkbook(true)],
  ['preservation.xlsx', await preservationWorkbook()],
  ['same-name/a/report.xlsx', await typesWorkbook()],
  ['same-name/b/report.xlsx', await typesWorkbook(true)],
  ['damaged.xlsx', Buffer.from('not a ZIP workbook: deterministic damaged fixture\n')],
  ['external-replacement.xlsx', await typesWorkbook(true)],
] as const
const hashes: Record<string, string> = {}
for (const [name, bytes] of files) {
  await mkdir(dirname(join(directory, name)), { recursive: true })
  await writeFile(join(directory, name), bytes)
  hashes[name] = createHash('sha256').update(bytes).digest('hex')
}
await writeFile(
  join(directory, 'expected.json'),
  JSON.stringify({ ...SALES_EXPECTED, fixtureSha256: hashes }, null, 2),
)
console.log(JSON.stringify({ directory, fixtureSha256: hashes }, null, 2))
