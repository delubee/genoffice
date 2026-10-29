import { describe, expect, it } from 'vitest'
import { CSV_IMPORT_MAX_CELLS, csvImportMatrix } from '../src/renderer/csv-import-model'

describe('csvImportMatrix', () => {
  it('pads ragged rows and types numeric cells', () => {
    expect(csvImportMatrix('a,b,c\n1,2.5\n"x,y",-3,')).toEqual({
      rows: 3,
      columns: 3,
      values: [
        ['a', 'b', 'c'],
        [1, 2.5, ''],
        ['x,y', -3, ''],
      ],
    })
  })

  it('reports empty input and the desktop cell budget', () => {
    expect(csvImportMatrix('')).toEqual({ error: 'appCsvEmpty' })
    const tall = Array.from({ length: CSV_IMPORT_MAX_CELLS / 10 + 1 }, () => '1,2,3,4,5,6,7,8,9,10')
    expect(csvImportMatrix(tall.join('\n'))).toEqual({ error: 'appCsvTooLarge' })
    const wide = Array.from({ length: 20_000 }, () => '1').join(',')
    expect(csvImportMatrix(wide)).toEqual({ error: 'appCsvTooLarge' })
  })
})
