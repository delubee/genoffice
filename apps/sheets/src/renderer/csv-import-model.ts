/**
 * Data → From Text/CSV: host-neutral pieces shared by the desktop renderer
 * and the VS Code editor (file pick, size budget, charset choice, typed matrix).
 */
import { isNumericCell, parseCsv } from '@genoffice/xlsx-gateway/gateway/csv-import'

/// Larger files should open as their own workbook instead.
export const CSV_IMPORT_MAX_BYTES = 32 * 1024 * 1024

/// Matches the paste ceiling in spirit: one setValues command, journaled and
/// undoable; larger files should open as their own workbook instead.
export const CSV_IMPORT_MAX_CELLS = 50_000

/// Excel writes CSV in the system's legacy charset; the UI language is the
/// best tie-breaker we have (same map as the main-process open path).
export const CSV_CHARSET_BY_LANG: Record<string, string> = {
  zh: 'gb18030',
  'zh-TW': 'big5',
  ja: 'shift_jis',
  ko: 'euc-kr',
}

export function pickCsvFile(): Promise<File | null> {
  return new Promise((resolve) => {
    const input = document.createElement('input')
    input.type = 'file'
    input.accept = '.csv,.tsv,.txt,text/csv,text/tab-separated-values,text/plain'
    input.onchange = () => resolve(input.files?.[0] ?? null)
    input.oncancel = () => resolve(null)
    input.click()
  })
}

export type CsvImportMatrix =
  | { readonly rows: number; readonly columns: number; readonly values: (string | number)[][] }
  | { readonly error: 'appCsvEmpty' | 'appCsvTooLarge' }

/** Sniffed delimiter, rectangular padding and numeric typing, as the desktop import writes. */
export function csvImportMatrix(text: string): CsvImportMatrix {
  let rows: string[][]
  try {
    rows = parseCsv(text)
  } catch {
    // the parser's own row/column caps; the desktop reports these as too large
    return { error: 'appCsvTooLarge' }
  }
  const columns = rows.reduce((width, row) => Math.max(width, row.length), 0)
  if (rows.length === 0 || columns === 0) return { error: 'appCsvEmpty' }
  if (rows.length * columns > CSV_IMPORT_MAX_CELLS) return { error: 'appCsvTooLarge' }
  return {
    rows: rows.length,
    columns,
    values: rows.map((row) =>
      Array.from({ length: columns }, (_, index) => {
        const cell = row[index] ?? ''
        return isNumericCell(cell) ? Number(cell) : cell
      }),
    ),
  }
}
