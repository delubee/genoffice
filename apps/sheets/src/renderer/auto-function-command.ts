import { formatAddress, parseRange } from '@genoffice/xlsx-gateway/domain/cell-address'

/** Desktop AutoSum writes one formula below each selected column. */
export function autoFunctionCells(functionName: string, address: string) {
  if (!/^(SUM|AVERAGE|COUNT|MIN|MAX)$/.test(functionName)) throw new Error('invalid_auto_function')
  const range = parseRange(address)
  if (range.endRow <= range.startRow) return []
  if (range.endRow >= 1048575) throw new Error('invalid_summary_target')
  return Array.from({ length: range.endColumn - range.startColumn + 1 }, (_, offset) => {
    const column = range.startColumn + offset
    return {
      address: formatAddress(range.endRow + 1, column),
      formula: `=${functionName}(${formatAddress(range.startRow, column)}:${formatAddress(range.endRow, column)})`,
    }
  })
}
