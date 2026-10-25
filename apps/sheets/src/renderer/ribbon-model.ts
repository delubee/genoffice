/** Pure state shared by desktop and embedded spreadsheet command bands. */
export interface TableStyleOptionFlags {
  readonly headerRow: boolean
  readonly totalsRow: boolean
  readonly bandedRows: boolean
  readonly bandedColumns: boolean
  readonly firstColumn: boolean
  readonly lastColumn: boolean
  readonly filterButton: boolean
}

export type TableOptionKey = keyof TableStyleOptionFlags

export interface SelectedTableRibbon extends TableStyleOptionFlags {
  readonly key: string
  readonly name: string
  readonly rangeA1: string
  readonly style: string | null
  readonly isFile: boolean
}

export interface HeaderFooterParts {
  readonly left?: string
  readonly center?: string
  readonly right?: string
}
