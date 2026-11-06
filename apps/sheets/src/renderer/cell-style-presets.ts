/** Shared desktop Home cell styles, independent of the application runtime. */
import type { CellFormatPatch } from '@genoffice/xlsx-gateway/domain/workbook-dsl'

export const CELL_STYLE_PRESETS: Record<string, CellFormatPatch[]> = {
  good: [{ fillColor: '#C6EFCE', fontColor: '#006100' }],
  bad: [{ fillColor: '#FFC7CE', fontColor: '#9C0006' }],
  neutral: [{ fillColor: '#FFEB9C', fontColor: '#9C6500' }],
  input: [
    { fillColor: '#FFCC99', fontColor: '#3F3F76', border: { type: 'all', color: '#7F7F7F' } },
  ],
  output: [
    {
      fillColor: '#F2F2F2',
      fontColor: '#3F3F3F',
      bold: true,
      border: { type: 'all', color: '#3F3F3F' },
    },
  ],
  calculation: [
    {
      fillColor: '#F2F2F2',
      fontColor: '#FA7D00',
      bold: true,
      border: { type: 'all', color: '#7F7F7F' },
    },
  ],
  'warning-text': [{ fontColor: '#FF0000' }],
  title: [{ bold: true, fontSize: 18, fontColor: '#44546A' }],
  'heading-1': [
    {
      bold: true,
      fontSize: 15,
      fontColor: '#44546A',
      border: { type: 'bottom', color: '#4472C4' },
    },
  ],
  'heading-2': [
    {
      bold: true,
      fontSize: 13,
      fontColor: '#44546A',
      border: { type: 'bottom', color: '#8EA9DB' },
    },
  ],
  // Excel's Total is thin-top + double-bottom; the double edge is not in the
  // border command set, so the top rule alone approximates it.
  total: [{ bold: true, border: { type: 'top', color: '#4472C4' } }],
  'accent1-20': [{ fillColor: '#DDEBF7' }],
  'accent1-40': [{ fillColor: '#BDD7EE' }],
  accent1: [{ fillColor: '#4472C4', fontColor: '#FFFFFF' }],
}
