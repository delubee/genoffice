import { expect, it } from 'vitest'
import { moveFormulaRefs, type FormulaMoveSpec } from '../src/domain/formula-shift'
const spec: FormulaMoveSpec = {
  sourceSheetName: 'Data',
  targetSheetName: "Other's Data",
  source: { startRow: 1, endRow: 2, startColumn: 0, endColumn: 1 },
  target: { startRow: 4, startColumn: 3 },
}
it('moves relative and absolute dependencies without treating a cut as a fill', () => {
  expect(moveFormulaRefs('=A2+$B$3+C4+SUM(A2:B3)+SUM(A1:B3)', spec, 'Data')).toBe(
    "='Other''s Data'!D5+'Other''s Data'!$E$6+C4+SUM('Other''s Data'!D5:E6)+SUM(A1:B3)",
  )
  expect(moveFormulaRefs('=Data!A2+Data!$B$3+Other!A2', spec, 'Third')).toBe(
    "='Other''s Data'!D5+'Other''s Data'!$E$6+Other!A2",
  )
})
it('keeps moved formulas tied to original outside dependencies and relocates internal ones', () => {
  expect(moveFormulaRefs('=A2+$C$4+SUM(A:A)+SUM(1:3)', spec, 'Data', "Other's Data")).toBe(
    "=D5+'Data'!$C$4+SUM('Data'!A:A)+SUM('Data'!1:3)",
  )
  expect(moveFormulaRefs('=A2+C4', { ...spec, targetSheetName: 'Data' }, 'Data')).toBe('=D5+C4')
})
it('preserves strings, table column names, external workbooks, 3D refs and names', () => {
  expect(
    moveFormulaRefs(
      '="A2"&Table1[A2]+\'[Book.xlsx]Data\'!A2+[Book.xlsx]Data!A2+Sheet1:Data!A2+LOG10(1)+A2name',
      spec,
      'Data',
    ),
  ).toBe(
    '="A2"&Table1[A2]+\'[Book.xlsx]Data\'!A2+[Book.xlsx]Data!A2+Sheet1:Data!A2+LOG10(1)+A2name',
  )
})

it('does not mistake short 3D sheet names for moved cell coordinates', () => {
  expect(
    moveFormulaRefs(
      '=SUM(S1:S2!A2)+S1!A2',
      { ...spec, source: { startRow: 0, endRow: 1, startColumn: 18, endColumn: 18 } },
      'Data',
    ),
  ).toBe('=SUM(S1:S2!A2)+S1!A2')
})
