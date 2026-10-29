/**
 * Pure mapping from the Fields-pane configuration (field indices) to the
 * add_pivot DSL parts (field captions), shared by the desktop and VS Code.
 */
import type { AddPivotOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import type { PivotLayoutErrorCode } from '@genoffice/xlsx-gateway/domain/pivot-layout'
import { t } from './i18n/locale'
import type { OoXmlPivotConfig, PivotField } from './pivot-field-model'

export function pivotConfigToOpParts(
  config: OoXmlPivotConfig,
  fields: readonly PivotField[],
):
  | string
  | Pick<
      AddPivotOperation,
      'rowFields' | 'columnField' | 'pageFields' | 'groupings' | 'filters' | 'values'
    > {
  // Multi-level rows: mapped to field captions in the dialog's order (outer
  // first).
  const rowFieldLabels: string[] = []
  for (const rowFieldIndex of config.rowFieldIndices) {
    const label = fields[rowFieldIndex]?.label
    if (!label) return t('appInvalidRowField')
    rowFieldLabels.push(label)
  }
  if (rowFieldLabels.length === 0) return t('appNeedRowField')
  // Multi-level columns: mapped to field captions in the dialog's order (outer
  // first); empty array = no column dimension.
  const columnFieldLabels: string[] = []
  for (const colFieldIndex of config.colFieldIndices) {
    const label = fields[colFieldIndex]?.label
    if (!label) return t('appInvalidColumnField')
    columnFieldLabels.push(label)
  }
  const pageFields: NonNullable<AddPivotOperation['pageFields']> = []
  for (const page of config.pageFields) {
    const label = fields[page.fieldIndex]?.label
    if (!label) return t('appInvalidPageField')
    pageFields.push(page.item === null ? label : { field: label, item: page.item })
  }
  // Grouping modes: field index → field caption (DSL groupings reference
  // fields by caption).
  const groupingEntries: NonNullable<AddPivotOperation['groupings']> = []
  for (const { fieldIndex: groupedField, rule } of config.groupings) {
    const label = fields[groupedField]?.label
    if (!label) return t('appInvalidGroupField')
    groupingEntries.push(
      rule.kind === 'date'
        ? { kind: 'date', field: label, dateUnit: rule.dateUnit }
        : { kind: 'range', field: label, rangeStep: rule.rangeStep },
    )
  }
  // Filters: label filters by field caption; value filters apply to the
  // level-1 row field.
  const filterOps: NonNullable<AddPivotOperation['filters']> = []
  for (const { fieldIndex: filteredField, rule } of config.labelFilters) {
    const label = fields[filteredField]?.label
    if (!label) return t('appInvalidLabelFilterField')
    filterOps.push({ kind: 'label', field: label, op: rule.op, value: rule.value.trim() })
  }
  for (const { valueIndex, rule } of config.valueFilters) {
    filterOps.push({
      kind: 'value',
      field: rowFieldLabels[0]!,
      valueIndex,
      op: rule.op,
      ...(rule.count !== undefined ? { count: rule.count } : {}),
      ...(rule.from !== undefined ? { from: rule.from } : {}),
      ...(rule.to !== undefined ? { to: rule.to } : {}),
    })
  }
  let valueFields: AddPivotOperation['values']
  try {
    valueFields = config.values.map((v) => {
      // Calculated fields: field is the new field name, aggregation fixed to
      // sum, formula passed to the DSL verbatim.
      if (v.formula !== undefined) {
        return {
          field: (v.calcName ?? '').trim(),
          agg: 'sum' as const,
          formula: v.formula.trim(),
          ...(v.showDataAs !== undefined ? { showDataAs: v.showDataAs } : {}),
          ...(v.name?.trim() ? { name: v.name.trim() } : {}),
        }
      }
      const label = fields[v.fieldIndex]?.label
      if (!label) throw new Error(t('appInvalidValueFieldIndex', { index: v.fieldIndex }))
      return {
        field: label,
        agg: v.agg,
        ...(v.showDataAs !== undefined ? { showDataAs: v.showDataAs } : {}),
        ...(v.name?.trim() ? { name: v.name.trim() } : {}),
      }
    })
  } catch (error) {
    return error instanceof Error ? error.message : t('appInvalidValueField')
  }
  return {
    rowFields: rowFieldLabels,
    ...(columnFieldLabels.length > 0 ? { columnField: columnFieldLabels } : {}),
    ...(pageFields.length > 0 ? { pageFields } : {}),
    ...(groupingEntries.length > 0 ? { groupings: groupingEntries } : {}),
    ...(filterOps.length > 0 ? { filters: filterOps } : {}),
    values: valueFields,
  }
}

export const LAYOUT_ERROR_KEYS = {
  sourceNeedsRows: 'appPivotSourceNeedsRows',
  sourceRowLimit: 'appPivotSourceRowLimit',
  sourceColLimit: 'appPivotSourceColLimit',
  headerBlank: 'appPivotHeaderBlank',
  headerDuplicate: 'appPivotHeaderDuplicate',
  fieldNotHeader: 'appPivotFieldNotHeader',
  calcFieldNameClash: 'appCalcFieldNameClash',
  calcFieldNameDuplicate: 'appCalcFieldNameDuplicate',
  valueFilterFieldMissing: 'appValueFilterFieldMissing',
  tooManyRowItems: 'appPivotTooManyRowItems',
  tooManyColItems: 'appPivotTooManyColItems',
  tooManyColLines: 'appPivotTooManyColLines',
  tooManyRowLines: 'appPivotTooManyRowLines',
  needsValues: 'appPivotNeedsValues',
  pageItemMissing: 'appPivotPageItemMissing',
} satisfies Record<PivotLayoutErrorCode, Parameters<typeof t>[0]>

const SESSION_ERROR_KEYS = {
  ...LAYOUT_ERROR_KEYS,
  onAddedSheet: 'appPivotOnAddedSheet',
  overlapSource: 'appPivotOverlapSource',
  overlapExisting: 'appPivotOverlapExisting',
  overlapSession: 'appPivotOverlapSession',
  nameUsed: 'appPivotNameUsed',
  sheetNoPivot: 'appSheetNoPivot',
  defNotLoaded: 'appPivotDefNotLoaded',
  sourceSheetMissing: 'appPivotSourceSheetMissing',
  layoutMismatch: 'appPivotLayoutMismatch',
  growUnsupported: 'appPivotGrowUnsupported',
  growConflict: 'appPivotGrowConflict',
} as const satisfies Record<string, Parameters<typeof t>[0]>

/** Localizes a session `pivot_layout:<code>[:<params>]` rejection; other errors pass through. */
export function pivotErrorMessage(error: unknown): string {
  const message = error instanceof Error ? error.message : String(error)
  const match = /pivot_layout:(\w+)(?::(\{[^}]*\}))?/.exec(message)
  const key = match && SESSION_ERROR_KEYS[match[1] as keyof typeof SESSION_ERROR_KEYS]
  if (!key) return message
  let params: Record<string, string | number>
  try {
    params = match[2] ? (JSON.parse(match[2]) as Record<string, string | number>) : {}
  } catch {
    params = {}
  }
  return t(key, params)
}
