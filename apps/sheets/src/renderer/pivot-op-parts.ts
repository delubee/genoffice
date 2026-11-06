/**
 * Pure mapping from the Fields-pane configuration (field indices) to the
 * add_pivot DSL parts (field captions), shared by the desktop and VS Code.
 */
import type { AddPivotOperation } from '@genoffice/xlsx-gateway/domain/workbook-dsl'
import type { PivotLayoutErrorCode } from '@genoffice/xlsx-gateway/domain/pivot-layout'
import { t, type StringKey } from './i18n/locale'
import type {
  OoXmlPivotConfig,
  PivotField,
  PivotModelError,
  PivotLayoutModel,
  PivotValueSpec,
} from './pivot-field-model'
import {
  defaultDataFieldCaption,
  isPivotAggregation,
} from '@genoffice/xlsx-gateway/domain/pivot-value-modes'
import type {
  WorkbookPivotAdd,
  WorkbookPivotDefinition,
} from '@genoffice/xlsx-gateway/shared/workbook-protocol'

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
      // "Show values as" base fields travel as header names; the base field
      // must sit on rows or columns for the engine to resolve it.
      let baseField: string | undefined
      let baseItem: string | undefined
      if (v.baseFieldIndex !== undefined) {
        baseField = fields[v.baseFieldIndex]?.label
        if (!baseField) throw new Error(t('appInvalidPivotBaseField'))
        if (!rowFieldLabels.includes(baseField) && !columnFieldLabels.includes(baseField)) {
          throw new Error(t('appInvalidPivotBaseField'))
        }
      }
      if (baseField === undefined) baseItem = undefined
      else if (v.baseItem !== undefined) baseItem = v.baseItem
      return {
        field: label,
        agg: v.agg,
        ...(v.showDataAs !== undefined ? { showDataAs: v.showDataAs } : {}),
        ...(baseField !== undefined ? { baseField } : {}),
        ...(baseItem !== undefined ? { baseItem } : {}),
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
  baseFieldMissing: 'appPivotBaseFieldMissing',
  baseFieldNotAxis: 'appPivotBaseFieldNotAxis',
  baseItemMissing: 'appPivotBaseItemMissing',
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
  relayoutPending: 'appPivotDefNotLoadedSave',
  relayoutOverlap: 'appPivotRelayoutOverlap',
  pivotStale: 'appSlicerPivotStale',
  editUnsupported: 'appPivotEditUnsupported',
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

/// Layout model of a journaled pivot (session add or pending relayout).
export function modelOfPivotAdd(add: WorkbookPivotAdd): PivotLayoutModel {
  const columns =
    add.columnFieldIndices ?? (add.columnFieldIndex === undefined ? [] : [add.columnFieldIndex])
  const columnFieldIndices = add.columnFieldIndices ?? columns
  // "Show values as" base item: index into the base field's member list.
  const baseItemLabelOf = (value: WorkbookPivotAdd['values'][number]): string | undefined => {
    if (value.baseFieldIndex === undefined || value.baseItemIndex === undefined) return undefined
    const rowLevel = add.rowFieldIndices.indexOf(value.baseFieldIndex)
    const items =
      rowLevel >= 0
        ? (add.rowLevelItems?.[rowLevel] ?? add.rowItems)
        : (add.colLevelItems?.[columnFieldIndices.indexOf(value.baseFieldIndex)] ??
          add.columnItems ??
          [])
    return items[value.baseItemIndex]
  }
  const values: PivotValueSpec[] = add.values.map((value, index) => {
    const filter = add.filters?.find((entry) => entry.kind === 'value' && entry.dataField === index)
    return {
      fieldIndex: value.fieldIndex,
      agg: value.agg,
      ...(value.showDataAs !== undefined ? { showDataAs: value.showDataAs } : {}),
      ...(value.baseFieldIndex !== undefined ? { baseFieldIndex: value.baseFieldIndex } : {}),
      ...(value.baseItemIndex !== undefined ? { baseItem: baseItemLabelOf(value) ?? '' } : {}),
      ...(value.name !== undefined ? { name: value.name } : {}),
      ...(value.formula !== undefined ? { formula: value.formula, calcName: value.calcName } : {}),
      ...(filter !== undefined && filter.kind === 'value'
        ? {
            filter: {
              op: filter.op,
              ...(filter.count !== undefined ? { count: filter.count } : {}),
              ...(filter.from !== undefined ? { from: filter.from } : {}),
              ...(filter.to !== undefined ? { to: filter.to } : {}),
            },
          }
        : {}),
    }
  })
  const groupings: Record<number, PivotLayoutModel['groupings'][number]> = {}
  for (const grouping of add.groupings ?? []) {
    groupings[grouping.fieldIndex] =
      grouping.kind === 'date'
        ? { kind: 'date', dateUnit: grouping.dateUnit }
        : { kind: 'range', rangeStep: grouping.rangeStep }
  }
  const labelFilters: Record<number, PivotLayoutModel['labelFilters'][number]> = {}
  for (const filter of add.filters ?? []) {
    if (filter.kind === 'label') labelFilters[filter.field] = { op: filter.op, value: filter.value }
  }
  return {
    filters: (add.pageFieldIndices ?? []).map((fieldIndex, page) => {
      const item = add.pageItems?.[page] ?? null
      return {
        fieldIndex,
        item: item === null ? null : (add.pageLevelItems?.[page]?.[item] ?? null),
      }
    }),
    columns,
    rows: [...add.rowFieldIndices],
    values,
    groupings,
    labelFilters,
  }
}

/// Layout model of a file pivot from its parsed definition. Returns a message
/// key when the definition is outside what the pane can edit.
export function modelOfDefinition(
  definition: WorkbookPivotDefinition,
): { model: PivotLayoutModel; sourceFieldCount: number } | string {
  if (definition.unsupported.length > 0) {
    return t('appPivotEditUnsupported', { reasons: definition.unsupported.join('; ') })
  }
  if (definition.rowFields.some((field) => field < 0)) return t('appPivotEditValuesOnRows')
  if (definition.dataFields.some((dataField) => !isPivotAggregation(dataField.subtotal))) {
    return t('appPivotEditAggUnsupported')
  }
  const sourceFieldCount = definition.fields.filter((field) => field.formula === undefined).length
  const memberLabel = (field: number, item: number | null): string | null => {
    if (item === null) return null
    const x = definition.fieldItems[field]?.[item]?.x
    const shared =
      x === null || x === undefined ? undefined : definition.fields[field]?.sharedItems[x]
    return shared === undefined || shared === null ? '' : String(shared)
  }
  const values: PivotValueSpec[] = definition.dataFields.map((dataField, index) => {
    const source = definition.fields[dataField.field]
    const agg = dataField.subtotal as PivotValueSpec['agg']
    const isCalc = source?.formula !== undefined
    const defaultName = defaultDataFieldCaption(agg, source?.name ?? '')
    const filter = definition.filters.find(
      (entry) => entry.kind === 'value' && entry.dataField === index,
    )
    // "Show values as" base item: a fieldItems index of the base field; the
    // model stores its label for the picker.
    const baseItemLabel = (() => {
      if (dataField.baseField === undefined || dataField.baseItem === undefined) return undefined
      const sharedIndex = definition.fieldItems[dataField.baseField]?.[dataField.baseItem]?.x
      const shared =
        sharedIndex === null || sharedIndex === undefined
          ? undefined
          : definition.fields[dataField.baseField]?.sharedItems[sharedIndex]
      return shared === undefined || shared === null ? '' : String(shared)
    })()
    return {
      fieldIndex: isCalc ? -1 : dataField.field,
      agg,
      ...(dataField.showDataAs !== undefined ? { showDataAs: dataField.showDataAs } : {}),
      ...(dataField.baseField !== undefined ? { baseFieldIndex: dataField.baseField } : {}),
      ...(baseItemLabel !== undefined ? { baseItem: baseItemLabel } : {}),
      ...(dataField.name !== defaultName && !isCalc ? { name: dataField.name } : {}),
      ...(isCalc ? { formula: source.formula, calcName: source.name } : {}),
      ...(filter !== undefined && filter.kind === 'value'
        ? {
            filter: {
              op: filter.op,
              ...(filter.count !== undefined ? { count: filter.count } : {}),
              ...(filter.from !== undefined ? { from: filter.from } : {}),
              ...(filter.to !== undefined ? { to: filter.to } : {}),
            },
          }
        : {}),
    }
  })
  // The pane's value filter always targets the level-1 row field.
  if (
    definition.filters.some(
      (filter) => filter.kind === 'value' && filter.field !== definition.rowFields[0],
    )
  ) {
    return t('appPivotEditHasFeatures')
  }
  const groupings: Record<number, PivotLayoutModel['groupings'][number]> = {}
  definition.fields.forEach((field, index) => {
    if (!field.grouping) return
    groupings[index] =
      field.grouping.kind === 'date'
        ? { kind: 'date', dateUnit: field.grouping.dateUnit }
        : { kind: 'range', rangeStep: field.grouping.rangeStep }
  })
  const labelFilters: Record<number, PivotLayoutModel['labelFilters'][number]> = {}
  for (const filter of definition.filters) {
    if (filter.kind === 'label') labelFilters[filter.field] = { op: filter.op, value: filter.value }
  }
  return {
    sourceFieldCount,
    model: {
      filters: definition.pageFields.map((page) => ({
        fieldIndex: page.field,
        item: memberLabel(page.field, page.item),
      })),
      columns: definition.colFields.filter((field) => field >= 0),
      rows: [...definition.rowFields],
      values,
      groupings,
      labelFilters,
    },
  }
}

/** Fields-pane validation messages, shared by the desktop and VS Code panes. */
export const PIVOT_MODEL_ERRORS: Record<PivotModelError, StringKey> = {
  needRow: 'appNeedRowField',
  needValue: 'dlgPivotErrNeedValue',
  colNeedsSingleValue: 'dlgPivotErrColNeedsSingleValue',
  calcIncomplete: 'dlgPivotErrCalcFieldIncomplete',
  rangeStep: 'dlgPivotErrRangeStep',
  labelFilterText: 'dlgPivotErrLabelFilterText',
  oneValueFilter: 'dlgPivotErrOneValueFilter',
  filterConflict: 'dlgPivotErrFilterConflict',
  dupValueName: 'dlgPivotErrDupValueName',
}
