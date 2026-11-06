import { useRef, useState } from 'react'
import { normalizeLang } from '@genoffice/i18n'
import {
  timelineDomainOf,
  timelineSelection,
  type MonthKey,
} from '@genoffice/xlsx-gateway/domain/pivot-timeline'
import type { WorkbookPivotDefinition } from '@genoffice/xlsx-gateway/shared/workbook-protocol'
import { LocaleScope, tFor } from '../../../apps/sheets/src/renderer/i18n/locale'
import {
  SlicerFieldPicker,
  SlicerPanels,
  type SlicerMember,
  type SlicerUiState,
} from '../../../apps/sheets/src/renderer/SlicerPanel'
import {
  TimelineFieldPicker,
  TimelinePanels,
  type TimelineUiState,
} from '../../../apps/sheets/src/renderer/TimelinePanel'
import type { WorkbookRuntime } from './runtime'

interface Picker {
  readonly kind: 'slicer' | 'timeline'
  readonly sheetId: string
  readonly pivotPath: string
  readonly fields: readonly { field: number; name: string }[]
}

/**
 * The desktop Insert → Slicer / Timeline flow: a field picker, then floating
 * panels whose selections filter the PivotTable through host transactions.
 * Panels are session UI, as on the desktop; the filtered output and hidden
 * items persist through the pivot refresh write-back.
 */
export function usePivotSlicers(
  runtime: WorkbookRuntime | null,
  locale: string | undefined,
  onError: (message: string) => void,
) {
  const lang = normalizeLang(locale ?? navigator.language)
  const text = (key: Parameters<typeof tFor>[1], params?: Record<string, string | number>) =>
    tFor(lang, key, params)
  const [slicers, setSlicers] = useState<readonly SlicerUiState[]>([])
  const [timelines, setTimelines] = useState<readonly TimelineUiState[]>([])
  const [picker, setPicker] = useState<Picker | null>(null)
  const [notice, setNotice] = useState<string | null>(null)
  const applying = useRef<Promise<unknown>>(Promise.resolve())
  const fieldName = (definition: WorkbookPivotDefinition, field: number) =>
    definition.fields[field]?.name || text('appFieldN', { n: field + 1 })
  /// One filter panel per field: slicers and timelines share the hidden items.
  const taken = (pivotPath: string) =>
    new Set(
      [...slicers, ...timelines]
        .filter((entry) => entry.pivotPath === pivotPath)
        .map((entry) => entry.field),
    )
  const open = (kind: Picker['kind']) => {
    setNotice(null)
    if (!runtime) {
      onError(text('appSlicerNeedsFile'))
      return
    }
    let found: ReturnType<WorkbookRuntime['filePivotAtSelection']>
    try {
      found = runtime.filePivotAtSelection()
    } catch (error) {
      onError(error instanceof Error ? error.message : String(error))
      return
    }
    onError('')
    const { definition } = found
    const seen = new Set<number>()
    const fields: { field: number; name: string }[] = []
    for (const field of [
      ...definition.rowFields,
      ...definition.colFields,
      ...definition.pageFields.map((page) => page.field),
    ]) {
      if (field < 0 || seen.has(field)) continue
      seen.add(field)
      const meta = definition.fields[field]
      if (kind === 'timeline') {
        if (!meta || meta.grouping !== undefined) continue
        if (timelineDomainOf(meta.sharedItems, definition.fieldItems[field] ?? []) === null)
          continue
      }
      fields.push({ field, name: fieldName(definition, field) })
    }
    if (!fields.length) {
      onError(text(kind === 'slicer' ? 'appPivotNoSlicerFields' : 'appTimelineNoDateFields'))
      return
    }
    const occupied = taken(found.pivotPath)
    const available = fields.filter((entry) => !occupied.has(entry.field))
    if (!available.length) {
      onError(text('appFieldFilterTaken'))
      return
    }
    setPicker({ kind, sheetId: found.sheetId, pivotPath: found.pivotPath, fields: available })
  }
  const create = (field: number): string | null => {
    const definition = picker && runtime?.filePivotDefinition(picker.sheetId, picker.pivotPath)
    if (!picker) return text('appSlicerPivotStale')
    if (!definition) return text('appPivotDefNotLoaded')
    const base = { sheetId: picker.sheetId, pivotPath: picker.pivotPath, field }
    const name = fieldName(definition, field)
    if (picker.kind === 'timeline') {
      const domain = timelineDomainOf(
        definition.fields[field]?.sharedItems ?? [],
        definition.fieldItems[field] ?? [],
      )
      if (!domain) return text('appTimelineNoDateFields')
      setTimelines((current) => [
        ...current,
        {
          ...base,
          id: `timeline-${Date.now().toString(36)}-${current.length + 1}`,
          fieldName: name,
          members: domain.members,
          minMonth: domain.minMonth,
          maxMonth: domain.maxMonth,
          range: null,
        },
      ])
      setNotice(text('appTimelineCreated', { name }))
      return null
    }
    const members: SlicerMember[] = []
    const selected: number[] = []
    definition.fieldItems[field]?.forEach((item, member) => {
      if (item.x === null) return
      const shared = definition.fields[field]?.sharedItems[item.x]
      members.push({
        member,
        label:
          shared === null || shared === undefined || shared === ''
            ? text('appBlank')
            : String(shared),
      })
      if (!item.hidden) selected.push(member)
    })
    if (!members.length) return text('appFieldNoMembers')
    setSlicers((current) => [
      ...current,
      {
        ...base,
        id: `slicer-${Date.now().toString(36)}-${current.length + 1}`,
        fieldName: name,
        members,
        selected,
      },
    ])
    setNotice(text('appSlicerCreated', { name }))
    return null
  }
  /// Applies in order; the panel state follows only a successful transaction.
  const apply = (
    target: { sheetId: string; pivotPath: string; field: number },
    selected: readonly number[] | null,
    done: () => void,
  ) => {
    setNotice(null)
    applying.current = applying.current.then(async () => {
      const failure = runtime
        ? await runtime.filterPivot(target, selected)
        : text('appOpenXlsxFirst')
      if (failure !== null) {
        onError(failure)
        return
      }
      onError('')
      done()
    })
  }
  const toggle = (slicerId: string, member: number) => {
    const slicer = slicers.find((entry) => entry.id === slicerId)
    if (!slicer) return
    const next = slicer.selected.includes(member)
      ? slicer.selected.filter((entry) => entry !== member)
      : [...slicer.selected, member].sort((a, b) => a - b)
    if (!next.length) {
      onError(text('appSlicerKeepOne'))
      return
    }
    apply(slicer, next, () => {
      setSlicers((current) =>
        current.map((entry) => (entry.id === slicerId ? { ...entry, selected: next } : entry)),
      )
      setNotice(text('appSlicerApplied', { name: slicer.fieldName }))
    })
  }
  const selectAll = (slicerId: string) => {
    const slicer = slicers.find((entry) => entry.id === slicerId)
    if (!slicer) return
    apply(slicer, null, () => {
      setSlicers((current) =>
        current.map((entry) =>
          entry.id === slicerId
            ? { ...entry, selected: entry.members.map((item) => item.member) }
            : entry,
        ),
      )
      setNotice(text('appSlicerCleared', { name: slicer.fieldName }))
    })
  }
  const removeSlicer = (slicerId: string) => {
    const slicer = slicers.find((entry) => entry.id === slicerId)
    if (!slicer) return
    apply(slicer, null, () => {
      setSlicers((current) => current.filter((entry) => entry.id !== slicerId))
      setNotice(text('appSlicerRemoved', { name: slicer.fieldName }))
    })
  }
  const range = (
    timelineId: string,
    value: { readonly start: MonthKey; readonly end: MonthKey } | null,
  ) => {
    const timeline = timelines.find((entry) => entry.id === timelineId)
    if (!timeline) return
    const selected = timelineSelection(
      { members: timeline.members, minMonth: timeline.minMonth, maxMonth: timeline.maxMonth },
      value,
    )
    if (selected !== null && !selected.length) {
      onError(text('appTimelineEmptyRange'))
      return
    }
    apply(timeline, selected, () => {
      setTimelines((current) =>
        current.map((entry) => (entry.id === timelineId ? { ...entry, range: value } : entry)),
      )
      setNotice(
        text(value === null ? 'appTimelineCleared' : 'appTimelineApplied', {
          name: timeline.fieldName,
        }),
      )
    })
  }
  const removeTimeline = (timelineId: string) => {
    const timeline = timelines.find((entry) => entry.id === timelineId)
    if (!timeline) return
    apply(timeline, null, () => {
      setTimelines((current) => current.filter((entry) => entry.id !== timelineId))
      setNotice(text('appTimelineRemoved', { name: timeline.fieldName }))
    })
  }
  const panels = (
    <LocaleScope lang={lang}>
      {picker?.kind === 'slicer' && (
        <SlicerFieldPicker
          fields={picker.fields}
          onPick={(field) => {
            const failure = create(field)
            if (failure !== null) onError(failure)
            return failure
          }}
          onClose={() => setPicker(null)}
        />
      )}
      {picker?.kind === 'timeline' && (
        <TimelineFieldPicker
          fields={picker.fields}
          onPick={(field) => {
            const failure = create(field)
            if (failure !== null) onError(failure)
            return failure
          }}
          onClose={() => setPicker(null)}
        />
      )}
      <SlicerPanels
        slicers={slicers}
        onToggle={toggle}
        onSelectAll={selectAll}
        onRemove={removeSlicer}
      />
      <TimelinePanels
        timelines={timelines}
        onRange={(timelineId, start, end) => range(timelineId, { start, end })}
        onClear={(timelineId) => range(timelineId, null)}
        onRemove={removeTimeline}
      />
      {notice && (
        <p className="dialog-note" role="status">
          {notice}
        </p>
      )}
    </LocaleScope>
  )
  return { open, panels }
}
