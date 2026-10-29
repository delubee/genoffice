import { VisualDialog } from '../../../apps/sheets/src/renderer/VisualDialogs'
import { ChartTextDialog } from '../../../apps/sheets/src/renderer/ChartTextDialog'
import {
  chartSupportsSeriesReplace,
  applyChartStateEdit,
} from '@genoffice/xlsx-gateway/domain/chart-visual'
import { useRef, useState } from 'react'
import { normalizeLang } from '@genoffice/i18n'
import { ChartFormatPane, SelectDataDialog } from '../../../apps/sheets/src/renderer/ChartPanels'
import { LocaleScope, useI18n } from '../../../apps/sheets/src/renderer/i18n/locale'
import type { ChartEditData } from '../../../apps/sheets/src/renderer/WorkbookVisuals'
import type { WorkbookRuntime } from './runtime'

export function DesktopChartPanels({
  runtime,
  locale,
}: {
  runtime: WorkbookRuntime | null
  locale?: string | undefined
}) {
  const target = runtime?.chartFormatTarget
  const selection = runtime?.chartSelectionTarget
  const text = runtime?.chartTextTarget
  const visual = runtime?.visualDialogTarget
  return (target || selection || text || visual) && runtime ? (
    <LocaleScope lang={normalizeLang(locale ?? navigator.language)}>
      {visual && (
        <VisualDialog
          key={`${visual.visual.id}:${visual.stateId}:${visual.kind}`}
          kind={visual.kind}
          visual={visual.visual}
          frame={visual.frame}
          onApply={async (changes) => {
            await runtime.editVisualDialog(changes)
          }}
          onClose={() => runtime.closeVisualDialog()}
        />
      )}
      {text && (
        <ChartTextDialog
          key={`${text.key}:${text.stateId}:${text.target}`}
          target={text.target}
          initial={text.initial}
          onCommand={(command) => runtime.editChartText(command)}
          onClose={() => runtime.closeChartText()}
        />
      )}
      {target && <FormatAdapter key={target.visual.id} runtime={runtime} target={target} />}
      {selection && (
        <SelectDataDialog
          key={`${selection.key}:${selection.stateId}`}
          chart={selection.visual.chart}
          supported={chartSupportsSeriesReplace(selection.visual.chart.chartTypes)}
          onApply={() => {
            throw new Error('chart_reference_required')
          }}
          onSelectionApply={(intent) => runtime.editChartSelection(intent)}
          onClose={() => runtime.closeChartSelection()}
        />
      )}
    </LocaleScope>
  ) : null
}

function FormatAdapter({
  runtime,
  target,
}: {
  runtime: WorkbookRuntime
  target: NonNullable<WorkbookRuntime['chartFormatTarget']>
}) {
  const { t } = useI18n()
  const [pending, setPending] = useState(false)
  const [error, setError] = useState('')
  const [proposal, setProposal] = useState<ChartEditData | null>(null)
  const pendingRef = useRef(false)
  const edit = (changes: ChartEditData) => {
    if (pendingRef.current) return
    pendingRef.current = true
    setPending(true)
    setProposal(changes)
    setError('')
    void runtime
      .editChartFormat(changes)
      .catch((reason: unknown) => {
        const message = reason instanceof Error ? reason.message : String(reason)
        setError(
          message.includes('invalid_chart_axis_bounds') ? t('appChartAxisRangeInvalid') : message,
        )
      })
      .finally(() => {
        pendingRef.current = false
        setPending(false)
        setProposal(null)
      })
  }
  const current = runtime.chartFormatTarget ?? target
  // Text/number inputs already hold their draft. Do not replace their keyed
  // DOM while pending, so a rejected write preserves the user's correction.
  const preview = proposal ? { ...proposal } : undefined
  if (preview) {
    delete preview.title
    delete preview.valueAxis
    delete preview.dataLabelFormat
  }
  return (
    <ChartFormatPane
      chart={applyChartStateEdit(current.visual.chart, preview)}
      element={current.element}
      pending={pending}
      error={error}
      onEdit={edit}
      onClose={() => {
        if (!pendingRef.current) runtime.closeChartFormat()
      }}
    />
  )
}
