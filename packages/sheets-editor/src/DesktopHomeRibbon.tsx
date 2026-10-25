import { useState } from 'react'
import { normalizeLang } from '@genoffice/i18n'
import { SheetsRibbon } from '../../../apps/sheets/src/renderer/SheetsRibbon'
import { FormatCellsDialog } from '../../../apps/sheets/src/renderer/FormatCellsDialog'
import { LocaleScope, setModuleLang } from '../../../apps/sheets/src/renderer/i18n/locale'
import type { WorkbookRuntime } from './runtime'
import type { RibbonTarget } from './ribbon-intents'

/** The desktop Home command band, backed by the plugin's authoritative session. */
export function DesktopHomeRibbon({
  runtime,
  locale,
  onError,
}: {
  runtime: WorkbookRuntime | null
  locale?: string
  onError(message: string): void
}) {
  const lang = normalizeLang(locale ?? navigator.language)
  setModuleLang(lang)
  const [dialog, setDialog] = useState<{
    target: RibbonTarget
    stateId: string
    revision: number
  } | null>(null)
  const rawAnchor = dialog
    ? runtime?.projectedValues(dialog.target.address.split(':')[0]!)?.[0]?.[0]
    : null
  const anchorValue =
    typeof rawAnchor === 'number' || typeof rawAnchor === 'string' ? rawAnchor : null
  const run = async (work: () => Promise<unknown>) => {
    onError('')
    try {
      await work()
      return true
    } catch (error) {
      onError(String(error))
      return false
    }
  }
  const command = (value: string) => {
    if (!runtime) {
      onError('The document is still loading.')
      return
    }
    if (value === 'format-cells') {
      void run(async () => {
        if (!(await runtime.flushEditor()))
          throw new Error('Finish editing before opening Format Cells.')
        const target = runtime.ribbonTarget
        if (!target) throw new Error('Select cells before opening Format Cells.')
        const { stateId, revision } = runtime.selectionContext
        setDialog({ target, stateId, revision })
      })
      return
    }
    if (/^autofn:(SUM|AVERAGE|COUNT|MIN|MAX)$/.test(value)) {
      void run(() => runtime.ribbonSummary(value.slice(7) as 'SUM'))
      return
    }
    void run(() => runtime.ribbonCommands([value]))
  }
  const unused = (): never => {
    throw new Error('A non-Home ribbon callback was invoked in the Home adapter.')
  }
  const view = runtime?.projectedView()
  return (
    <LocaleScope lang={lang}>
      <header className="excel-header workbook-desktop-header" inert={!runtime}>
        <SheetsRibbon
          activeTab="Home"
          selectionFormat={runtime?.ribbonTarget?.format ?? null}
          formatPainterActive={false}
          sheetHasContent={runtime?.ribbonEcho.hasContent ?? false}
          sheetProtected={runtime?.ribbonEcho.protected ?? null}
          workbookProtected={null}
          formulaBarVisible={true}
          crossHighlightVisible={false}
          pageLayout={{
            showGridlines: view?.gridlines ?? true,
            showHeadings: view?.headings ?? true,
            rightToLeft: false,
          }}
          selectedChart={null}
          selectedShape={null}
          selectedTable={null}
          outlineSummary={{ below: true, right: true }}
          onCommand={command}
          onListNames={unused}
          calcManual={false}
          onAiRun={(instruction) => {
            if (runtime) void run(() => runtime.ribbonChat(instruction))
          }}
          aiOpen={false}
          onAiToggle={() => {
            if (runtime) void run(() => runtime.ribbonChat())
          }}
          onRefreshPivot={unused}
          onIsSelectionInPivot={unused}
        />
      </header>
      {dialog && runtime && (
        <FormatCellsDialog
          selectionFormat={dialog.target.format}
          anchorValue={anchorValue}
          onCommand={command}
          onClose={() => setDialog(null)}
          onApply={(commands) => run(() => runtime.ribbonCommands(commands, dialog))}
        />
      )}
    </LocaleScope>
  )
}
