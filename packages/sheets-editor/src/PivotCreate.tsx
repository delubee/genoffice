import { useState } from 'react'
import { normalizeLang } from '@genoffice/i18n'
import { LocaleScope } from '../../../apps/sheets/src/renderer/i18n/locale'
import { PivotDialog } from '../../../apps/sheets/src/renderer/PivotDialog'
import type { WorkbookRuntime } from './runtime'

/** The desktop Create PivotTable dialog, bound to a frozen host source read. */
export function usePivotCreate(
  runtime: WorkbookRuntime | null,
  locale: string | undefined,
  onError: (message: string) => void,
) {
  const [context, setContext] = useState<Awaited<
    ReturnType<WorkbookRuntime['pivotCreateContext']>
  > | null>(null)
  const open = () => {
    if (!runtime) {
      onError('The document is still loading.')
      return
    }
    onError('')
    runtime
      .pivotCreateContext()
      .then(setContext, (error: unknown) =>
        onError(error instanceof Error ? error.message : String(error)),
      )
  }
  const dialog =
    context && runtime ? (
      <LocaleScope lang={normalizeLang(locale ?? navigator.language)}>
        <PivotDialog
          hasFields={context.fields.length > 0}
          sourceRange={context.sourceRange}
          onCreate={(request) => runtime.createPivot(request, context)}
          onClose={() => setContext(null)}
        />
      </LocaleScope>
    ) : null
  return { open, dialog }
}
