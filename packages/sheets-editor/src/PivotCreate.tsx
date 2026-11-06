import { useEffect, useRef, useState } from 'react'
import { normalizeLang } from '@genoffice/i18n'
import { LocaleScope, tFor } from '../../../apps/sheets/src/renderer/i18n/locale'
import { PivotDialog } from '../../../apps/sheets/src/renderer/PivotDialog'
import { PivotFieldPane } from '../../../apps/sheets/src/renderer/PivotFieldPane'
import {
  validatePivotModel,
  type PivotLayoutModel,
} from '../../../apps/sheets/src/renderer/pivot-field-model'
import { PIVOT_MODEL_ERRORS } from '../../../apps/sheets/src/renderer/pivot-op-parts'
import type { WorkbookRuntime } from './runtime'

type Seed = Awaited<ReturnType<WorkbookRuntime['pivotEditSeed']>>
interface PaneState {
  readonly seed: Seed
  readonly model: PivotLayoutModel
  readonly deferred: boolean
  readonly dirty: boolean
  readonly error: string | null
}

/**
 * The desktop Create PivotTable dialog and PivotTable Fields pane, bound to
 * host transactions. The pane follows the cursor into PivotTables, closes when
 * it leaves (unless deferred edits would be lost) and opens after creation,
 * as in the desktop shell. Applies run in order, one canonical edit each.
 */
export function usePivotCreate(
  runtime: WorkbookRuntime | null,
  locale: string | undefined,
  onError: (message: string) => void,
) {
  const lang = normalizeLang(locale ?? navigator.language)
  const [context, setContext] = useState<Awaited<
    ReturnType<WorkbookRuntime['pivotCreateContext']>
  > | null>(null)
  const [pane, setPane] = useState<PaneState | null>(null)
  const dismissed = useRef<string | null>(null)
  const applying = useRef<Promise<unknown>>(Promise.resolve())
  const fail = (error: unknown) => onError(error instanceof Error ? error.message : String(error))
  const open = () => {
    if (!runtime) {
      onError('The document is still loading.')
      return
    }
    onError('')
    runtime.pivotCreateContext().then(setContext, fail)
  }
  const edit = (quiet = false) => {
    if (!runtime) return
    runtime.pivotEditSeed().then(
      (seed) => {
        dismissed.current = null
        setPane({ seed, model: seed.model, deferred: false, dirty: false, error: null })
      },
      (error: unknown) => {
        if (!quiet) fail(error)
      },
    )
  }
  const apply = (state: PaneState, model: PivotLayoutModel): PaneState => {
    const invalid = validatePivotModel(model)
    if (invalid !== null)
      return { ...state, model, dirty: false, error: tFor(lang, PIVOT_MODEL_ERRORS[invalid]) }
    const seed = state.seed
    applying.current = applying.current.then(async () => {
      const failure = runtime ? await runtime.editPivot(seed, model) : null
      setPane((current) =>
        current && current.seed.key === seed.key ? { ...current, error: failure } : current,
      )
    })
    return { ...state, model, dirty: false, error: null }
  }
  const key = runtime?.pivotSelectionKey ?? null
  useEffect(() => {
    if (key === null) {
      dismissed.current = null
      setPane((current) => (current && !(current.deferred && current.dirty) ? null : current))
      return
    }
    if (pane?.seed.key === key) return
    setPane(null)
    if (dismissed.current !== key) edit(true)
    // Only the pivot identity drives the pane, as in the desktop shell.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [key])
  const dialog = (
    <LocaleScope lang={lang}>
      {context && runtime && (
        <PivotDialog
          hasFields={context.fields.length > 0}
          sourceRange={context.sourceRange}
          onCreate={async (request) => {
            const failure = await runtime.createPivot(request, context)
            if (failure === null) edit(true)
            return failure
          }}
          onClose={() => setContext(null)}
        />
      )}
      {pane && (
        <PivotFieldPane
          fields={pane.seed.fields}
          model={pane.model}
          deferred={pane.deferred}
          dirty={pane.dirty}
          error={pane.error}
          onModelChange={(model) =>
            setPane(
              pane.deferred ? { ...pane, model, dirty: true, error: null } : apply(pane, model),
            )
          }
          onDeferredChange={(deferred) =>
            setPane(
              !deferred && pane.dirty
                ? { ...apply(pane, pane.model), deferred }
                : { ...pane, deferred },
            )
          }
          onUpdate={() => setPane(apply(pane, pane.model))}
          onGetFieldMembers={pane.seed.members}
          onClose={() => {
            dismissed.current = pane.seed.key
            setPane(null)
          }}
        />
      )}
    </LocaleScope>
  )
  return { open, edit: () => edit(), dialog }
}
