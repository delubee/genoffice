# Sheets Data parity ledger

Reference: `apps/sheets/src/renderer/SheetsRibbon.tsx` Data branch. Plugin:
`packages/sheets-editor/src/WorkbookRibbon.tsx` Data panel. The plugin currently
reuses the actual desktop Sort & Filter group inside a partial Data panel; the complete shared Data ribbon is not mounted.

| Desktop group | Control / variants                            | Current plugin evidence and gap                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   |
| ------------- | --------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Pivot         | Create PivotTable, Refresh selected pivot     | Actual shared desktop `PivotTableRibbonGroup` (Data) and Insert PivotTable button open the actual desktop `PivotDialog` (now async with locked pending dismissal). Session `add_pivot`: desktop default placement, shared layout/source formatting, localized subtotal caption, overlap/name/added-sheet checks, baked output in bounded private stages, one revision; native pivot parts written in a final save stage; real-sidecar and native create/history/save/reopen verified. Refresh / Refresh All: file pivots recompute through the shared engine (in-place data or grown layout with occupied-area refusal), one revision, refreshOnLoad and new location saved; native refresh/history/save/reopen verified. Session-created pivots refresh after save (desktop message). Edit PivotTable: actual desktop PivotTable Fields pane (follows the cursor, deferred Update, validation messages) relayouts session pivots and file pivots through `edit_pivot`; file relayouts save as regenerated definitions with refreshOnLoad. Slicers, timelines, PivotChart missing |
| Get Data      | From Text/CSV, Merge Workbooks, Refresh All   | Actual shared desktop `GetDataRibbonGroup`. From Text/CSV: shared picker/charset/sniffing/typed matrix, desktop 50,000-cell budget through a host staged transaction (one revision/history entry); native 8,400-cell history/save/reopen verified. Merge Workbooks: host native picker, private read-only source snapshots, deduped sheets with values (formulas as cached values), styles, merges, custom row heights and column widths as one staged commit; real-sidecar and native dialog/history/save/reopen verified; 50,000-cell total budget. Refresh All: recomputes every file PivotTable as one revision                                                                                                                                                                                                                                                                                                                                                                                                                                                               |
| Sort & Filter | A→Z, Z→A                                      | Actual shared Sort dropdown submits canonical range sorting without an assumed header, matching desktop simple-sort behavior; browser cross-tab and native Custom Sort regressions pass; broader selection/formula/filter/platform variants remain open                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           |
| Sort & Filter | Custom Sort                                   | Shared three-level desktop dialog mounted; frozen canonical multi-key operation; browser selection/retry verified; native two-key history/save/reopen with typed payloads and style movement verified; broader selection/platform acceptance pending                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              |
| Sort & Filter | Filter toggle                                 | Shared desktop Filter tool and Ctrl/Meta+Shift+L remove existing filters or derive new source-backed smart ranges. Region/browser and Linux native shortcut/history/save/reopen/imported-filter checks pass; native header value/condition/sort popups now connected and verified; larger regions and platform/geometry variants remain open                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      |
| Sort & Filter | Clear, Reapply                                | Actual shared Clear/Reapply controls use distinct canonical operations: Clear retains the range, Reapply recomputes visibility from edited values while preserving conditions. Browser stale-read/Save and native AI-edit/history/save/reopen pass; broader formula/filter/protection/platform variants remain open                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               |
| Sort & Filter | Advanced                                      | Actual shared desktop one-column/two-condition AND/OR dialog, frozen host context, canonical custom criteria and original-filter composition. Browser retry/pending/stale/row-visibility and native raw-value/history/save/reopen/imported-condition preservation pass. 2,000-cell limit, broader variants and installed platforms remain open                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    |
| Data Tools    | Text to Columns: comma, semicolon, space, tab | Actual shared desktop four-option dropdown; sparse blanks and non-A/nonzero-row output supported, writes grow used extent without structural insertion, pending reads/commit join Save and duplicate submissions are rejected. Browser delimiter/pending/retry and native acceptance recorded below; current 2,000-cell limit and broader fidelity/platform variants remain open                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  |
| Data Tools    | Flash Fill                                    | Shared inference algorithm and frozen chunked source reads; broader examples/locales/save acceptance remain open                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  |
| Data Tools    | Remove Duplicates                             | Actual shared desktop header-checkbox dialog, frozen selection and atomic comparison/compaction; browser retry/pending/stale/no-op/merge guards and native formula/type/style/history/save/reopen pass. Larger/sparse/date/error/filter/protection/platform variants remain open                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  |
| Data Tools    | Data Validation                               | Plugin editor and OOXML patch path implemented; original-rule projection, comparisons, prompts, error alerts and common list-source resolution added; full desktop rules pane, extended OOXML and parameter/platform acceptance remain open                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       |
| Data Tools    | Consolidate                                   | Position and left-label modes, five desktop aggregation functions; shared matrix algorithms, frozen reads and typed atomic output. Browser retry/stale tests and native VS Code history/save/independent reopen pass. Actual desktop dialog now shared; larger output, broader formula/locale/appearance variants and installed-platform coverage remain open                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     |
| Forecast      | What-if Analysis → Goal Seek                  | Actual shared Forecast/What-If Analysis dropdown and desktop Goal Seek dialog with frozen host solver connected; one final canonical write; native cross-sheet solve/history/save/reopen verified; broader resource/variant/platform acceptance pending                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           |
| Outline       | Group rows, Group columns                     | Actual shared desktop Group/Ungroup dropdowns route to relative row/column grouping using the desktop planner and complete frozen host reads; native nested grouping/history/save/reopen pass; shared desktop gutter connected and browser mouse acceptance passes; native depth-button acceptance passes; exact-group pointer/geometry and broader variants remain open                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          |
| Outline       | Hide detail rows/columns                      | Desktop nested group plans produce atomic canonical visibility and collapsed-summary updates; source/save and browser grid acceptance pass; native menu folding/save acceptance passes, native depth buttons verified; exact-group native pointer variants pending                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| Outline       | Summary below, Summary right                  | Actual shared Group dropdown check marks echo source/journal flags; both summary toggles persist and native history/save/reopen pass; shared desktop gutter uses the saved placement flags; full geometry/variant acceptance remains open                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| Outline       | Ungroup rows, Ungroup columns                 | Level-aware decrement reuses desktop planner; native row/column nested removal/history/save/reopen pass; shared desktop gutter connected; broader selection/geometry variants remain open                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| Outline       | Show detail rows/columns, Clear Outline       | Show Detail retains independently folded children; Clear Outline restores folded detail and clears selected axes or whole-sheet outlines as one transaction. Browser grid/gutter and native menu history/save/reopen verified; native depth-button history verified; exact-group glyph and broader variants pending                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               |
| Outline       | Subtotal                                      | Actual shared desktop dialog/button; one canonical transaction produces inserted group/grand totals, SUBTOTAL formulas and two-level row outline. Browser retry/Save/stale/cancel and native history/save/reopen verified; existing-feature transforms, computed results/caches, large inputs and installed platforms remain open                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 |

This ledger does not accept the Data tab as complete. Next implementation units
are whole workflows: Custom Sort, Goal Seek, Outline/Subtotal and Get Data/Pivot.
Each needs the desktop control/dialog parameters, explicit session/projection/save
semantics, shared history and real UI acceptance. The source inventory separately
tracks exact control locations and dynamic expressions for review.

Goal Seek prerequisite: the desktop now uses a shared, evaluator-driven secant
solver with at most 100 evaluations and the existing 0.001 target tolerance.
Five focused tests cover linear/nonlinear convergence, already-satisfied inputs,
constant/error-valued functions, finite parameter validation, bounded evaluation
and evaluator cancellation. Desktop restoration after failure remains in its
adapter. The plugin must use an isolated frozen host calculation source and
publish only the accepted final input-cell change; the existing desktop adapter's
per-guess grid writes must not be reused for the plugin. The plugin now reuses the desktop dialog with a controlled locale and a host-owned solve/commit controller. Native cross-sheet solve/history/save/reopen passes; broader variants remain open.

Custom Sort prerequisite: the shared `sort_range` operation now accepts an
optional ordered `levels` array (up to three column/direction pairs). When present
it supplies the comparison keys; the original byColumn/order fields remain the
fallback for existing callers. Small-range expansion, session style/type movement
and desktop bulk sorting consume the same keys. Five focused sort tests pass,
including secondary descending keys, stable ties, blank ordering, exact planned
cell rewrites and rejection of an out-of-range key. The actual shared dialog adapter is now connected; native two-key history/save/reopen acceptance now passes; broader variants and platforms remain open.

Column-header popup follow-up: the plugin now uses the desktop Univer filter
preset. Values, blanks, two-condition criteria and column clearing route through
the authoritative session after a complete bounded source read. Browser retry,
pending/Save, stale/offscreen reads and authoritative formula-cache/mode-switch
checks pass. The temporary column/value-textarea editor has been removed. Native
pointer/history/save/reopen acceptance passes in VS Code 1.109.5 with a
1200×900 CDP workbench viewport; imported-popup variants and color persistence,
filters above 2,000 cells and the broader platform/geometry/locale cases remain
open. The Data tab is still incomplete.

Filter-popup sorting is now connected through the same desktop sort preset and
embedded Ascending/Descending buttons. Canonical sorting uses complete host row
visibility and leaves filtered/manual hidden slots stationary; typed values and
styles move together across visible slots. New filter ownership preserves source
manual hides in the existing visibility journal. Browser retry/Save and native
source-hide/history/save/reopen checks pass. Formula movement and the remaining
sort/locale/platform variants are not accepted as complete.

Color-filter persistence follow-up (2026-10-29) supersedes the earlier color
persistence gap for the current bounded base-style workflow. The actual By Colors
panel now supports fill/font, multiple choices, no fill and default black through
canonical criteria. Source styles and other-column criteria determine hidden
rows; Reapply responds to style edits. Single colors use standard DXF references,
while all selections also have a GenOffice extension for exact desktop round
trips. Real sidecar import/save and native VS Code history/imported editing/clear
acceptance pass. External Excel handling of multicolor extensions, rich text,
conditional/table/theme/pattern cases, pending-formula UI, larger filters and
installed platforms remain open.

Data Tools controls follow-up (2026-10-29): the desktop and plugin now mount
the same `DataToolsRibbonGroup` (Text to Columns, Flash Fill, Remove Duplicates,
Data Validation and Consolidate), retaining desktop labels, icons, descriptions
and the four delimiter menu options. This shares the control surface, not all
underlying dialogs: Data Validation still uses a plugin editor; Consolidate now reuses its actual desktop dialog.
The plugin's additional Clear Validation/Clear Contents remain outside this
shared group. Forecast and the Group/Ungroup menus are now shared, as described
below. Subtotal/Get Data are not accepted as complete desktop control reuse.

Text to Columns reads complete bounded host rectangles; omitted cells in a sparse
response are blank. Output grows through cell writes instead of inserting columns
at the used extent, which would shift unrelated formulas. It retains the native
split command's inclusive trailing clear column and numeric/blank intermediate-row
clearing semantics; desktop command row-offset quirks are not copied. Reads and
canonical commit share one pending-work slot, rejected duplicate calls create no
second operation, and Save waits for it. Full Excel wizard behavior, larger
selections, rich text, formula/protection/merge variants and Windows remain open.

Forecast/Outline shared-control follow-up (2026-10-29): desktop and plugin now
mount `ForecastRibbonGroup` and `OutlineRibbonGroup`. The former exposes the
original What-If Analysis → Goal Seek entry; the latter exposes both original
Group/Ungroup dropdowns with row/column, hide/show detail, summary-direction and
Clear Outline options. The desktop appends its Subtotal button; the plugin still
needs the Subtotal workflow, so the full Outline group remains incomplete.

Summary marks derive from the canonical source/journal and update only after an
accepted transaction. Save now retains promises that were active when it began,
even if cell-editor flushing awaits a failed recoverable transaction and clears
its work slot. A shared menu-positioning correction constrains dropdowns to the
actual webview viewport, uses the side with available space, and allows scrolling
for entries that cannot all fit above the VS Code status bar. Broader menu/locale/
appearance/lifecycle, installed platforms and all missing Data workflows remain
open.

Subtotal follow-up (2026-10-29): `OutlineRibbonGroup` now includes its original
Subtotal button in both hosts. The actual desktop dialog is shared, retains all
three selectors and Sum/Count/Average, and supports an async frozen plugin
transaction with failed-draft retention. Both implementations use the same pure
adjacent-group planner. Count uses SUBTOTAL function 3 (COUNTA), matching desktop.
Group labels and Grand Total retain desktop literal wording.

The plugin sends one standalone `subtotal` operation; the candidate performs
private structural then content/outline stages and returns only the final root.
Original-root inverse history stays exact. Host materialization covers every
sheet (up to the existing 200,000-cell per-sheet structural budget), and the
selected source retains a 2,000-cell limit and desktop 200-group limit. Private
content stages are bounded to preserve existing operation/write ceilings.
Current generic structural feature guards still apply; existing source outlines
are explicitly refused. This is not full imported-feature or large-file parity.
Subtotal is now included in native session tool schemas; standalone CLI/legacy
desktop AI/generic in-memory execution require an authoritative session and
reject it explicitly. Normal desktop ribbon subtotal remains supported.

Shared Consolidate dialog follow-up (2026-10-29): both hosts now mount the actual
`apps/sheets/src/renderer/ConsolidateDialog.tsx`. The plugin-only replica was
removed. Function dropdown, reference Add/Enter/Delete/deduplication, pending
unadded reference on OK, left-label mode, target descriptions and Max/Min reset
to Sum in label mode all come from one component. The shared component supports
a controlled locale and async host submission, with locked pending dismissal,
failed-draft retention and synchronous duplicate-submit protection.

Sparse omitted source labels are now blank and are skipped exactly as in the
desktop matrix builder. Pending/error label formulas and ready formulas without
an authoritative cache are refused before output. Existing frozen root/revision,
source reads, overlap/bounds/merge checks, canonical write and Save barrier remain
in use. Four browser tests and the real VS Code shared dropdown/history/save/
independent-reopen test pass. The 2,000-cell output, 50-reference and 200,000-row
label limits remain; complete formula-cache, geometry/locales/appearance and
installed Linux/Windows acceptance remain open.

Get Data follow-up (2026-10-29): desktop and plugin now mount the same
`GetDataRibbonGroup`. From Text/CSV shares the desktop picker, 32 MB byte
budget, locale charset preference, delimiter sniffing and numeric typing through
`apps/sheets/src/renderer/csv-import-model.ts`; the plugin binds the write to
the selection and revision current after the pick and submits one canonical
transaction. Outputs above the session's 2,000-cell ordinary batch are refused
explicitly; desktop accepts 50,000. Lifting that needs a manual path through the
session import pipeline (or a raised, journaled bulk-write budget). Refresh All
matches the desktop message for workbooks without pivots but cannot refresh
existing pivots until the Pivot unit lands. Merge Workbooks reports that it is
not available. Neither is accepted as parity.

Staged bulk transactions (2026-10-29): `packages/workbook-session/src/staged-candidate.ts`
lets the host commit a list of ordinary bounded stages (each within the 2,000-cell
batch and the structural/content separation) as one private candidate, one
revision and one history entry, up to 50,000 written cells. Stages may target a
sheet added by an earlier stage through a `new-sheet:<name>` placeholder. It is
reachable only from the manual webview bridge (`view.stagedIntent`) and from
host-built workflows; model-facing tools keep the ordinary 2,000-cell limit.
CSV import above 2,000 cells and Merge Workbooks use it. Merge Workbooks runs on
the host (`workbook.mergeWorkbooks`): VS Code's open dialog, private snapshot
copies of the sources, shared `dedupeSheetName`, then one staged commit. Unlike
the desktop it is a single undo step. Its 50,000-cell budget counts each source
sheet's full used rectangle; the desktop has no total cap. Source formulas
without a stored result import as blanks, as on the desktop.

PivotTable create follow-up (2026-10-29): `add_pivot` is a standalone session
operation. The candidate reads the source from hydrated material (ready formula
results only), applies the desktop source-label rule (`pivot-placement.ts`),
builds the layout with `buildPivotLayout`, writes the report-filter block and
body through bounded private stages and records a `pivotAdditions` sheet
feature. The package writer emits those as native pivot parts in a final stage
after all structural and sheet stages. While a session pivot is pending,
row/column insert/delete and rename/delete of its sheets are refused, because
the written definition pins final coordinates and names. Existing target
content is overwritten without a prompt, as on the desktop (Excel asks first).
`add_pivot` is now also exposed through the regenerated native session tools.
Open: Refresh/Refresh All recompute, Edit PivotTable and the Fields pane,
slicers, timelines, PivotChart, file-pivot editing and installed platforms.

PivotTable refresh follow-up (2026-10-29): file pivot definitions are parsed
from the private snapshot when a workbook opens (`readArchiveText` port) and
ride in each sheet's original metadata. `refresh_pivot` (one op per sheet;
Refresh All batches them) reads complete sheet material, applies the shared
source-label rule, `growPivotDefinition`/`recomputePivotData`, and writes the
data area or the grown block (`grownPivotBlock`, now shared with the desktop's
`applyGrownPivotOutput`). Growth into occupied cells, layout drift, missing
sources and unsupported definitions are refused with the desktop messages.
Grown definitions persist in the session for later refreshes. Save sets
refreshOnLoad on refreshed caches and widens locations in the same stage as the
cell edits (the writer checks growth against its input).

Edit PivotTable follow-up (2026-10-29): the plugin mounts the actual desktop
`PivotFieldPane` (its `@genoffice/ui` import now uses the chrome-only subpath,
keeping the AI model picker out of the webview). A host-side controller mirrors
the desktop shell: the pane follows the cursor into session and file pivots,
closes when it leaves unless deferred edits are pending, honours dismissal,
opens after Create, and applies each change as one ordered `edit_pivot`
transaction. The candidate shares the add planner, rewrites the union of the old
and new blocks (shrinking clears leftovers), refuses growth into occupied cells
and excludes the pivot itself from overlap checks. Session pivots replace their
pending addition; file pivots keep a relayout payload that saves through the
existing PivotRefreshUpdate.relayout path, and refresh waits for the save, as on
the desktop. Model builders and pane error keys moved to the shared pure
`pivot-op-parts.ts`.
