# Implementation approach

1. Separate browser presentation from desktop application orchestration. Extract
   the exact Sheets command band, retaining the desktop shell as a consumer.
   Expose office chrome independently of the AI UI barrel. Provide controlled
   locale scopes. Keep command state and callbacks explicit.
2. Inventory each standard/contextual tab, dropdown option, dialog and pane in
   Sheets and Docs. Map controls to desktop command helpers and available session
   operations. Treat missing operations as work to implement, not as reasons to
   remove controls from the goal.
3. Adapt the shared Sheets command band to `WorkbookRuntime`. Extend validated
   operations and projection where necessary, and reuse dialog models. All
   mutation paths must acknowledge host acceptance before being considered saved.
4. Extract privileged ports from the Docs ribbon and its transitive helpers.
   Reuse the same ribbon and dialogs with the DOCX runtime. Extend the policy,
   conversion, session validation, native operations and serializer together for
   rich features; remove initial slice restrictions only with save/history proof.
5. Route native chat, file actions, clipboard, image pickers, export/print and
   window actions through typed host ports. Preserve URI-based resource handling
   and the Windows junction fallback. Keep desktop behavior intact.
6. Share the appropriate desktop chrome styles with both embedded appearances.
   Verify compact/collapsed layouts, narrow panes, contextual transitions,
   popover layering, focus, shortcuts and locale changes in actual browsers.
7. Verify representative actions in each command family end to end; keep a
   complete coverage ledger for all controls. Build and run installed VSIX
   packages on Linux and Windows before final acceptance.

## First implementation evidence

- `SheetsRibbon.tsx` now contains the original desktop command band; the desktop
  shell consumes it. No second menu definition was introduced for VS Code.
- `@genoffice/ui/chrome` supplies browser-only office controls without importing
  the AI UI barrel.
- `paste-special-model.ts` isolates options and availability from Univer's
  clipboard implementation; both the dialog and command band consume the model.
- Docs and Sheets expose `LocaleScope` for controlled embedded rendering.
- The Docs ribbon, font/paragraph dialogs, color palette, insertion/reference
  tabs and wheel zoom now import the browser-only chrome entry. Both desktop
  ribbon dependency graphs pass the embedded browser boundary.
- Embedded component tests exercise all 11 Sheets tab bodies, formatting/menu
  callback parameters, view-state echo, language changes and Escape dismissal.
  A separate browser bundle test checks the transitive dependency boundary.

These are preparation for host integration. They do not prove mutations or save
behavior from the new shared ribbon inside the production extension.

## DOCX dialog integration

The embedded DOCX editor now opens the same Font and Paragraph components used
by the desktop ribbon. It imports the desktop Docs stylesheet and shared dropdown
styles, with native host theme tokens mapped for the VS Code appearance. The
document page retains its authored colors. Dialog labels use the controlled
desktop locale scope.

The document policy explicitly validates the added font effects, advanced
character settings, paragraph spacing/indents and pagination flags. Existing
source mark values remain preservable when outside editable dialog bounds;
newly edited values must pass validation. Original mark values are indexed once
per source block during validation, avoiding repeated subtree scans.

Authoritative resynchronization now restores the ProseMirror selection and emits
a transaction notification so selection-dependent controls update before they
receive focus. This fixes font input remounting after history/resync.

These dialogs are currently launched from the existing simplified toolbar. The
complete desktop ribbon, other dialogs and all XLSX adapters remain open work.

## Native file and chat menu ports

Both embedded editors now expose a shared five-entry file/chat action group.
The browser sends only the typed intent (`open`, `save`, `saveAs`, or `chat` with
document/selection scope). A strict discriminated schema and bridge method
allowlist reject extra URIs, command identifiers and unsupported payloads.

The extension captures the resource/session and panel when registering the
handler. It synchronizes pending edits, checks the captured session and active
panel again, and prevents overlapping menu actions. Open uses the native picker
and dispatches supported files to the appropriate GenOffice custom editor.
Save/Save As use VS Code's native commands and existing custom-document lifecycle.

Chat actions capture a native tool binding before moving focus. Workbook
bindings reuse an extracted selection-capture helper without touching the
clipboard. DOCX bindings use the existing context implementation. The prompt
contains resource/binding/revision metadata, not document contents. Generic chat
opens a partial request in native Agent mode; nonempty ribbon presets can submit
their explicit instruction through the same port. Model configuration, skills,
terminal/Python capabilities and tool consent stay with VS Code.

These entry points prepare the callbacks for the full shared ribbons. Other
desktop file/media/print/window capabilities and every preset's actual ribbon
connection still need implementation and verification.

## Shared chart presentation and command ownership

Desktop chart drawing is now a browser-only SheetChart component. Desktop keeps
media retrieval, native visual installation, source hydration and inline/context
edit services outside that presentation module. Both desktop manual insertion
and session authoring share the same chart construction model; the host retains
source access and transaction authority. All nine ordinary Insert types serialize
through the existing drawing writer. Recommended Charts keeps its actual desktop
presentation with asynchronous hosted apply and state-bound source reads.

The plugin uses desktop in-grid visual installation and inline chart properties.
Original charts display from immutable metadata plus validated property/geometry
overlays without becoming new additions. Full contextual controls, live refresh of all original references and broad
formula freshness remain open; bounded range editing and missing-cache hydration
are connected through the vector port described below.

Native ribbon focus belongs to the selected tab until the user leaves it.
Univer's hidden normal-editor selection/position synchronization must not override
that ownership. A per-render-instance focus boundary guards that one entry point,
restores it on renderer/runtime disposal and releases on a real pointer action.
Explicit Go To focus restoration releases tab ownership before focusing the grid.
This does not disable the engine's global keyboard or other editor focus paths.

## Floating chart frame integration

The embedded workbook now installs chart visuals through the original desktop
installWorkbookVisuals implementation and the free drawing preset. Browser chrome
exports the existing shape-preview path helper so this module no longer imports
the AI UI barrel. Workbook media reads have a narrow optional reader port with a
capability-checked desktop fallback; the chart-only embedded installer performs
no media reads. Duplicate side preview drawings are removed; temporary chart
property controls remain collapsed until the actual chart editor/panes are wired.

set_visual_anchor is a shared validated operation. Session-created charts update
their visual state directly; original drawings retain immutable metadata and store
an anchor/removal patch keyed by visual id. The host resolves drawing paths and
anchor indices from the original logical sheet mapping, validates the locator and
uses existing surgical save patches. Original-chart deletion uses the existing
cascading drawing writer, including optional deletion of an empty drawing part.
Desktop operation execution routes the new operation through its existing shape
edit port. Raw package paths are never accepted from the browser operation.

Projection installation waits while a drag is active. Visual transactions bind to
the installed state/revision and reject overtaken gestures. Editor flushing waits
for the gesture and authoritative submission. Escape cancellation runs in the
window capture phase so engine focus changes cannot hide it from the frame;
pointer-capture loss and window blur also cancel. The native Webview stylesheet
clears VS Code's default body padding, keeping canvas viewport measurements inside
the visible editor. Full contextual menus, all-source live-cache acceptance
and cross-platform/pointer/geometry audits remain open; basic original-file
properties are now connected through the host overlays described below.

The plugin uses the desktop inline chart property editor for session-added and
original-file charts. Original edits resolve a unique trusted chart part and owner
in the host, enforce sheet protection and removal state, and store validated
property overlays separately from immutable source metadata. The package writer
replays these overlays through surgical chart XML patches, omitting overlays for
deleted visuals. Hosted apply is asynchronous: it keeps failed drafts,
disables controls while pending, and releases the float installation guard during
submission so the committed projection can replace its own editor. Title, supported
conversions, series colors and point colors use the existing canonical edit_chart
operation. Source data ranges use authoritative seriesData operations and the
vector reader described below; the contextual menu remains open work. The package writer must forward both authored series colors and
point-color maps rather than silently rebuilding the default palette on save.

The inline desktop editor now has an authoritative bounded vector reader. Its
range edits are converted back to canonical seriesData references; the host reads
the referenced cells again and produces caches, resolving values and categories
on separate sheets when needed. Original chart overlays store merged series edits,
and exact referenced material is hydrated so later scalar edits refresh those
journaled caches. Session charts retain their new references rather than refreshing
from the old range. Explicit series names replace their previous name reference.
An inline-apply lifetime counter covers both range reads and commit/receipt, with
idempotent disposal cleanup, so native Save cannot finish between the read and
transaction. Missing original caches hydrate through the same read port for display
without authoring the document. Full live refresh of untouched/reopened original
references and broad formula freshness remain separate open work.

The plugin now hosts the actual desktop ChartFormatPane outside the grid so its
lifetime survives projection replacement. Context-menu Format Selection, data-label
toggling and deletion use the shared menu, while Select Data and unconnected visual
commands remain disabled until their authoritative paths exist. Format controls
submit through the same frozen chart transaction route and serializers as inline
edits. Pending checkbox/dropdown proposals echo locally without authoring data;
text/number inputs retain their drafts on failure. Public editor flush blurs active
pane inputs and waits for their edit. Pointer-captured non-drag clicks delegate to
the original chart hit element so existing title/axis/series/point selection handlers
work, preserving the established movement/resize capture. Full contextual ribbon
bands and the remaining commands are still required for the goal.

Select Data now uses the shared desktop dialog and a declarative canonical
`edit_chart.seriesSelection` input. Source indices and optional bounded references
are resolved against the frozen current root in the host. Original and session
series replacements share one pure selection builder; source overlays clear old
index-based formatting before consolidation, and retained references remain live.
The dialog awaits commit and participates in Save flushing. Contextual chart
ribbons and the broader feature acceptance remain unfinished.

Chart Design now mounts the actual desktop SheetsRibbon command band in a dynamic
contextual tab. Visual selection updates the descriptor from the current session;
keyboard navigation includes the contextual tab and deselection returns to Home.
The desktop and plugin share `chartRibbonEdit` for presets, palettes, labels,
legend, grouping and text edits. The desktop title/axis-title prompt is extracted
as a browser-only shared component with asynchronous commit/error handling. Plugin
text dialogs freeze the chart identity and opening root/revision and join chart
Save flushing. Source and session charts use the same canonical edit/history/save
path already used by the chart pane and Select Data dialog.

Background recovery explicitly opts out of blurring an unsubmitted chart-format
draft while waiting for canonical work and flushing cell edits. Explicit Save,
history and current-state tool flushes keep their default chart-draft commit. The
bridge validates the flush options; real native backup/focus/save tests verify the
distinction.

Floating visual properties now use canonical `edit_visual` (`visualId`, optional
`altText`, `editAs`, `zIndex`). Targets are editor visual identities; the standalone
file CLI rejects this identity-scoped operation explicitly, as it already does
editor visual deletion. The host resolves trusted drawing ownership, verifies
protection/removal and validates placement/layer bounds. Source edits compose in
the drawing journal and session chart metadata/order stays in the current model.

The desktop and plugin share `projectVisualOrder`, preserving the drawing writer's
edited-slot permutation rule. Absolute source layer edits expand into a consistent
part-wide order; session charts reorder their model array. The desktop dispatcher
also accepts the canonical metadata and absolute layer operations through its
existing undo-aware journal paths. Both applications reuse VisualDialog with
asynchronous commit/error handling. Chart dialogs retain the desktop's placement
controls; chart width/height resizing remains on the existing in-grid handles.

Source shapes and pictures now render with the existing desktop WorkbookVisuals
components alongside charts. Selection exposes the actual Shape Format/Picture
Format SheetsRibbon bands. Their commands use the same source drawing journal,
canonical history and package writer as the earlier visual-property operations.
`edit_visual` extends the shared drawing properties for paint, rotation, flips and
true frame size; unsupported target families are rejected. Frame-size changes pair
with an earlier anchor operation in the same transaction. Desktop lazy dispatch
applies equivalent target/placement/frame checks.

The desktop frame-center/rotated-AABB conversion is extracted into
`frameChangeFromBox`. The plugin uses it for size/rotation and desktop align/distribute
helpers for selected installed frames. All affected anchors are submitted together
against the captured root/revision, with full visual-operation Save flushing.

Floating-source images use `source.media` through the validated bridge. The host
resolves a unique current image/shape identity and its surviving logical owner,
reads the fixed source session, validates the bounded image response and rejects
reads overtaken by another root. The browser receives only MIME/base64 data and
uses the desktop media decoder. New-shape/new-picture insertion and clipboard/link
flows remain separate unfinished requirements.

New session shapes share the desktop Shapes gallery and `startSheetShapeDraw`
gesture. The gesture accepts the plugin grid container, returns idempotent
cancellation and blocks underlying visual drags; Escape, window blur and disposal
cancel it. Its frozen sheet/root submits an explicit bounded anchor, so stale
completion rejects. Text boxes retain the desktop's immediate cell-based defaults.

Session `features.visuals` now holds an ordered chart/shape union; `features.charts`
remains a filtered compatibility projection. Shape IDs derive from the canonical
transaction allocation. Existing source visuals retain their drawing journal.
The package writer walks the ordered union and writes at most 50 additions per
stage, preserving the complete list across repeated saves. Shared shape text
editing awaits canonical acceptance, retains failed drafts and participates in
explicit Save flushing without making automatic backups commit live drafts.

Pictures reuse `pickPictureFile`, `readPictureFile`, `pictureFileProblem` and
`fitPictureFrame` from the desktop picture module. The host-facing operation is
`add_picture`: validated PNG/JPEG/GIF bytes, name and a bounded explicit anchor,
without browser-supplied filesystem handles. Existing desktop `add_image` path/URL
loading remains a separate entry point. Native VS Code tools advertise the new
canonical operation using the shared generated schema.

Pictures join the ordered session visual union and existing formatting/history
paths. Their accepted bytes live inline in the canonical root, so the existing
chunked history and recovery serialize them without a separate transient media
cache. Per-file and general transaction/history bounds still apply. Picture-bearing
candidates execute the same shared interpreter directly in the extension host:
repeating a near-limit image in candidate material, next/forward/inverse over the
bridge would otherwise exceed its transfer budget. The accepted root alone is
projected to the renderer. Source media retains its identity-scoped host reader.
The package writer emits pictures alongside charts and shapes in authored order,
retaining the existing 50-addition stage bound and original-byte replay model.

The plugin now installs the actual desktop `installPictureTransfer` lifecycle
against its grid container. An optional scoped key-target check recognizes the
embedded hidden focus host without relying on the desktop container ID. The
scene/scroll/scale `cellAtClientPoint` helper is extracted unchanged for both
applications. Pictures freeze their sheet/root and explicit cell anchor before
async file processing, and the pending work participates in Save flushing.

Native cell clipboard input is distinct from a floating image, including tables
with preview bitmaps and text with an incidental image. Both applications pass
such payloads to the desktop HTML/text parser without file items. The plugin
explicitly initializes `SheetClipboardController` after workbook creation; its
registration is unavailable before that lifecycle point when native toolbars
are hidden.

`planNativeCellPaste` temporarily captures public clipboard-hook mutation plans,
uses a sentinel operation to stop native mutation/history application, restores
the hooks in a finally block and converts supported value/style/merge operations
into one authoritative batch. Frozen target/root validation and existing host
protection/budget checks remain in force. Unsupported mutation families and mixed
rich text reject before any batch is published. General full clipboard acceptance
(cut, source freshness, Paste Special, rich content and automatic row sizing)
remains open. The shared neutral-style mapper now accepts computed CSS RGB colors
so external HTML font/fill/border colors survive serialization.

The shared Home Paste button reads the first ClipboardItem using the same browser
API as the desktop command, dispatches via the shared picture/table classifier
and submits the same canonical picture or external-cell operations. A separate
tracked clipboard-read promise participates in Save flushing alongside the nested
visual commit. The selected sheet/anchor/root is captured before the read; denied
reads leave the root unchanged and provide a localized native-shortcut hint.
Platform clipboard-cache fallback, cut/internal source semantics and Paste Special
remain separate acceptance items.

Home Copy and native grid Ctrl/Cmd+C now read the complete selected typed block
from the authoritative frozen root before invoking the desktop clipboard service.
The native copy identity is paired with a detached typed snapshot, retained across
projection reinstalls and source edits. Matching HTML pasted within that editor
uses `paste_cells`: one bounded rectangular typed operation with replacement
formats and optional replacement merges. This preserves literal strings beginning
with `=`, errors, date systems and independent border edges; relative formulas
shift from the original copied coordinates and become pending without old caches.
The shared scalar cell schema lives below both session state and the DSL to avoid
a circular schema dependency. Existing merge/protection/write-budget checks apply;
partial destination merges and mixed date systems reject atomically. Full copies
can tile a compatible selected target, including translated source merges.

Copy reads and writes participate in Save flushing; revalidate the selected sheet,
range and root before writing the native clipboard. This is a same-editor copy
snapshot path, not complete clipboard parity. Cut requires reference-aware moves;
Paste Special, copied links/notes/validation/conditional formatting/rich text,
filter/hidden-row parity, cross-editor snapshot portability, denied-copy reporting
and platform clipboard fallback remain open. The desktop DSL executor supports
typed values/styles/merges but explicitly rejects literal error cells before a
batch, because its legacy journal cannot preserve that type. The VS Code canonical
writer supports literal errors and is independently verified through save/reopen.

Home Cut and grid Ctrl/Cmd+X now use the same authoritative source capture and
native desktop CUT clipboard serialization as Copy. Cut does not edit the root
until Paste. A matching cut identity dispatches canonical `relocate_range`, not
legacy `move_range` (whose adapter expansion remains copy followed by clear).
The operation materializes complete worksheets to find dependents beyond the
viewport, checks source/destination/dependent write permissions, captures the
source before clearing and supports overlapping source/destination areas.
It moves typed values and replacement formatting, contained merges, plain notes
and hyperlinks as one reversible transaction. Same-editor clipboard identities
are consumed after successful Cut/Paste; Escape cancels a pending cut. A root
changed since Cut rejects before clearing either block.

`moveFormulaRefs` shares the existing A1 tokenizer's literal/table protections.
References wholly contained in the source follow it, including `$` references;
partial ranges retain their coordinates, matching the desktop formula controller
MOVE_RANGE condition. Moved formulas retain outside dependencies on their original
sheet by qualifying formerly bare references. Cross-sheet dependents use escaped
sheet qualifiers; copied formulas continue to use the separate fill-offset path.
External workbook prefixes, short 3D sheet names and table-column names are kept.

This is an incremental cut implementation. Affected defined names/native chart
references, source/target tables/pivots/sparklines, filters, conditional formatting,
data validation and rich/image/threaded notes still need their own reference and
metadata transforms and currently reject before commit. Session chart reference
strings are transformed with the cell references; wider chart acceptance remains
open. The ordinary 2000 expanded-write budget includes source clears, target writes
and dependent formula writes. Full dependency materialization uses the existing
200,000-cell per-sheet material budget. Standalone legacy desktop/CLI DSL paths
explicitly reject the canonical-only opcode; the desktop's native Cut UI remains
unchanged. Paste Special, cross-editor cut, clipboard permission/fallback behavior,
remaining ribbons, DOCX and installed-platform acceptance remain open.

Paste Special reuses the actual desktop dialog, availability model and extracted
pure transform (`paste-special-transform.ts`); the desktop module re-exports its
existing API. The plugin connects its five quick entries, full dialog and grid
Ctrl/Cmd+Alt+V. Clipboard source/target/root are captured before asynchronous reads,
read/apply work participates in Save flushing, failures keep choices for retry,
and successful application closes only after the transaction commits.

`paste_cells` now supports selective value/format fields and bounded skipped-cell
addresses. Canonical composition retains target data for Formats, target styling
for Values/Formulas, all unrelated fields for number-format variants and target
edges for All Except Borders. It can carry a styled 101-cell block as one operation
rather than exceeding the scalar command count. Classic standalone desktop/CLI
DSL executors explicitly reject these canonical-only variants; native desktop
Paste Special continues using the shared implementation.

Source rules follow the installed desktop implementations rather than treating
all choices as one frozen copy: plain Values/Formulas retain cached copied data;
plain All/Formats/Except Borders read current source formatting; advanced own-layer
operations read current source data; Column Widths reads current source widths.
Links use the original source coordinates and current sheet name. External basic
HTML choices use the desktop clipboard hooks; external text transforms use the
shared parser/transform and retain the desktop source-dependent disabled choices.
Comments/Validation stay disabled exactly as in the shared desktop model.

Full completion is still unproven. Pending source/target formula values are
rejected where arithmetic/value/transpose requires a verified computed result;
Formats and Links do not require such a result. Comprehensive CF/validation/rich
metadata, filtered/merged target parity, source deletion/structural changes,
cross-editor clipboard ownership, all locales/appearances and installed-platform
acceptance remain open alongside the other base/Insert/DOCX ribbon work.

### Shared Page Layout checkpoint

The plugin now mounts the actual desktop Page Layout band. Desktop and plugin
share a schema-validated command decoder for orientation, paper, margin presets,
scale/fit, print area/titles and page breaks. The plugin reads indexed source
settings before preparing a frozen canonical transaction and includes that work
in Save flushing. Saved source settings and canonical changes drive the ribbon
labels; screen grid/headings and print grid/headings remain separate.

Page-only transactions do not hydrate chart cell ranges or recalculate chart
caches, while still indexing source protection metadata. This preserves pending
formula chart caches when editing print metadata. Shared fit controls use plain
layout containers rather than label wrappers, preventing native VS Code menu
activation from swallowing subsequent clicks.

Document Themes, Theme Colors and Theme Fonts are visibly unavailable until a
canonical workbook-theme transaction and save/projection path exist. This
checkpoint does not establish full Page Layout or full desktop-menu parity.

### Shared Function Library checkpoint

Extract the desktop Function Library group rather than build a second catalog.
Reuse live engine descriptions and the actual catalog/arguments flow, supporting
awaited canonical writes alongside the desktop's synchronous host. Keep the
frozen owning target/root across pointer range picks, retry failed writes, and
include preparation/commit work in Save flushing. Hook fx only once a workbook
projection has initialized its formula editor service.

Resolve bounded static preview references from current host reads before engine
preview and retain host authority/history. Do not turn unverified formula caches,
names, dynamic or external references into zero values. Those previews, remaining
Defined Names/Auditing/Calculation groups, and full platform acceptance remain
open in `sheets-formulas-ledger.md`; rendering a shared library does not complete
the Formulas tab or the full goal.

### Canonical defined-name foundation

Enable existing add/delete defined-name operations in the authoritative session
with an optional logical `scopeSheetId` (omitted means workbook scope). Journal
the complete modeled name set separately from immutable source metadata and use
the existing declarative defined-names writer. Map logical scopes to current
sheet indices only at save. Preserve unmodeled source names, hidden names and
Excel built-ins. Reject case-insensitive duplicates within a scope and workbook
protection; replacing a definition through delete/add remains one transaction.

Share the pure name projection installer with desktop and use current journal
names in Go To. Name-only edits do not hydrate or rewrite unrelated chart data;
original formula caches become pending after name changes. The standalone CLI
and legacy desktop DSL explicitly reject the new logical-scope field rather than
silently treating a local name as global.

The shared Name Manager now uses an authoritative ordered reference plan.
Existing overlay formulas rewrite at action time, while immutable source formulas
replay the plan on read and save. Each action captures logical local/global scope
and existing local shadows, including hidden reserved bindings. Later formulas do
not inherit earlier renames. Removal produces `#REF!` tokens. Bulk label-derived
names use one canonical operation and bounded authoritative formatted-label reads.

The native reader reports name-binding completeness and reserved hidden/empty
names with scope IDs, without publishing hidden formula content. Reference-aware
and bulk actions reject older readers with incomplete binding metadata. Source
hidden/empty definitions, print names and foreign extension payloads remain intact.
The same shared dialog supports synchronous desktop and awaited plugin writes;
failures retain drafts and Save waits for reads/commits. Use in Formula freezes
its target and accepts only a verified empty cell.

Destination renames that introduce a new local shadow currently reject pending
workbook-qualified resolution. Full scoped computation, structural-reference
replay, all remaining menu groups and broad installed-platform/locale/appearance
acceptance remain open. This does not complete the full desktop-menu goal.

### Formula Auditing adapter

Share the desktop auditing group, pure reference parser, float arrows, formula
view interceptor and Watch Window. The plugin uses a dedicated read-only
source.audit method, validated at both bridge ends and tied to an immutable root.
It scans the same complete indexed source plus current overlays as Find, but has
no shared query ticket that could invalidate an active Find/Replace panel.
Name-reference logs also replay during formula searches, so renamed source
formulas do not retain their old spelling in audit/search results.

Show Formulas extends set_page_setup and uses the existing worksheet-view writer.
Its display interceptor installs only after workbook creation, with per-projection
cleanup and active-sheet context synchronization. Trace/error queries freeze the
root/target and reject stale publication; Save waits for pending audit/view work.
Watches use stable logical sheet IDs and authoritative bounded cell reads; pending
formula values render as pending. Native source/save/history/reopen, complete
browser regression and wider calculation/platform acceptance remain required.

Calculation follow-up evidence and required canonical/cache integration are in
[calculation-notes.md](calculation-notes.md). A Node experiment proves that the
same Univer engine can calculate a sparse workbook without a DOM; this is not a
production calculation service or completed menu. Host-only authority is required
for volatile results, with explicit result/cache freshness and manual-mode behavior
through projection rebuilds and save/reopen.

### Protected worksheet view access

Candidate material carries an optional viewWritableSheetIds list alongside normal
writableSheetIds. The host grants it only after complete indexed protection
metadata and for non-readonly targets. The candidate kernel permits only a pure
set_page_setup patch of showFormulas/showGridlines/showHeadings/zoomScale through
this grant. Content writes and mixed page/print patches still use ordinary write
checks. Both host and browser share the predicate, the strict bridge explicitly
validates the new list, and the host re-executes accepted candidates. The grant
is not a replacement for broader desktop protection/readonly permission mapping.

The normalized-snapshot computation adapter is now implemented in workbook-session
(calculation/univer.ts). It observes rich values at the runtime writer boundary,
uses private engine unit IDs, and supports actual workbook/sheet dirty semantics.
It is not a canonical document calculator yet: complete source/context hydration,
worker ownership, frozen-root result cache publication and calculation-mode/UI/
save behavior remain required. See calculation-notes.md for observed engine
pitfalls and the remaining integration contract; do not route the ribbon to an
unverified viewport/projection calculation as an intermediate final state.

Complete stored-cell loading now lives in calculation/source.ts. It uses the
native indexed stored_cells mode, current coordinate/name transformations and
bounded sparse reads. The result is host-private logical data with separate
source-space features and array ownership, not a certified engine snapshot.
Private package definitions/resources, array/table transformation, row/column
context, worker bounds and canonical result/cache publication remain required
before connecting Calculation commands. See calculation-notes.md and the real
sidecar page-boundary tests for the exact implementation boundary.

The complete-source loader now also requires host-private read_calculation_context.
The source session binds the file path; hidden/empty/built-in/scoped definitions,
calcPr attributes and unresolved external IDs stay outside public metadata/state
and renderer routes. Raw source context is retained separately from logical
journals. Calculation-time name binding remains open: the installed engine's
getValueByName cache does not include worksheet scope even though its name type
has localSheetId. Resolve and verify this before treating a registered resource
set as a correct authoritative engine snapshot. See calculation-notes.md.

The normalized computation boundary now resolves logical definedNames through a
private alias binding pass over the actual engine tokenizer, with lexical LET/
LAMBDA scope, runtime dynamic-name lookup and shared-formula expansion before
provenance capture. This addresses spelling-only engine caches without changing
authored formulas. Source-context/journal transformation into this normalized
input remains required, along with recursive/iterative and arbitrary dynamic-name
semantics, table/array/external/date/default resources and canonical result-cache/
mode/save/UI ownership. Real engine checks and limitations are in calculation-notes.md.
