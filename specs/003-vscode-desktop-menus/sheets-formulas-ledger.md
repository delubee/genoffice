# Sheets Formulas parity ledger

This is an intermediate checkpoint toward the complete desktop Formulas band.
The desktop and plugin now share the Function Library group, actual function
catalog and Function Arguments dialog. The plugin still exposes its legacy
manual formula input/calculation command in additional controls. The other
Formulas groups are not accepted as complete.

| Desktop group | Controls | Current state |
| --- | --- | --- |
| Function Library | Insert Function | Shared live catalog, search, recent history and two-step arguments flow; canonical frozen-target writes |
| Function Library | AutoSum, Average, Count Numbers, Max, Min | Shared desktop per-column formula geometry; one canonical transaction |
| Function Library | Recently Used, Financial, Logical, Text, Date & Time, Lookup, Math & Trig, More Functions | Actual shared category buttons and catalog filters |
| Defined Names | Name Manager, Define Name, Use in Formula, Create from Selection | Actual shared group and asynchronous Name Manager connected; canonical scope-aware rename/removal reference replay, bulk label creation and empty-cell insertion implemented; broader acceptance remains open |
| Formula Auditing | Trace Precedents, Trace Dependents, Remove Arrows, Show Formulas, Error Checking, Watch Window | Shared actual group, arrows/interceptors/panel; full-source immutable dependency/error queries, canonical formula-view setting; native history/save/reopen on protected and ordinary sheets plus 106-case full browser regression pass; wider calculation/permission/platform acceptance remains open |
| Calculation | Automatic/Manual, Calculate Now, Calculate Sheet | Desktop menu/settings parity open; legacy recalculation control remains available |

The shared arguments flow supports synchronous desktop writes and awaited host
writes. Pending submission disables actions and parameter changes; failed writes
retain the dialog and arguments. The plugin freezes owning sheet/cell/root so a
range pick or selection change does not retarget the formula. Overtaken roots
reject instead of authoring against newer content. Save waits for pending commits.

Preview resolves bounded static precedents from the authoritative host before
using the display engine, including off-page and cross-sheet values. It changes
only the projection, never session roots/history. Preview currently declines
named/external/dynamic references, over-2,000-cell precedent reads and formula
precedents (including cached ones; preview computation is not yet established). Those are preview gaps,
not zero-valued inputs; formula insertion itself remains available.

The native formula-bar fx listener is scoped to a ready workbook projection and
disposed with that projection; grid Shift+F3 is scoped to the embedded grid.
The shared picker accepts the embedded grid root while retaining the desktop
container default. Full editing/caret/cancel, computation freshness, locales,
appearances and installed-platform acceptance remain open.

The Defined Names group now reuses the desktop dialog and commands. Existing
cell overlays rewrite references at the action revision; original formulas replay
an ordered binding plan on read and save, including off-page cells. Later user
formulas do not inherit earlier actions. Hidden/local bindings are reserved by
the native reader without exposing their formula content. Hidden and empty source
definitions, print names and foreign worksheet extension payloads are preserved.

Create from Selection reads authoritative formatted labels (including number/date
formats), shares desktop label cleanup, and creates up to 2,000 names in one
transaction. Pending writes block dialog actions, failed drafts remain available,
and Save waits for reads and commits. Stale roots reject without a definition write.

A global rename whose destination introduces a new local shadow is currently
rejected; workbook-qualified resolution, complete scoped computation and broader
locale/appearance/platform acceptance remain open. The complete Formulas tab is
not yet accepted.

Formula Auditing shares the desktop reference parser and rendering helpers.
Precedents use the authoritative owning formula; dependents scan the complete
immutable source plus current overlays, including off-page and cross-sheet cells.
Cross-sheet arrows follow desktop behavior: count and report them, without drawing
an arrow onto another sheet. Remove Arrows and tracing are presentation-only and
are disposed with their workbook projection. Error Checking uses the full index,
cycles in workbook/row/column order and does not consume Find/Replace tickets.
Overtaken queries reject before drawing or navigation. Saved formula values that
await recalculation are explicitly disclosed; complete calculation freshness is
still open.

Show Formulas uses the actual desktop content interceptor and raw-formula context,
installed after workbook creation and disposed with each projection. It is a
per-sheet canonical setting with native history/save/reopen support, rather than
a rewrite of authored cells. Ctrl/Cmd+Backquote stays scoped to the embedded grid.
Watch Window shares the actual desktop panel and 20-new-watches-per-add geometry;
pinned cells resolve by logical sheet IDs, use authoritative reads, and show
pending computation instead of presenting stale values as current. The panel
survives ribbon-tab changes. Scope-specific computation, destination qualification,
large-file/performance and broad locale/appearance/installed-platform acceptance
remain open, alongside the Calculation group and all other pending menus.

Protected-sheet view writes now have a separately indexed host grant for
showFormulas/showGridlines/showHeadings/zoomScale. The strict bridge validates the
grant; missing grants, readonly state, content and mixed print/view operations
retain ordinary restrictions. Real protected-source save/reopen preserves the
protection/password and source cells, and both native protection states pass
history/reopen tests. Broader print/layout and readonly presentation policies,
calculation freshness and installed-platform/locale/appearance acceptance remain
open. The historical intermittent shape case remains an audit item even though
the latest full 106-case browser run passed.
