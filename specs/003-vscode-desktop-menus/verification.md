# Verification: shared desktop ribbons and DOCX dialogs

## Outline persistence checkpoint

Menu/native checkpoint: the plugin Data panel now exposes Group/Ungroup Rows
and Columns, plus controlled summary-below/summary-right checkboxes using the
desktop labels. The runtime reuses the desktop groupLevelOps planner over a
complete host outline read, captures the worksheet/root/revision, and submits
each relative action as one canonical transaction. Equal-level spans are split
at the canonical 10,000-line operation limit; levels stay within 0–7. Pending
outline work participates in Save flushing, concurrent outline actions are
rejected, and stale reads cannot commit.

Native VS Code 1.109.5 acceptance passes for rows and columns: group twice,
ungroup once, switch both summary flags, assert eight revisions and exact native
Undo/Redo roots, then save and independently reopen in a new clean session.
Saved row/column level 1 and both false summary flags read back from the native
metadata; numeric source content and the custom XML sentinel remain unchanged.
The test waits for the asynchronous checkbox transaction before asserting state.
The full extension build/browser-boundary checks, TypeScript, changed-file lint,
source-inventory check and whitespace check pass. This checkpoint covers the
plugin-owned menu adapter, not the complete shared desktop Data band. Shared
gutter rendering, collapse/expand and Clear Outline are covered by the follow-up
below; full edge/lifecycle variants and installed Windows acceptance remain open.

Complete-outline read follow-up: the sidecar source now exposes its existing
read_row_outline capability. The host waits for complete indexing, maps source
row/column coordinates through earlier structural changes, and replays authored
absolute-level/settings edits. A strict `source.outline` JSON route checks the
root before and after reading. No cell viewport is used as a complete-outline
substitute. The editor host port now drives the relative group/ungroup menus;
outline gutter projection is now connected through a presentation-only port.

Six tests pass across read-outline and outline-save, including rows beyond the
viewport, source column spans, insertion shifts, partial-index waiting, journal
override/clear, foreign-journal rejection, cancellation and a real sidecar save/
reopen whose complete read returns the expected levels and collapsed summary.

The session DSL now supports absolute row/column outline levels (0–7), optional
collapsed flags and summary placement. Canonical operations append to a feature
journal and the package writer validates/replays that journal through the existing
set-rows-outline/set-cols-outline/set-outline-pr XML path. Coordinate-changing
operations on sessions with authored outline edits remain rejected until their
reference transform is implemented; this is an explicit parity gap.

Two focused tests pass, including real sidecar source/save/reopen: row levels,
column levels, a collapsed summary row, summaryAbove, native outline maxima,
the original numeric cell and custom XML sentinel are verified. The inverse
root is unchanged and grouping does not author cell contents. Schema checks
reject column-grid overflow and levels above seven. Session typecheck and
changed-source lint pass. The plugin group/ungroup menus are now connected;
the shared outline gutter is now mounted as described below.

## Outline detail and shared gutter follow-up

The existing set_rows_hidden/set_cols_hidden operations now enter the session
whitelist and write a validated current-coordinate visibility journal. Current
range rows, complete outlines, initial grid construction and later row hydration
replay that journal. The package writer applies visibility after the other
worksheet feature stages, preserving source contents, sizes and outline levels.
Coordinate-changing operations remain guarded after authored visibility edits
until exact reference transforms are implemented. Seven read-outline/outline-save
tests pass, including overriding imported hidden flags, last-write precedence,
newly grouped detail and real source/save/reopen with row and column hiding.

Hide/Show Detail reuses desktop outlineGroups/groupTogglePlan and summary flags;
expanding an outer group keeps independently folded children hidden. Clear Outline
uses the desktop single-cell/selection scope, restores folded detail visibility
and removes levels/collapsed flags in the same canonical transaction. Contiguous
clear spans are grouped to avoid one operation per line. Native VS Code acceptance
passes for nested row/column menu folding, saved hidden/collapsed XML, exact
one-step Clear Undo/Redo, and independent clean-session reopen with the numeric
source and custom XML sentinel intact.

The actual desktop outline-gutter module now accepts outline/group presentation
ports instead of importing Electron/lazy-workbook action and synchronization
modules. Desktop App provides its original adapters; the plugin supplies complete
host outline reads bound to the active sheet/root, discards overtaken background
reads and disposes the shared gutter with its view. Group clicks carry exact
group/root identity; level buttons reuse desktop levelButtonPlan. The same canvases,
geometry, header sizing and styles render in both applications, with a reserved
strip beside the embedded grid. Browser pointer acceptance passes for exact inner
row/column group buttons, level 1/3 visibility, both gutters present and strip
removal after clear. Native VS Code acceptance also passes for the row/column depth buttons with
both gutters present and exact Undo back to the folded root. The initial native
run exposed an actual narrow-window defect: shared desktop CSS imposed a 720px
body minimum width and a tall Data panel left only 49px for the grid. The
embedded theme now removes that minimum, constrains the panel height relative
to the viewport and confines the grid content. A 591×365 browser acceptance
verifies no horizontal document overflow, navigates to offscreen summary lines
and clicks the shared group/depth buttons using actual viewport scroll geometry.
Native test coordinates wait for the raised column header and scroll the gutter
into view before clicking. Full native exact-group glyph/scroll/zoom variants
remain open.

Full desktop and plugin typechecks, changed-source lint (desktop App retains
three existing React hook warnings), inventory consistency and the complete
extension build/browser-boundary check pass. The final native run passes both
Outline UI scenarios (41 seconds) without a Canceled diagnostic. Earlier attempts
emitted bare Canceled messages during view replacement; their broader lifecycle
audit remains open. Imported-outline/gutter reopen, filtered/protected selections,
zoom/frozen geometry, theme/locale and installed-platform variants remain open.

## Goal Seek private-source evaluation checkpoint

Current UI/native checkpoint: the actual desktop GoalSeekDialog is mounted in
the plugin with controlled locale and a JSON-validated host port. Opening flushes
pending edits before capturing the sheet/root. Save waits for the solve; duplicate
actions are blocked; returned public state/revision updates the dialog binding.
Native VS Code 1.109.5 acceptance passes for `=A2*3+Other!A1`: the dialog reaches
22 by changing A2 to 5 with exactly one revision, exact native Undo/Redo restores
the roots, and independent reopen reads the solved input, unchanged formula,
source value and custom XML sentinel from the saved file.

The first native run exposed a delayed progress notification after the Node
engine's locale service had been disposed. The engine adapter now drains terminal
command microtasks before disposal. It also disposes the browser-only image-size
interceptor in Node; otherwise draining image results invoked the unavailable
browser Image constructor. All 27 calculation-runtime/private-solver tests pass,
including delayed-progress observation after Goal Seek; type/lint checks and the
full extension build/boundary check pass. Nonlinear/no-solution UI, broader
permissions/resources/lifecycle and installed Windows acceptance remain open.

Host controller/bridge follow-up: `source.goalSeek` now validates a strict request
containing state/revision, worksheet, target cell, changing cell and numeric goal.
It preflights writable/merge policy through host material, loads a frozen source,
rejects stale or disposed jobs and submits one normal manual set_cell transaction
only on convergence with a changed input. Nonconvergence returns a JSON-safe null
for an error-valued reached result. The overall controller deadline is shorter
than the bridge timeout. Six controller tests pass for final-only writes, unchanged
roots, nonconvergence, protection, stale roots, disposal, invalid requests, no-op
history and overlapping runs. Existing search/picture bridge regressions pass.
The menu/dialog and native linear solve path are now covered by the checkpoint above.

`seekCalculationSource` clones a hydrated calculation source and changes only
that private copy for each trial. It uses the existing Node formula engine and
bounded shared secant solver, rejects non-scalar target results and nonnumeric
changing cells, checks the owning root before/after evaluation, and propagates
cancellation with a 30-second abort deadline. Existing calculation snapshot
limitations (including unsupported array ownership/date systems/resources) still
apply. This is not worker isolation or a UI/commit adapter.

Three bundled Node tests pass with the real formula engine: cross-sheet linear
convergence, nonconvergence without source mutation, and stale-root/cancellation/
invalid-input rejection. The tests compare the original source before and after;
no workbook transaction is issued. Session typecheck and changed-source lint pass.
The bridge, permission checks, final root-checked commit and desktop dialog are
connected; the checkpoint above describes the specific native acceptance scope.

## Shared Custom Sort checkpoint

Native VS Code 1.109.5 acceptance now passes for a two-key sort: the first key has
ties that require the second key, and payload cells include boolean true, the
string `00123`, and independently styled cells. The test uses normal pointer
interaction with the shared dialog, checks one revision, verifies exact native
Undo/Redo roots, saves, closes the original document, and opens a new session.
Host snapshot reads after reopen verify sorted numbers, preserved boolean/string
types, movement of bold formatting with its source row and removal from the old
destination. Header values and the package sentinel remain intact. Extension
typecheck, test lint and whitespace checks pass. Three-key edge cases, broader
selection states and installed Windows acceptance remain open.

The desktop SortDialog is extracted into a browser-safe shared component and
mounted in the plugin Data panel. It retains three levels, per-level directions
and the header flag, with controlled locale, asynchronous disabling and rejected
draft retention. Column labels are read from the captured host revision. The
target range/revision remains fixed and one sort_range operation carries all keys.

A real plugin browser scenario passed: it selects Sales as the primary key and
Region as the secondary key, rejects the first submission, retries the same draft,
checks exactly one revision and verifies preserved headers plus reordered typed
cell values. Five shared sorting tests cover multi-key tie/blank semantics and
exact authoritative rewrites. Editor and desktop Sheets typechecks pass; changed
source lint has no errors (the existing ExcelShell onListFunctions dependency
warning remains). Native custom-sort history/save/reopen, all selection variants
and installed-platform acceptance remain open.

The full extension build and browser boundary guard pass after using the shared
UI chrome entry for SortDialog. Native tool schemas were regenerated for the new
ordered sort levels. The control inventory includes the extracted dialog source.

## Consolidate implementation checkpoint

The plugin Data tab now opens Consolidate with multiple references, five
aggregation functions, position matching and left-column label matching. It
uses the desktop matrix builders and commits typed labels and live formulas
through one frozen-revision host transaction. Formula-looking labels remain
literal strings. Oversized output is rejected before matrix generation; the
current host transaction limit is 2,000 output cells. Source/destination overlap,
merged destinations and revisions changed during reads are rejected.

The shared consolidation algorithm suite passed all nine tests. Sheets editor
TypeScript, changed-source ESLint and whitespace checks passed. The full VS Code
extension build, including its browser boundary check, passed. These checks do
not establish actual dialog interaction, native history or save/reopen acceptance
for this new command; those remain required alongside larger-output parity.

Follow-up browser acceptance passed on the real plugin ribbon/dialog with an
in-memory session authority: a rejected commit retains its draft, retry creates
exactly one revision with label values and SUMIF formulas, and a dialog whose
revision changed is rejected without another write. Consolidation is now included
in the editor's Save flush barrier and concurrent consolidation is rejected.
Sheets editor typecheck and changed-source lint passed after this change. Native
VS Code undo/save/reopen and larger-output acceptance remain outstanding.

Native VS Code 1.109.5 acceptance subsequently passed (one integration test):
normal pointer clicks select the Data Consolidate command, combine Data!A2 and
Other!A1 into C1, and produce exactly one revision. Native Undo restores the full
prior state; Redo restores the committed state. Save writes the expected SUM
formula into worksheet XML and preserves the custom XML sentinel. The first run
was obstructed by workbench notifications/auxiliary chat; the test now clears
notifications and closes sidebars before normal pointer interaction. Independent
save/reopen and installed-platform/large-output coverage remain pending.

The extended native test now also passes independent reopen: it waits until the
old custom document is disposed, opens the saved file in a distinct session,
checks revision zero and a clean document, then reads C1 through the host snapshot
and verifies the exact cross-sheet SUM formula. Data!A2 and Other!A1 retain their
original numeric values. This proves persistence and source preservation for the
tested position-mode case; installed Windows and larger-output coverage remain
pending. Extension TypeScript, test lint and whitespace checks pass.

This record covers component extraction, dependency separation and the first
embedded DOCX dialogs. It does not prove complete menu parity in the production
VS Code extension.

## Passed checks

- Sheets: 31 tests across `embedded-ribbon`, `ribbon-compact`,
  `insert-function-catalog`, `paste-special`, and `ribbon-dropdown-zorder`.
  Embedded tests render all seven standard and four contextual tab bodies
  without a desktop preload. They check formatting command parameters, number
  format dropdown selection, view-state echo, controlled language changes and
  Escape dismissal. They do not assert that the plugin session executes every
  desktop command.
- Docs: 56 tests across font dialog advanced settings, paragraph breaks/spacing,
  wheel zoom, ribbon font combos, paragraph direction and conversion menus.
- Extension: two browser bundle boundary tests, one per desktop ribbon. These
  bundle the real transitive graph with the same forbidden-import guard used by
  the webviews, and reject inclusion of AI chat/composer, spreadsheet UI services
  or the formula engine. Word still has privileged actions requiring host ports;
  compilation is not proof those actions run outside the desktop.
- Both desktop applications build with `npx electron-vite build`.
- Typechecks pass for Docs, Sheets, the shared UI package and the extension
  (including its separate DOCX configuration).
- Source comparison of the 19 moved Sheets declarations against branch HEAD
  found no changes beyond exports, the component name and formatting.
- Changed-source ESLint has no errors. The existing `ExcelShell` hook dependency
  warning for `onListFunctions` remains. Formatting and `git diff --check` pass.

The isolated, unminified Sheets ribbon probe decreased from 19,800,298 to
4,977,829 bytes after separating Paste Special's UI model from its runtime.
This probe uses esbuild with `write: false`; these are not production VSIX size
measurements.

## Still required

The plugin still uses simplified toolbars, with the real desktop Font and
Paragraph dialogs now connected. Full command adapters, rich DOCX
policy/session support, shared styles, real native chat routing, saved output,
undo/redo and installed Linux/Windows package verification remain to be built
and checked against the complete desktop control inventory.

## DOCX Font and Paragraph integration

The embedded editor now launches the real desktop Font and Paragraph components,
with desktop styles and all existing locale dictionaries. These are working
dialogs on the current simplified toolbar; the full ribbon is still outstanding.

Checks for this change:

- `npx vitest run --root packages/docx-session`: 23 tests passed, including four
  new session tests. Font tests round-trip font slots, size, color, baseline,
  double strike, caps, hidden text, scale, spacing, position and kerning through
  real DOCX bytes and shared undo/redo. They also preserve untouched runs and
  unchanged foreign formatting outside editable bounds, while rejecting invalid
  newly supplied values without committing.
- Paragraph session tests round-trip indents, character-unit first-line indent,
  exact/multiple line spacing, paragraph spacing, pagination flags and outline
  level. They verify explicit off/removal values, original paragraph style,
  untouched paragraphs, invalid-value rejection and undo/redo. Automatic spacing
  flags emitted by the desktop UI executor are explicitly supported.
- Desktop Font/Paragraph component/model regressions: 42 tests passed in three
  files. Extension boundary/manual/save regressions: seven tests passed in three
  files.
- `npm run test:docx:browser -w apps/vscode-office`: one expanded browser scenario
  passed on the built extension webviews. It clicks Font/Advanced controls,
  cancels without modifying state, commits scale/spacing, reopens saved DOCX,
  and verifies one-step history. It also clicks Paragraph spacing and pagination
  controls, saves/reopens, and verifies one-step history. Normal typing,
  headings, table editing, and final undo-to-original checks remain in the same
  scenario.
- The browser scenario switches between Office appearance and simulated native
  VS Code dark theme variables. The page and font preview retain white paper
  and authored ink; appearance changes do not mutate the document. Visual
  inspection caught and corrected unreadable dark-theme font preview ink.
- Screenshots are local ignored artifacts under
  `test-results/desktop-menus/docx-font-advanced.png` and
  `test-results/desktop-menus/docx-font-vscode-dark.png`.
- Extension build, including browser dependency guards, passes. Extension/DOCX
  configuration and DOCX session typechecks pass. Changed-source ESLint,
  formatting and whitespace checks pass.

Browser verification also exposed a resync bug: replacing editor state collapsed
the selection and left selection-dependent inputs stale. The runtime now
restores the full selection where valid and emits a transaction notification
after resync. The browser scenario continues editing after undo/redo/resync and
verifies those controls remain usable.

Installed VSIX checks on Linux/Windows have not yet been repeated for this menu
feature. The complete menu/variant ledger and remaining ribbon adapters are
still required before feature acceptance.

## Shared native file/chat entry points

- Eight host policy tests pass for native action dispatch, frozen DOCX/XLSX
  bindings, partial versus preset chat requests, strict payload rejection,
  session/focus changes during async work and overlapping action rejection.
- One shared component test passes for all five entries, pending-state disabling,
  error reporting/recovery and document/selection chat scope.
- Both desktop ribbon browser-boundary tests pass. Their entry paths now resolve
  relative to the test file, so workspace-root test invocation also works.
- The built DOCX browser scenario passes with the shared action group. It clicks
  Save through the JSON bridge, reopens actual resulting DOCX bytes, then clicks
  document and selection chat. It verifies distinct native bindings, frozen
  selection data and absence of document text in the chat payload. The host ports
  in this browser scenario are test ports; it does not run a real model.
- A new VS Code 1.109.5 integration test passes using the real native action
  adapter and a real custom XLSX editor. It edits A1, invokes the native Save
  command, checks the written worksheet XML, and checks that custom XML survives.
  After switching to another editor it verifies that the original menu action is
  rejected and the file stays unchanged. It also confirms the native chat command
  is registered in the supported VS Code version.
- The two native DOCX product regressions pass, covering native tools, save,
  shared history, Python exchange, recovery and independent XLSX/DOCX targets.
  The revert/close path emits a `The document view has closed` bridge diagnostic;
  disposal/cancellation behavior remains an explicit follow-up, rather than
  evidence of an entirely clean close flow.
- The full extension build and browser boundary checks pass. The host bundle was
  rebuilt after the final native chat argument change. Extension/DOCX, shared UI
  and Sheets editor typechecks pass; dependency declarations and lockfile include
  the new workspace imports. License checking passes with no new external
  dependency. Changed-source formatting, ESLint and whitespace checks pass.

The native chat command arguments (`mode: 'agent'`, `query`, `isPartialQuery`)
were checked against the locally cached VS Code 1.109.5 workbench source.
Native Open/Save As dialogs, real-model chat/skills and installed Windows/Linux
packages are not yet verified for the full menu feature.

## XLSX desktop Home integration checkpoint

The plugin now renders the shared desktop Home ribbon and stylesheet. The other
six standard tabs still use their existing plugin controls. Additional legacy
Home controls remain in a collapsible section while command migration continues.
This checkpoint does not establish full Home or full desktop menu parity.

- The command adapter validates a complete operation batch before submitting it
  to the authoritative session. Font, alignment, borders, cell styles, merges,
  basic structure and sorting commands have initial adapters. Unsupported
  commands report an explicit error; clipboard/paste special, format painter,
  conditional formatting, tables and other desktop actions remain incomplete.
- The real Format Cells dialog submits one batch and one history entry. It
  rejects changes if the document changed while the dialog was open, leaving the
  dialog open and the document untouched by that attempted apply.
- Rich format support now includes double underline, superscript/subscript,
  signed/vertical rotation, indentation, protection flags, independent border
  edges and explicit fill/color clears through projection and serialization.
- Candidate/session regression checks passed: 13 tests in five files. Desktop
  dialog/ribbon regressions passed: 64 tests in five files. Gateway style checks
  passed: eight tests in three files. Sheets, embedded Sheets and extension/DOCX
  typechecks and the full extension build passed.
- A real VS Code 1.109.5 integration test passed: projected double underline,
  superscript, rotation, indentation and outline corners matched the requested
  settings; saved XLSX XML retained those settings and protection flags, an
  untouched sentinel survived, and native Undo/Redo worked.
- The new browser scenario passed, including the actual desktop Home controls,
  font state echo, atomic Format Cells apply, cancel and stale-dialog rejection.
  Existing browser tests needed scoped selectors because the shared ribbon adds
  controls with the same accessible labels as legacy forms. After updating those
  selectors, all 13 Home/manual browser regression scenarios passed. Changed-source
  ESLint reported zero errors (one existing hook warning); formatting, theme-color
  and whitespace checks passed.

Full command coverage, DOCX ribbon integration and installed Linux/Windows VSIX
verification remain open. No new Windows package is produced by this checkpoint.

## Relative Home commands

- Mixed indentation now reads each selected cell's composed style, adjusts its
  own indent, clamps at 0/250, and submits the whole selection in one host batch.
  It retains the desktop 10,000-cell indentation limit. Duplicate, incomplete or
  foreign-cell snapshots are rejected before submission.
- Increase/Decrease Decimal now use `decimal-step-model.ts`, shared with the
  desktop executor. Explicit General records follow the raw-number path;
  percent/custom patterns keep their pattern and advance one decimal place.
  The initial embedded relative-format snapshot path is limited to 10,000 cells;
  support for larger desktop decimal selections remains pending.
- Directional fill now borrows the preceding/following row or column for a
  one-row/column selection, exactly matching the desktop source direction.
  Edge selections produce no writes. Multi-cell selections copy from the
  first/last row or column. The donor is excluded from the destination to satisfy
  range-overlap validation; formulas retain relative and absolute references.
  **Fill parity remains incomplete:** the current host range-copy contract
  retains destination formatting. Desktop COPY also copies source formatting;
  that behavior still needs a host operation and save/projection verification.
- Candidate/session checks: 17 tests passed across the relative Home, rich
  formatting and range-copy suites. Existing desktop decimal regressions:
  seven tests passed. Browser Home/manual editing regressions: 14 passed,
  including mixed indent state, one-revision apply, decimal increase/decrease,
  and adjacent-donor fill through the real shared Home controls.
- Real VS Code 1.109.5 rich formatting regression passed with additional mixed
  relative-indent and decimal transactions, projected padding/number formats,
  actual saved XLSX style XML, and atomic native Undo/Redo. This uses an unpacked
  extension; it does not establish installed Windows/Linux VSIX acceptance.
- Sheets desktop, embedded Sheets and both extension configurations typecheck.
  Browser test servers ignore generated dist output to avoid build-time page
  reloads; the new number-format dependency is explicitly pre-optimized.

The complete desktop menu objective remains open, including the other six Sheets
ribbon tabs, full DOCX ribbon, every missing dialog/pane and the platform package
checks listed in tasks.md.

## Formatting fill checkpoint

The previous value/formula-only fill limitation is now addressed for the modeled
cell-style fields. Home fills request `copyFormat: true`; callers that omit it
retain the existing destination-format behavior. The authoritative session takes
immutable source cell/style snapshots and replaces destination formatting. An
explicit `styleReset` marker survives later edits and is passed to the existing
XLSX writer, which resets the original target xf before applying donor fields.
The browser projection starts from Univer defaults for replacement cells, so old
target formatting cannot leak into the copied cell.

Evidence:

- Nineteen candidate/copy/relative-format tests passed, including removal of
  destination-only italic/border fields, typed dates and number formats, shifted
  formulas, blank donors, source/target tiling and subsequent format edits.
- Sixteen read-current/material/transaction/merged-edit regressions passed.
- Fourteen real browser Home/manual editing scenarios passed. They verify
  copied per-cell indents and explicit replacement flags, then fill from a blank
  donor and verify that the old number format and indent disappear.
- A direct XLSX writer regression passed for replacement formatting and default
  xf restoration on a target that originally had fill/font/indent settings.
- The real VS Code 1.109.5 regression passed. A formatted donor replaced A2's
  original fill and copied its new indent/number format; a blank donor removed
  C2's original rich format. Assertions inspect the saved target cell's actual
  xf (including fillId/fontId/numFmtId), browser projection and native Undo/Redo.
  The XML test reader was corrected to distinguish self-closing xf entries from
  entries with alignment/protection children; the saved XML itself was correct.
- Embedded Sheets, workbook session and extension/DOCX typechecks passed. The
  extension build, browser dependency guards and generated tool-schema check
  passed; the optional fill parameter is reflected in the native tool manifest.

This is not full fill/menu acceptance. Inherited row/column/default formatting,
source features outside the modeled style schema and the complete control ledger
still need audit. The legacy in-memory demo adapter explicitly rejects formatting
fills; production VS Code edits use the authoritative session path. Installed
Windows/Linux package verification and the other unfinished ribbon families
remain pending.

## Clear commands

The real Home Clear menu now routes formats, comments/notes, hyperlinks and
modeled Clear All to the authoritative session. Formats reset the original xf
before applying General format, without clearing values or formulas. Cell
content/format clears split around filtered rows as the desktop does. Notes and
hyperlinks use selected object addresses and retain objects outside the selection.
Clear All decodes the complete batch before a single host submission.

The optional format-reset flag is retained by the in-memory planner, desktop lazy
planner/executor and CLI writer. Hyperlink session edits persist through the
existing package patcher. Unsupported structural transformations of new hyperlink
edits reject before mutation. Known image-note preservation restrictions remain.

Evidence includes 22 candidate/copy/relative/dimension/clear tests, 15 browser
scenarios at the initial clear checkpoint, 16 desktop planner/format/ribbon tests,
and the CLI format-reset regression. The native VS Code 1.109.5 scenario saves
seeded links/notes, clears the selected target, inspects actual worksheet/comments
XML, retains an unselected link/note and verifies atomic Undo/Redo. This does not
prove every Clear All interaction with conditional formats, validation or images,
or large-selection parity.

## Axis dialogs and dimensions

The plugin now opens the extracted desktop AxisSizeDialog from the original
Home Format dropdown. The original desktop is still a consumer. The embedded
adapter adds an asynchronous host acknowledgment, prevents repeated submission,
keeps a failed/stale dialog open and freezes its selection/revision when opened.
Cancel does not change the workbook.

Rows use points and the desktop pixel rounding. Columns use character units and
the workbook Normal font's calibrated maximum digit width. Font digit tables,
availability probing and Normal-font resolution were extracted to pure shared
modules; desktop public imports remain compatible. Column display pixels and
explicit character-width provenance are stored together. Provenance follows
column insertions/deletions and is refreshed by later resizes, so save does not
reinterpret a dialog width using a fixed digit width. Original/default column
projection uses the same font calibration.

Range validation and the package writer accept the desktop zero/maximum row
height (409.5pt) and character width (255) boundaries. The save-layer schema still
had old minimum sizes; the native zero-size test exposed it and it was corrected.
The CLI also honors explicit character widths instead of dropping that field.

Verification:

- 16 real Home/manual browser scenarios passed, including comma decimal input,
  numeric boundaries, cancel, stale apply rejection and unchanged existing header
  drag/numeric controls.
- 26 desktop Normal-font/narrow-column calibration tests passed. 22 candidate
  regression tests passed, covering size unit conversion, chunked row operations,
  character provenance and column insertion shifts alongside recent clear/fill
  work.
- The real VS Code 1.109.5 scenario passed actual XLSX row-height/column-width XML,
  native Undo/Redo, save/close/reopen with the same 34px rows and 80px columns, and
  zero-height/zero-width XML. It also re-runs prior rich-format/fill/clear cases.
- Desktop Sheets, embedded Sheets, gateway, workbook session and extension/DOCX
  typechecks passed. Extension build and generated schema/boundary guards passed.
  Changed-source lint has no errors; existing hook warnings remain.

AutoFit remains disconnected. Font variants requiring canvas measurement in a
headless pixel-only operation, fractional-width behavior and the complete
control/variant inventory still require audit. No installed Windows/Linux VSIX
acceptance is implied by these unpacked-extension tests.

## AutoFit measurement checkpoint

The actual Home Format dropdown now executes AutoFit Row Height and AutoFit
Column Width. The runtime resolves full row/column spans, hydrates original data
pages for the measurement, and calls the same Univer skeleton measurement API
used by desktop. The desktop wrap-budget and calibrated font-pitch patches are
shared. Forced-row automatic-height flags are scoped to synchronous measurement
and restored before committing the resulting dimensions.

The runtime freezes the source state/revision and sheet before async reads,
checks them after hydration and measurement, and checks them again inside the
submission queue. Measurement does not author original source cells into the
canonical overlay. One user action creates one host transaction/history entry.
The rendering package was already installed; its direct workspace dependency is
now declared, with no new external package added to the lockfile.

Evidence:

- 18 Home/manual browser scenarios passed, including an original wrapped D2
  outside selection A2 and an original A80 beyond the initial viewport. They
  verify whole-row height, original-data reads, expanded column width and one
  revision per action. A separate delayed-read case interleaves another edit and
  verifies that stale measurements do not create a resize revision.
- 24 desktop AutoFit span/multi-row/wrap-budget/font-pitch tests passed.
- A VS Code 1.109.5 integration test uses optional local CDP in the isolated test
  profile to click the real webview menu. It targets only the current test-file
  window, verifies original D1/A80 measurements, inspects saved worksheet sizing
  XML, preserves the custom XML sentinel and verifies separate native Undo/Redo
  for each menu click. It passed with `GENOFFICE_TEST_CDP_PORT=9345` and the
  `Desktop Home AutoFit UI` integration grep. Default test launches expose no
  CDP port; this UI test skips unless that explicit test environment is enabled.
- Browser test servers use separate Vite caches and explicit core/render/React
  preoptimization; this avoids concurrent optimizers replacing each other's
  dynamic locale chunks. The expanded UI scenarios allow cold initialization
  time while keeping individual assertions bounded.
- The extension build and browser boundary/schema guards pass. Embedded Sheets
  and extension/DOCX typechecks pass.

This is a measurement-and-apply checkpoint, not full AutoFit parity. Dimensions
currently persist as explicit sizes; automatic row-height mode after future
content changes remains pending. Hydration currently allows 100 source pages,
and the inherited skeleton's viewport measurement extent still requires audit
for larger columns. Font variants, installed packages and the remaining complete
DOCX/XLSX menu inventory remain open.

## Go To checkpoint

The Home Go To entry now opens the actual desktop GoToDialog. Its callback can
await source hydration, keeps failed navigation open with the existing localized
errors and prevents repeated submission while pending. The runtime reuses desktop
reference validation, the same Univer range parser and the extracted frozen-pane
reveal routine. The desktop remains a consumer of that reveal routine.

Navigation supports sheet-qualified addresses and original workbook defined
names. Defined-name priority is shared with the desktop loader: live workbook
names precede local definitions and REF residues, matching the desktop engine's
current first-definition behavior. Duplicate entries are collapsed for the list;
full scope-pair behavior still needs audit.

Distant navigation grows only the display coordinate allocation before creating
the range, then hydrates the target page. It creates no canonical cell/dimension
operation or history entry. Subsequent canonical resynchronization retains the
active selection's display extent. Grid keyboard focus is restored after the
modal closes. The formula engine was already installed through the preset; its
public parser is now an explicit dependency with no new resolved package.

Evidence:

- 19 real Home/manual browser scenarios passed after the build completed. The
  new case checks invalid/out-of-grid and missing-sheet errors, quoted cross-sheet
  names, a live definition preceded by a REF residue, Q2000 beyond the initial
  display allocation, non-navigable formula names and cancel. Navigation leaves
  the entire canonical snapshot/revision unchanged; real typing after a jump
  produces the expected single edit at Q2000.
- 14 desktop address/reveal/defined-name-order tests passed. Shared extraction
  preserves the public desktop import surface and existing behavior.
- The native VS Code 1.109.5 CDP menu scenario passed with source-defined names.
  It opens Go To, switches to Other, double-clicks GoTarget to return to Data A80,
  checks actual selection sheet/range and proves canonical state/revision do not
  change. Prior AutoFit save/history assertions still pass in the same scenario.
  The run prints a `Canceled` diagnostic during lifecycle activity; clean
  cancellation/disposal acceptance remains unproven and requires follow-up.
- Embedded Sheets, desktop Sheets and extension/DOCX typechecks pass. The full
  extension build, schema/boundary checks and changed-source lint/format checks
  pass, with existing hook warnings. One earlier browser run exceeded its cold
  startup timeout while two Vite optimizers and the production build competed;
  the complete subsequent 19-scenario run passed without changing assertions.

Find/Replace is still disconnected from the shared Home ribbon: its desktop UI
requires a real query session and complete-source coverage, not a cosmetic panel.
Scoped-name behavior, grid-edge/full-axis navigation, IME and asynchronous
cancellation/focus remain explicit acceptance work. Installed Windows/Linux
packages and the complete DOCX/XLSX menu ledger are still pending.

Desktop Sheets also builds successfully with `electron-vite build`; production
dependency license checking passes with no new resolved external package.

## Find/Replace query and shared-panel checkpoint

This supersedes the disconnected Find/Replace status at the preceding Go To
checkpoint. Home Find and Replace now open the actual desktop `FindReplacePanel`;
the desktop retains its service and the plugin supplies an authoritative host
adapter through the same presentation contract.

The source query scans complete sidecar index pages, combines canonical overlays
and explicit clears, transforms source coordinates through inserts/deletes and
rewrites formula text through structural edits before matching. It never treats
a loaded grid viewport as a complete workbook. The per-view `source.find` bridge
accepts strict state-bound requests and returns a random query ticket, total and
200 previews per page (512 UTF-16 units per preview, avoiding a split surrogate).
Host source sessions and file handles do not cross the bridge. More than 100,000
matches or 64 MiB of full match data rejects the query rather than publishing a
partial total. Cancelled, superseded, stale and incomplete queries do not publish
success or a false no-match report.

The adapter connects sheet/workbook scope, row/column ordering, case, whole-cell,
values/formulas, live query, previous/next, Find All, history inputs and result
navigation. It gathers complete host pages before enabling replacement. IME
composition defers replace live search; busy controls cannot submit replacements.
Query navigation suppresses delayed grid focus and rejects obsolete late jumps.
Result-row selection updates the replacement match independently of its preview.

Replacement is planned against full host-cached text, not truncated previews or
renderer-authored addresses/content. Literal dollar signs stay literal, numeric
and boolean cells retain their type when the new text still parses, and formula
values are excluded from writes in values mode. Formula text is writable in
formulas mode. All proposed writes validate before submission through the normal
manual transaction/candidate/history path. The ticket is consumed before awaiting
submission so duplicate requests cannot apply it twice. Read-only/protected,
stale, invalid-index, oversized-cell and over-2,000-cell requests reject before
any write; there is no partial/chunked Replace All fallback.

Evidence at this checkpoint:

- Workbook-session suite: 102 tests in 30 files pass, including complete index
  coverage, source overlays, structural/formula transformations, cancellation,
  stale queries, budget overflow, literal/full-text replacement, scalar types,
  formula-value exclusions, consumed tickets and rejection-before-write cases.
  A real Rust sidecar source test verifies indexed formula search and rejects
  a foreign original sheet; search never uses viewport reads.
- Embedded adapter tests cover complete 201-result pagination, actual desktop
  option mapping, off-window navigation, selection of a different result,
  obsolete/failed/incomplete queries, duplicate replacement suppression and
  IME deferral. The four adapter scenarios pass.
- Real JSON bridge search routing and all three desktop presentation boundary
  probes pass. The panel bundles without embedded AI or desktop grid services.
- The existing 52 desktop streamed-find/session tests pass after extraction of
  shared scalar matching/coercion. Ordinary desktop service behavior remains
  behind the same presentation contract.

Full feature acceptance remains open. Fresh calculated values after edits are
not proven; responses conservatively flag possible stale cached formula values
and the plugin reports that state instead of implying fresh calculation. Large
Replace All still rejects above the existing 2,000-cell ordinary write budget.
Selection-relative navigation, native shortcuts, multi-view/lifecycle behavior,
all matching/formatting/protection variants and installed Windows/Linux packages
remain acceptance work. This checkpoint does not complete the XLSX/DOCX menu
ledger or the active full-desktop-menu goal.

The native VS Code 1.109.5 CDP scenario now also passes the shared Find/Replace
panel path. It clicks the real Home entry, selects formula-text lookup, types a
query, proves live-result navigation keeps the input focused, finds original A80
without changing canonical state/revision, replaces it through one host edit,
saves and checks the resulting worksheet XML plus preserved custom-XML sentinel,
then checks exact pre-replacement state after native Undo and the new value after
Redo. Prior AutoFit measurement/save/history and Go To assertions pass in the
same scenario. The final run exits 0 without the earlier cancellation/ENOENT
cleanup diagnostic. Earlier attempts exposed incorrect test roles/names (history
inputs are comboboxes; header/footer both expose Close), corrected against the
actual accessible UI before this successful run. This is one Linux VS Code
scenario, not evidence for installed Windows or all lifecycle/focus cases.

Final regression also checks query publication after the asynchronous indexed
scan and reopens a panel with an uncommitted query without searching its older
text. The complete embedded-editor unit suite now passes 14 tests in six files;
the complete workbook-session suite remains 102 passing tests in 30 files.
The final full extension build and desktop Sheets build pass, as do typechecks,
changed-source lint/format and production dependency license checks. Build size
warnings remain; these checks do not establish performance parity for large
workbooks or finish outstanding menu requirements.

## Native workbook shortcuts and cross-tab dialogs

The extension now registers native Find, Replace and Go To commands with scoped
Ctrl+F/H/G keybindings and Cmd variants. The contribution limits matching to the
workbook custom editor and excludes native input, terminal and Quick Open focus.
The provider additionally verifies the actual active custom tab URI/view type,
active panel and current session both before and after awaiting view readiness.
The browser accepts only a strict `view.command` enum: arbitrary command IDs,
URIs and operation payloads are not presentation requests.

Shared Home dialogs now render through a body portal, so a native command can
open the actual Find/Replace panel or Go To dialog while another ribbon panel is
selected. This retains locale context and the global appearance tokens without
switching the ribbon back to Home. Repeated Find focuses and selects the current
query through the shared service focus signal. Query research now honors the
desktop initial `stayIfOnMatch` behavior for an already selected matching cell;
other selection-relative and workbook-segment ordering cases remain an audit.

Evidence:

- A real VS Code 1.109.5 CDP test passes actual Control+F/H/G from the Insert
  ribbon, repeated Find text selection, Replace-tab switching and Go To Other.
  With original A1 and A80 both containing the query, selecting A80 first keeps
  the query on match two of two. The Insert tab stays selected throughout.
- The test rejects an arbitrary view command and a valid command with an extra
  URI. Canonical state and revision remain unchanged through these presentation
  operations. After opening an ordinary text editor, the workbook command
  returns false and actual Control+F opens the native Monaco find widget.
- The native shortcut scenario and the existing Home AutoFit/Go To/Find/Replace
  save, original-content preservation, exact Undo and Redo scenario pass together
  (two scenarios, exit 0). The final run has no prior cleanup diagnostic.
- The complete embedded-editor unit suite passes 15 tests, including the
  selected-match research anchor. All three desktop browser presentation probes
  plus the actual JSON bridge search test pass. Embedded Sheets, desktop Sheets
  and extension/DOCX typechecks pass; the full extension build passes with the
  existing chunk-size warning. Changed-source lint is clean.

This proves the exercised Linux shortcuts and dialog placement. It does not
prove Cmd+H behavior on macOS, installed Windows shortcuts, every locale or
appearance, chat/terminal/modal focus precedence, command dispatch during a
view-loading/disposal race, F5 variants, non-matching selection-relative search,
formula-value freshness or large atomic Replace All. Those acceptance items and
the complete XLSX/DOCX menu ledger remain open.

## Shared Insert ribbon and Link dialog checkpoint

The Insert panel now renders the actual desktop `SheetsRibbon` Insert branch,
including all ten groups and the original generated dropdowns. The single shared
adapter remains mounted outside the individual ribbon panels: its command band
portals into the selected Home/Insert anchor and its dialogs remain in the body
portal. Tab changes therefore keep dialog/query state instead of constructing a
second adapter. Source/current PivotTable output ranges provide the pivot hint
predicate. Existing simplified Insert controls remain in temporary additional
controls; this is not full Insert command acceptance.

`sheets-insert-ledger.md` records the desktop Insert command families and their
generated option sources. All rows except the connected Link editing checkpoint
remain open; rendering a button is not evidence that its command works.

The desktop's actual Link dialog and target normalizer are extracted as pure
reusable modules. Desktop callers keep their synchronous command behavior;
plugin callers await the host before closing, lock pending actions and leave the
dialog open after a rejected edit. URL normalization, internal-address syntax,
labels, layout and remove/cancel controls come from the desktop source. Edits use
the frozen top-left selected cell and revision/state guard, with a final check
inside the ordinary operation queue. Link target changes use the existing
validated `set_hyperlink` transaction, package writer and shared native history.

`source.read` additionally returns only current in-range hyperlink targets.
Original coordinates map through structural edits and canonical overrides and
removals replace source entries. Opening Edit Link reads that authoritative
current target rather than assuming the grid's cell style contains a URL. The
read-current unit case verifies clipping, overrides, removal, coordinate shifts
and immutable input state.

Evidence before the final reopen assertion:

- The complete workbook-session suite passes 103 tests in 30 files; the complete
  embedded-editor suite passes 15 tests in six files.
- The Link presentation joins Sheets, Docs and Find/Replace in the pure-browser
  dependency boundary probes; together with the actual JSON search bridge test,
  five scenarios pass. Embedded, host/extension/DOCX and desktop typechecks pass.
- The real VS Code Link scenario proves source-link prefill, cancel without an
  edit, rejection of an invalid target, bare-domain normalization, current edited
  target prefill, removal, internal-link creation, and rejection of a dialog
  overtaken by a concurrent edit. Saved external relationship targets and native
  internal `location` XML match the requested values. Worksheet `sheetData` is
  byte-for-byte identical after the link-only save; the original custom-XML
  sentinel survives. Native Undo restores the exact pre-edit state and Redo the
  exact accepted state.
- Link, native shortcut and Home AutoFit/Go To/Find/Replace save/history scenarios
  pass together (three scenarios, exit 0). An earlier Link test reverted and
  immediately closed/deleted its fixture, producing an ENOENT diagnostic during
  asynchronous reload. Saving the accepted fixture state before closing avoids
  that reload/delete race; the subsequent complete run has no such diagnostic.
- Full plugin and desktop Sheets builds pass with the existing chunk-size
  warning. Changed-source lint reports only the pre-existing ExcelShell hook
  dependency warning, with no errors.

Link activation/navigation and grid decoration, every existing target form and
length boundary, all locales/appearances, PivotTable/table/chart/media/note/text/
symbol insertion, installed Windows/macOS behavior and the remaining tab/context
inventories remain open. This checkpoint does not complete Insert or the overall
XLSX/DOCX desktop-menu objective.

The final Link-only native run additionally closes the custom editor, waits for
the original document to dispose, opens the file in a distinct new session, and
checks that no canonical hyperlink overlay exists at startup. The shared dialog
then reads `#Other!A1` from the saved source. The grid still presents numeric 12.5
at A2; the reopened descriptor stays at revision zero and is not dirty after
navigation/cancel. This scenario passes with exit 0. It prints a bare `Canceled`
diagnostic during lifecycle activity, so clean cancellation/disposal acceptance
remains open; the successful save/reopen assertions do not resolve that diagnostic.

The final 19-scenario browser regression passes after the Insert rearrangement.
It retains all existing format, structure, size/drag, merge, formula, appearance,
chart and navigation assertions. The chart test now expands Insert's temporary
additional controls; two older global summary selectors were scoped to Home so
hidden Insert controls cannot create an ambiguous match. Explicit preoptimization
of the shared find/RxJS/React DOM imports keeps the two browser fixtures stable.
The Link extraction also preserves desktop's existing empty-address Enter-close
path; plugin async apply continues requiring a valid accepted edit.

## Insert Symbol checkpoint

The plugin opens the actual desktop `SymbolDialog` from the shared Insert ribbon.
The same categories, characters, localization, modal contract and `appendSymbol`
helper remain in desktop use. Async hosted inserts lock pending symbol/category/
close actions and Escape, retain the dialog after an accepted click, and report
errors without an unhandled rejected Promise. Synchronous desktop callbacks keep
their immediate behavior.

The adapter reads only the frozen selection's top-left cell from the host. It
checks state/revision before and after the source read and again inside the
ordinary operation queue. It uses `covertCellValue`, the actual desktop facade
conversion: appending a currency sign may preserve a number and change its
number format, and formula-looking output becomes a formula. Value and format
writes share one transaction/history action. Subsequent clicks adopt only the
committed receipt's state/revision, never whatever newer state happens to be
installed when the Promise resolves. Explicit null cached values remain blank;
pending formula results and final text exceeding 32,767 UTF-16 units reject before a
write. General calculation freshness is still incomplete.

Evidence:

- Three intent tests cover original leading-zero text, numeric/boolean/blank
  values, actual facade currency parsing with its number format, formula-looking
  output, ready/error/explicit-null caches, and pending/oversized rejections.
  The complete embedded-editor suite passes 18 tests in seven files. The existing
  three desktop append-helper tests pass unchanged.
- The browser Symbol scenario holds a source read and proves disabled duplicate
  clicks/Escape cannot commit or close the panel, then accepts exactly one edit.
  A second symbol appends to that accepted value, only the selected range anchor
  changes, and a later concurrent note invalidates the old dialog. All 20 browser
  regressions pass. Handled tab keys now stop event propagation, and the keyboard regression
  also checks that the selected grid range does not move. A repeated run exposed
  intermittent focus loss; focused-page preconditions and failure diagnostics
  are retained rather than attributing a cause from a single passing retry.
- The native VS Code Symbol scenario passes original source ready-formula,
  number, boolean, date, leading-zero text and blank-range cases, and facade
  conversion of formula-looking text. Each primitive case returns to the exact
  initial state through native Undo. A formula with no usable cache is explicitly
  refused without advancing revision. The fixture restores B1's independently
  known 12.5 * 2 = 25 cache after the workbook writer invalidates caches; C1
  deliberately remains uncached, so the test does not mask that limitation.
- The native scenario holds submission before commit and verifies no optimistic
  state change, duplicate edit or modal close. Currency and the next ordinary
  symbol use two accepted receipts and exactly two Undo/Redo actions; the original
  bold/fill remain. It also holds an acknowledgment after committing a source A80
  insertion, commits a different note edit, then delivers the older receipt. The
  next Symbol click rejects instead of adopting the unseen concurrent state.
- Saved XML includes the authored symbols and preserves the custom-XML sentinel.
  Closing the old custom document and opening a distinct session reads the
  expected A2 and A80 text back; original A2 bold/fill survive and navigation leaves
  the reopened revision at zero with no dirty state. The run exits 0.
- All five desktop presentation boundary probes plus the actual JSON search
  bridge test pass. Host/embedded/desktop typechecks pass, changed-source lint is
  clean, and full plugin and desktop builds pass with existing size warnings.
  Desktop typechecking also caught an omitted local `normalizeLinkTarget` import
  after the preceding extraction; the import is restored while retaining its
  public re-export.

This checkpoint does not establish full Symbol or Insert acceptance. Pending or
stale formula results still need a verified calculation path, including cache
validity across save/reopen and style-only edits; current read flags are
conservative. Every palette/input-conversion case, locale, appearance, modal
shortcut priority and installed Windows/macOS behavior remain an audit. Other
Insert commands, remaining tabs/contextual menus and the complete DOCX ribbon
are still open under the unchanged full-menu objective.

The focused keyboard scenario additionally passes three repetitions after making
page activation and initial Home focus explicit preconditions. Insert focus and
unchanged grid selection remain required. Five diagnostic repetitions had also
passed after an earlier intermittent failure; this evidence does not establish
every focus/lifecycle case, so broader keyboard acceptance remains open.

Symbol size checks apply to the final converted text, preserving desktop numeric
conversion even when a long zero-padded currency input converts to a short scalar.
A missing requested source cell rejects rather than being inferred to be blank;
real blank cells are explicit records. No additional record or transaction is
created before these validations complete.

The final native run also includes selecting B6 inside the original A5:B6 merge.
The host supplies its content anchor, so Symbol appends to A5 and does not author
B6. Exact Undo restores the original state and source geometry. Symbol, Link,
native shortcuts and Home save/history scenarios all pass together (four
scenarios, exit 0), including distinct-session reopen checks. That run has no
previous cancellation/ENOENT diagnostic; broader lifecycle acceptance remains
open rather than being inferred from a clean sample.

Symbol's source-read and commit stages are tracked as one pending view operation.
Public editor flushing waits for it, while the Symbol operation itself uses the
cell-editor-only flush to avoid a self-wait. The browser explicitly holds the
source read, starts a flush, proves the flush remains unresolved, then releases
the read and verifies a ready result only after the accepted edit. The final
20-scenario browser run passes with this save barrier in place.

The final combined native run passes all four scenarios after adding the Symbol
flush barrier. During a deliberately held submission, native Save remains pending;
after release it completes without error, marks the accepted root saved and writes
the new currency format to `styles.xml`. Repeated Symbol insertion, exact Undo/Redo,
receipt/concurrent-edit rejection, merged-anchor handling and distinct-session
reopen checks still pass. This run prints a bare `Canceled` during Symbol lifecycle
activity, so the prior clean sample is not treated as resolution of the existing
cancellation/disposal audit. The test exits 0 and all persisted-content assertions
pass. The final full plugin build, typechecks, 18 editor unit tests and 20 browser
regressions also pass; desktop build/typecheck and six presentation/bridge boundary
scenarios pass for the unchanged corresponding sources.

## Insert Header & Footer shared-dialog checkpoint (2026-10-25)

The plugin now renders the actual desktop HeaderFooterDialog. Its six sections,
field-code buttons and cursor insertion behavior are shared with desktop. The
header/footer decoder was extracted into a presentation-independent module and
re-exported at its original desktop entry point. Opening reads the captured
sheet's source odd sections and overlays canonical pageSetupEdits; an explicit
null remains cleared. Applying creates one set_page_setup transaction bound to
the opening state/revision and sheet. The existing surgical writer retains
first/even variants and headerFooter attributes.

Hosted apply supports asynchronous success/failure while preserving the desktop
synchronous callback contract. Pending work disables fields, field buttons and
actions and guards Escape/backdrop close. Parameter validation happens before
queueing: the browser initially exposed invalid input poisoning later editor
flushes; moving schema validation ahead of enqueue permits correction and retry.
A stale opening state is also rejected before queueing and checked within the
queue again. This does not claim a global resolution of all rejected-operation
flush behavior elsewhere.

The focused browser scenario and full 21-scenario Home/manual-edit regression
pass, covering original prefill, cancel, cursor field insertion, section-length
failure and retry, explicit clearing and stale-dialog rejection without cell
changes. Existing desktop print/page-setup tests pass (99 tests in two files),
editor unit tests pass (18 in seven files), and browser presentation/search
boundary probes pass (seven in two files), including the new Header/Footer
presentation probe. Extension, embedded editor and desktop typechecks, plugin
production build, ESLint, formatting and diff whitespace checks pass. Build
emits the existing bundle-size warning.

Native save/history/reopen verification is recorded below. Full
print output/format-code fidelity, all language and appearance variants,
installed Windows/macOS testing and broader modal/lifecycle acceptance remain
open; this checkpoint does not complete Insert or the feature.

The actual Linux VS Code 1.109.5 native run passes all three scenarios: Header &
Footer, Link and Symbol (exit 0). Header & Footer starts from source odd sections
plus different-first/even flags and variant text. Holding the real engine submit
proves pending inputs/actions are disabled, duplicate clicks and Escape cannot
close or optimistically mutate the document. Exact native Undo/Redo restores the
accepted roots. Saved XML retains first/even content and flags, preserves the
entire source sheetData and custom sentinel, writes literal-dollar/ampersand text
and removes a cleared odd footer. A distinct reopened session reads the saved
section values with revision zero and a clean descriptor.

The run includes two bare Canceled messages during the Symbol scenario and a
transient unresponsive extension-host startup that recovers. These diagnostics
remain part of the existing lifecycle/performance audit; passing content tests
does not establish their resolution. No Windows VSIX was produced and the
workspace changes have not been committed or pushed in this checkpoint.

## Insert chart presentation and authoring checkpoint (2026-10-25)

The desktop chart presentation now lives in SheetChart.tsx without workbook
media, document or editor service imports. WorkbookVisuals retains privileged
visual installation, hydration, inline editing, selection stores and context
menus; it uses the extracted body and re-exports the prior rendering helpers.
The SVG plot implementations, authored colors, axis layout/text, labels,
legends and element-selection callbacks are shared. The plugin's independent
four-type BasicChart has been removed. Its current chart preview pane renders
the actual desktop component for all nine ordinary Insert families.

Desktop manual insertion and authoritative add_chart now share buildChartVisual.
Numeric-year vectors become categories, scatter uses its numeric category vector
as X, references remain absolute and quote sheet names, and default anchors match
desktop. The existing desktop shorter-dimension orientation heuristic remains;
square test data was initially incorrectly assumed to be column-oriented, so
the final source fixtures use four rows by three columns. Newly inserted charts
retain desktop's absent data-label setting instead of the former gateway-only
forced value labels. Existing authored labels are unchanged.

The session's four-type gate has been removed. The package writer recognizes
exact supported plot sets (including bar+line combo) and routes all nine through
the existing drawing serializer, rejecting unknown mixed sets instead of guessing
a different type. Original file chart parts remain preserved. Opening a saved
file now displays original source chart visuals using their metadata and the
logical sheet mapping; it does not promote source charts into duplicate additions.

Every ordinary Insert gallery entry is connected. Recommended Charts uses the
actual desktop ranked previews/dialog. Opening and applying bind to the captured
sheet/state/revision. Selection reads are bounded at the desktop 2,000-cell limit,
read the complete source selection beyond the visible window and reject missing
cells or pending formula values. Authoring is prevalidated before enqueueing and
rechecked within the queue. The entire read/commit operation participates in
editor flush, so Save waits for accepted chart work. The recommendation dialog
retains synchronous desktop callbacks and supports hosted pending/error behavior;
duplicate clicks, Escape and backdrop/Close cannot cancel a pending submission.

Evidence: the focused new insertion/recommendation browser scenarios pass; the
full existing Home/manual suite passes 23 scenarios with shared rendering. The
native Linux VS Code 1.109.5 scenario clicks all nine actual gallery buttons,
checks source-derived caches and exact absolute references, restores exact roots
with independent Undo/Redo after each insertion, rejects a stale recommendation,
saves all nine plot families and reopens a distinct session with nine rendered
charts and a clean revision-zero descriptor. Saved worksheet sheetData and the
custom XML sentinel remain unchanged. The sidecar-backed repeated-save regression
now covers all nine types and verifies two separately saved files each contain
exactly nine charts and the expected current unsaved values.

The session suite passes 112 tests in 30 files; editor units pass 18 tests in seven
files; isolated browser presentation/search boundary tests pass nine tests in two
files, including chart and recommendation probes. Existing desktop chart/visual
regressions pass 147 tests in 17 files and the numeric-year/scatter/plot-set model
checks pass alongside recommendations (12 tests in two files). Desktop, embedded
editor and extension typechecks pass. Both plugin and desktop production builds
pass, with the existing bundle-size warnings. Formatting, lint and whitespace
checks pass.

This is partial chart parity. The plugin still uses a side preview pane rather
than desktop in-grid anchors, dragging and resizing. Source charts display but
cannot yet be edited/deleted through session operations; source chart cache
hydration/refresh, full contextual Chart Design/format menus, every chart option,
formula-result freshness, cross-platform installed behavior, appearance/locale
and lifecycle audits remain open. The native run prints bare Canceled diagnostics
while replacing views for history/reopen; the existing cancellation audit remains
unresolved. No Windows VSIX, commit or push was produced in this checkpoint.

The final error-path browser scenario rejects nonnumeric selections, a source
formula explicitly marked pending and a 2,020-cell range before reading that
range, then accepts a valid chart insertion. Its first oversized fixture selection
exceeded the display's allocated columns before reaching the chart command; the
fixture range was corrected to A1:T101 and the focused scenario passes.

The expanded final regression exposed the previously recorded tab-focus race.
Failure-only instrumentation captured actual calls from EditorBridgeService's
setEditCell through normal DocSelectionRenderService.focus after the Insert tab
had already gained focus. Initial selection synchronization can do the same
before the arrow key, causing that key to move the grid selection instead.
React bubbling stopPropagation alone cannot prevent those engine focus calls.
The diagnostic run reproduced two failures in five repetitions; its temporary
global focus instrumentation was removed after collecting the stacks.

ribbon-focus-boundary.ts now guards only the current engine instance's normal
selection renderer while this workbook's ribbon tab owns DOM focus. Selection
and geometry updates continue normally. A real pointer action releases tab
ownership before engine handlers; renderer/runtime disposal restores the original
method and unsubscribes listeners. Explicit Go To grid-focus requests release the
boundary before layout focus, preserving the intentional return from a dialog.
This is scoped to tab ownership and does not claim to solve every input, modal,
IME, platform or lifecycle focus path.

The original strict seven-tab scenario passes five consecutive repetitions after
the boundary. A separate deterministic scenario selects B2 while Home owns focus,
checks tab and selection preservation through ArrowRight, then clicks the actual
canvas and types 77 successfully into B2. The final combined browser suite passes
all 25 scenarios (58.6 seconds). The rebuilt native extension passes both shortcut
and nine-chart scenarios together (two scenarios, exit 0), including an added DOM
focus assertion after native Go To from Insert. This final native sample prints
no Canceled message; the earlier diagnostics remain an unresolved broader audit.
Final plugin build, embedded-editor and extension typechecks, lint and whitespace
checks pass with the focus boundary. The desktop build/typecheck covers the shared
chart extraction; subsequent focus changes are plugin-only. Changes remain in the
workspace without a new commit, push or VSIX package.

## Floating charts and original drawing transactions (2026-10-26)

The plugin now uses the desktop installWorkbookVisuals implementation and drawing
preset to place charts in the worksheet. It reuses the actual frame selection,
move ghost, eight resize handles, marker/EMU conversion, pointer capture and delete
affordances. Sidebar SVG copies are removed; temporary property forms are collapsed.
The visual module now imports browser chrome instead of the AI UI barrel. Its
browser dependency probe passes without AI or a second grid/formula service. Media
reads use a narrow optional port and a capability-checked desktop fallback; the
embedded chart-only path performs no media request.

set_visual_anchor is a shared validated layout operation with positive frame size
and grid bounds. Session chart anchors update in canonical visual state. Original
chart drawing anchors and removal tombstones live in visualEdits keyed by original
visual id. Both candidate and writer resolve original sheet/drawing ownership from
immutable metadata and validate the trusted locator. Browser operations supply
sheet/visual ids and anchor data, never a package path/index. Protection, unknown
or removed ids and invalid anchors reject without publishing a partial root. The
desktop executor also routes this operation through its existing shape edit port.

Projection installation waits for an active gesture to end. Gesture callbacks bind
to the installed state/revision; an overtaking transaction rejects the old anchor.
Flushing waits for drag completion and accepted visual work. The shared frame now
cancels on Escape in the window capture phase, pointer-capture loss or window blur.
The capture phase is necessary because engine focus can otherwise keep Escape from
reaching the frame's React handler. A new test initially exposed flushing remaining
blocked after Escape; the capture-phase cancellation fixes that case. These changes
do not establish complete cross-Webview pointer/lifecycle acceptance.

Native diagnosis exposed VS Code's default 20px body side padding: the stylesheet
set body width to 100% without clearing that padding, so the measured grid extended
past the visible iframe. A screenshot and hit-test showed a resize handle at x798
outside the 792px viewport. The Webview stylesheet now clears padding. The native
geometry scenario uses an in-viewport NW handle and verifies the original-size
chart rather than clicking a clipped SE handle. An earlier cross-iframe drag/flush
timeout remains part of the pointer lifecycle audit; all supported cancellation
paths and broader cross-platform behavior still need continued verification.

The final browser run passes all 29 scenarios, including original source move,
resize and deletion, unchanged cells/metadata, a held Save flush, a stale concurrent
gesture, Escape cancellation and mixed original/session Shift-selection deletion
as one transaction. Source anchors remain immutable and only canonical patches
change. Existing menu/format/theme/history projection cases continue to pass.

The final Linux VS Code 1.109.5 run passes three scenarios together: native shortcuts,
all nine chart insertions, and original floating charts (exit 0). The source-chart
scenario verifies exact Undo/Redo after moving/resizing, saved marker offsets,
byte-identical original chart XML and worksheet XML for position-only edits, and
the custom XML sentinel. It closes the original document, opens a distinct clean
revision-zero session, checks every saved anchor field and the displayed geometry,
then deletes the chart with native history. The saved chart part is removed. If
its now-empty drawing part is removed too, the worksheet drawing reference, its
relationship and content-type override are absent. The test initially assumed an
empty drawing part survived and failed by dereferencing null; it now checks the
actual required no-dangling-reference behavior. The chart's content-type override
is absent and unrelated sentinel content remains intact.

Additional checks pass: 115 session tests in 30 files (including original drawing
ownership/protection and atomic anchor rejection), 18 editor unit tests in seven
files, ten presentation/search boundary tests in two files, and 221 desktop
chart/visual/drawing regressions in 24 files. Desktop, embedded-editor and extension
typechecks pass. Both final plugin and desktop production builds pass with existing
bundle-size warnings. ESLint and whitespace checks pass. The native combined run
prints bare Canceled messages during history/reopen; their broader lifecycle audit
remains unresolved despite passing all content assertions.

The final browser run was initially not executed because automatic approval review
hit a usage limit. After the limit reset the normal approval path succeeded; the
29-scenario run above completed without bypassing approval. No new commit, push or
installed Windows/Linux VSIX was produced. Full Chart Design/context menus, actual
property panes, source chart content/cache editing, RTL/zoom/freeze/large-anchor
and installed-platform coverage remain open, as do the other menus and DOCX scope.

### Shared inline chart properties (2026-10-26)

The plugin now renders the actual desktop inline `ChartEditor` for session-added
charts. The temporary plugin-specific property forms are removed. Original-file
chart property editing remains gated off; their existing movement/deletion support
is unchanged. The shared panel supports title, the desktop's six convertible types,
series colors, and pie/doughnut point colors. Advanced plot families retain their
plot type when only their title or colors change. No range reader or unconnected
context menu is enabled by this checkpoint.

Hosted apply accepts a promise, disables the fieldset/cancel controls, prevents
repeated submissions synchronously, and retains the draft/error after a rejection.
Escape cancels an idle draft without a transaction. An applying editor releases
its installation guard so its own committed projection can replace it; otherwise
the guard preserves drafts while an unrelated update is waiting. A frozen root and
revision plus the host transaction check reject overtaken drafts. The complete
chart commit participates in `view.flush`. A failed chart dialog transaction is
handled by its own editable error state rather than permanently preventing retry;
failed cell-editor commands still retain their save protection.

`edit_chart` now validates authored point-color maps as well as series colors, with
matching desktop AI and in-memory adapters and regenerated tool schemas. Native
sidecar save/reopen exposed a real serialization omission: the session package
writer discarded both series colors and point colors, rebuilding palette defaults.
It now forwards these authored colors through the existing chart-add writer. Two
successive saves/reopens confirm the pie point override survives alongside all nine
chart families and their source data. The saved pie includes default colors for the
other points, so the check verifies the authored point inside the full parsed map.

Verification:

- Final browser regression: **31 passed**. This includes actual inline title/type
  changes, failure retention/retry, held commit/Save flushing, duplicate clicks,
  pending Escape/cancel protection, idle Escape cancellation, pie palette selection,
  and an overtaken draft alongside all prior Home/Insert/floating scenarios.
- Real VS Code shared-panel scenario: **1 passed** after rebuilding native test
  artifacts. Actual pointer clicks edit the topmost combo chart's title and first
  series color; full canonical roots match native Undo/Redo; saved chart XML carries
  both properties; the untouched worksheet data/sentinel survive; a distinct clean
  reopened session renders the saved title and has no original-chart property pencil.
- The first native run also passed floating-original-chart and workbook-shortcut
  regressions, while the new property scenario failed because its right-edge pencil
  was clipped outside the canvas viewport. The final test uses native Go To to move
  past the chart and then reveal its middle, verifies `elementFromPoint` hits the
  actual pencil, and restores the source selection before Recommended Charts.
  No forced click, smaller fixture, or synthetic property action substitutes for
  the pointer interaction. The final property scenario passed in **33 seconds**.
- Workbook session suite: **116 tests / 30 files passed**, including atomic point
  color edits, preservation of cells/data, invalid-color rejection and merged overrides.
- Editor unit suite: **18 tests / 7 files passed**. Browser boundary suite: **9 passed**.
- All-nine-family repeated-save sidecar test: **1 passed** with the authored pie
  point color retained in both saved files.
- Workbook session, editor, extension/DOCX and desktop TypeScript checks pass.
  Plugin production build (including generated schema and boundaries) passes;
  desktop main/preload/renderer build passes. Changed-source lint and formatting
  checks pass, as does `git diff --check`.

Full contextual Chart Design/Format menus and panes, original-file chart property
transactions, source/cache/range authoring, the other chart options, broad locale
and appearance verification, cross-Webview pointer acceptance, all remaining Sheets
and Docs menu families, and installed Windows/macOS acceptance remain open. The bare
`Canceled` diagnostic observed in earlier native history/disposal scenarios remains
on the lifecycle audit; a passing property scenario does not close that audit.
The new capability/busy rejections reuse the existing localized chart-not-editable
message; no new locale keys or authored-color theme remapping were introduced.
No commit, push or VSIX packaging was performed in this checkpoint.

### Original-file chart property editing (2026-10-26)

The actual desktop inline panel now edits both original-file and session-added
charts. Original chart parts resolve against trusted global metadata and their
logical owner; unknown or multiply-owned chart paths, protected/unverified sheets,
removed visuals, and nonexistent series/points are rejected before publication.
Conversion accepts the same single-plot families as the desktop. Axis/grouping
incompatibilities fail before publishing a root. Original series-range editing is
still unconnected and is rejected explicitly.

Source property transactions retain validated `workbookChartEditSchema` overlays
in the owner sheet's `chartEdits`, including nested merges for authored series and
point colors. The raw chart metadata and worksheet cells stay immutable. The grid
composes these overlays through the shared chart state helper, so the title, type,
colors and reopened inline controls echo the accepted canonical state. Save resolves
the trusted owner again, patches the original chart part surgically, and skips a
property overlay when its drawing was deleted. Original charts are never duplicated
as chart additions. `SheetVisual` now records the optional original `chartPath`.

Real VS Code revealed an Electron-specific interaction failure: selecting Line
left the custom type listbox visible, blocking the series color button. The shared
chart editor previously put its multi-button Dropdown inside an implicit HTML
label. Replacing that wrapper with the same styled ordinary field container keeps
the Dropdown's existing accessible name and eliminates the label activation. Native
verification first reproduced a visible listbox after selection, then confirmed it
closed and the color picker became clickable after this shared fix. Browser
verification also asserts the listbox disappears.

Verification so far:

- Final native run: **3 passing, 50 seconds**, exit zero. Includes workbook
  shortcuts; all nine new chart families with property save/reopen; and original
  drawing movement, resizing, properties, agent-origin edits and cascading deletion.
- The original-chart scenario modifies the actual shared title/type/color controls,
  checks exact native Undo/Redo roots, saves a line chart with the new title and
  authored red series, and proves drawing/worksheet XML remain byte-identical during
  the property-only save. A distinct clean session reads the new type/title/color.
  An `origin: agent` engine transaction then changes the original title, is rendered,
  and returns to the exact pre-agent state through native Undo. This proves the
  shared canonical history path, not an end-to-end language-model/chat invocation.
  Final Delete/Undo/Redo/save removes the chart without dangling drawing relationships.
- Workbook session suite: **118 tests / 30 files pass** on the final isolated run.
  New cases cover nested source overlays, immutable metadata/cells, atomic bad-series
  rejection, unknown paths, protected/removed charts, unsupported axes and ambiguous
  ownership. An earlier concurrent full run timed out only in the heavy history
  storage case; that case passed alone in 4.64 seconds and the full suite passed
  separately in 15.08 seconds without changing limits or the implementation.
- The first native shortcut run asserted focus before the asynchronous command
  completed. It now waits for the actual input focus before checking the same full
  selected query; the isolated and final complete native runs pass. Bare `Canceled`
  diagnostics still occur in history/reopen flows and remain an open lifecycle audit.
- Final complete browser regression after the HTML container fix: **32 passed**
  in 1.7 minutes, with failure tracing enabled. One preceding final run timed out
  in the original-chart interaction without preserving its actionable wait stack;
  its isolated trace run and the subsequent complete run passed without changing
  product code. The test now uses bounded per-action waits and non-masking cleanup.
  This intermittent browser wait remains on the interaction/lifecycle audit; the
  successful full run does not prove universal focus/timing acceptance.
- Workbook session, embedded editor, extension/DOCX and desktop TypeScript checks
  pass. Plugin build/schema/boundary checks pass after the shared fix; changed-source
  lint and formatting are clean. Final desktop main/preload/renderer build passes
  in 32.69 seconds, and `git diff --check` passes.

Full source reference/cache hydration and data-range editing, multi-owner chart
parts, contextual Chart Design/Format menus and panes, AI tool discovery/composition
of current chart descriptors, all remaining chart parameters, appearance/locale
coverage and installed platform checks remain open alongside the other Sheets and
Docs menu families. No commit, push or VSIX packaging was performed.

### Chart series ranges, save barrier and missing-cache hydration (2026-10-26)

The actual desktop inline range inputs now work for original-file and session-added
charts. A shared bounded-vector resolver accepts row/column cell vectors (up to
1,000 cells), quotes/escapes the canonical sheet name and retains absolute cell
references. Values and categories can come from different worksheets, including
quoted names containing `!` or apostrophes. Same-series entries merge before applying.
The renderer reads for interaction but converts the result back into canonical
`seriesData` references; the host hydrates the exact referenced cells and derives
caches again. Browser-provided numeric/category arrays are not transaction authority.
Non-vector/oversized references, incomplete material, pending formula caches and
outdated submissions fail atomically.

Session series edits now retain their new references, fixing the old refresh path
that could overwrite a just-edited cache from its former range. Explicit series
names clear their former name reference, and numeric text uses the same finite
coercion in cache refresh and the range editor. Original-chart series overlays merge
with earlier properties and hydrate their exact references during later operations;
scalar edits refresh these journaled caches without mutating raw source metadata.
This does not yet implement live refresh for every untouched or reopened original
chart reference, nor prove general formula freshness after arbitrary dependencies.

The inline-apply counter covers range reads as well as commit/receipt, and releases
idempotently on completion or component disposal. Editor flushing waits for this
whole interval; the commit still uses the private cell-editor flush to avoid waiting
on itself. Controls remain disabled during reads, duplicate Apply does not create a
second transaction, and pending Escape/Cancel cannot abandon the apply. Background
missing-cache hydration uses the read port without dirtying or authoring the document.

Verified so far:

- Final native run: **3 passing in 36 seconds**, exit zero. Original-chart editing
  points values at Other!A1 and categories at Data!A1 through actual controls. The
  source read is held before any transaction exists on a clean revision-zero
  document; native Save remains pending, repeated Apply/Escape cannot submit or
  close, and Save completes cleanly after release. Exact native Undo/Redo roots
  match the accepted range/property edit. Saved XML contains both absolute refs,
  numeric cache 7 and category `00123`; the drawing and Data worksheet XML stay
  byte-identical. A distinct new session reads both refs/caches, title/type/color,
  and remains revision zero/clean. Agent-origin history and cascading deletion still
  pass alongside all-nine-family insertion/property and workbook shortcut scenarios.
- Final workbook session suite: **121 tests / 30 files passed** in 12.48 seconds.
  New tests cover different source sheets with quoted names, merged source overlays,
  later scalar-cache refresh, atomic invalid/oversized/pending-formula rejection,
  persisted session references, numeric-text values and explicit-name precedence.
- Targeted browser tests pass for held range reads/Save, input correction after
  invalid/pending data, duplicate and close protection, later journaled-cache refresh,
  stale read rejection, and missing original values/categories/name hydration. The
  hydration case verifies source caches remain empty, the original placeholder name
  remains unchanged, no chart overlay is authored and the revision stays zero.
- Plugin production build/schema/boundary checks pass. Embedded editor and
  extension/DOCX TypeScript checks pass; changed-source lint/formatting and
  `git diff --check` pass. The final complete browser regression is **35 passed**
  in 1.8 minutes with failure tracing enabled. Final desktop main/preload/renderer
  build passes in 38.59 seconds; desktop/session TypeScript checks, editor unit
  tests (**18 / 7 files**) and browser boundary/search bridge tests (**10 / 2 files**)
  pass. The desktop strict optional-property check required adding explicit
  `undefined` to the shared range-entry optional types; this final change is type
  only and the emitted behavior is unchanged.

Broad formula-result freshness, original/live/reopened/imported reference refresh,
all typed/date/error/name-range cases, full Select Data and contextual Chart Design/
Format menus/panes/options, current chart discovery in AI tools and actual model/chat
flows, appearance/locale/platform audits and the previously observed intermittent
focus/cancellation diagnostics remain open. No commit, push or VSIX packaging was
performed in this checkpoint.

### Desktop chart context menu and format pane (2026-10-26)

The plugin now reuses the actual desktop context menu and ChartFormatPane. Format
Selection opens a persistent pane outside the grid; Add/Delete Data Labels and
Delete Chart use canonical transactions. Select Data, alt text, arrange and size
entries are visibly disabled pending their implementation. This is partial context
integration, not full Chart Design/Format ribbon acceptance.

The shared DSL now accepts the pane's label placement/number format, gridlines,
value-axis bounds, gap width, doughnut hole size and series/point explosions, in
addition to title, colors, legend and labels. The desktop AI adapter and both
source/session candidate routes carry the same fields. The session writer now
preserves these settings instead of silently dropping them; source saves continue
using surgical chart XML patches. Effective partial bounds reject min >= max,
family-incompatible settings fail before publication, and conversions remove
settings for XML elements that cease to exist. Point validation includes trusted
reference extents for uncached original charts without authoring their caches.

The hosted pane disables controls/close while pending and serializes edits. Pending
checkbox/dropdown proposals echo locally; text/number inputs retain their keyed DOM
drafts on rejected writes. Editor flush blurs active pane inputs and waits for their
transaction, so Save publishes a focused text draft. Hidden editor synchronization
cannot steal focus from active pane controls. The axis-range message is translated
in all 21 locale shards; no new authored-color theme mapping is introduced.

Real pointer testing exposed capture retargeting: the float frame received clicks
intended for SVG slices, preventing element-scoped controls. A non-drag captured
press now delegates its click to the original hit element, using existing chart
handlers. Moves/resizes retain their established capture/ghost path. Custom dropdown
and color-well rows use ordinary field containers with explicit accessible names,
avoiding implicit-label reactivation of multi-button controls.

Verification completed before the final production rebuild:

- Complete browser regression: **38 passed / 2.4 minutes**, including actual context
  actions, parameter echoes/retry, active-input Save, pending close/flush, rejected
  draft retention, keyboard sliders producing one transaction each, real SVG slice
  hit-testing/pointer clicks and point-scoped fill/explosion, plus prior drag/resize,
  multi-delete, cancel and Home/Insert regressions.
- Workbook session suite: **124 tests / 30 files pass**. Includes original/session
  format fields, merged partial axis validation, invalid-family/point refusal,
  conversion cleanup and uncached point extent validation with immutable metadata.
- Repeated native sidecar save/reopen of all nine session families passes with
  column legend/labels/position/format/gridlines/bounds/gap and doughnut hole,
  series explosion and point explosion preserved. The presentation boundary suite
  now includes the complete ChartFormatPane entry without embedded AI/editor services.
- Native VS Code initial format run: **3 passed / 37 seconds**. Actual source-chart
  context actions and pane controls preserve exact Undo/Redo roots; real keyboard
  typing retains pane focus, native Save commits the still-focused title, and saved
  XML/reopened metadata retain title, legend, labels/format, gridline removal and
  axis bounds. Drawing/worksheet bytes remain unchanged. The final rebuild/regression
  after the element-click/draft/localization refinements is recorded below.

Full Select Data replacement and row/column switching, contextual ribbon bands,
remaining visual commands, noncanonical index-key normalization, every plot/input/
parameter edge case, general live-cache/formula freshness, end-to-end model/chat
flows, all-locale/appearance interaction and installed-platform verification remain
open. Earlier intermittent focus/cancellation observations also remain on the audit.
No commit, push or VSIX packaging was performed.

Final format-pane checkpoint verification:

- Final plugin build/schema/boundary checks pass (39.00 seconds for Vite).
- Final native VS Code run: **3 passing / 41 seconds**, exit zero, on the rebuilt
  element-click/draft/localization implementation. Source format input, focus,
  native Save/history/reopen and original drag/resize/deletion scenarios all pass.
  One bare `Canceled` remains during close/reopen, so the lifecycle audit stays open.
- Editor unit suite: **18 tests / 7 files pass**. Browser presentation/search bridge
  boundary suite: **11 tests / 2 files pass**, including ChartFormatPane.
- Final extension/DOCX and desktop TypeScript, changed-source lint, all 21 locale
  shard/changed-source formatting and `git diff --check` pass.
- The final shared desktop main/preload/renderer build passes in **33.98 seconds**.

## Select Data dialog checkpoint (2026-10-26)

The plugin now mounts the actual shared desktop SelectDataDialog from the chart
context menu for original and session-added charts. Rename, add, remove and reorder
series commit as one canonical `edit_chart.seriesSelection` transaction. Switch
Row/Column transposes the current authoritative data and replaces references with
literal caches, matching the shared desktop behavior. Single-plot column, bar,
line, area, pie, scatter, radar and doughnut charts are supported; combo replacement
remains explicitly unavailable through the shared dialog.

The browser submits source indices, authored names and optional range references,
never replacement point arrays. Material preparation resolves retained and new
references; candidate execution checks coverage, protection and formula readiness.
Original overlays are composed before deriving a new series set, preserve colors
with the reordered series and discard obsolete index/point overrides. Subsequent
inline edits use the effective series count; subsequent cell edits refresh retained
references. Source metadata remains unchanged. A parser-only field leak during
replacement-cache consolidation was found by regression and fixed.

The dialog captures its opening root/revision. Full asynchronous commits are awaited
before closing, inputs and cancellation are disabled while pending, repeated clicks
are guarded synchronously, rejected drafts remain visible, and editor flushing waits
for the chart operation. The desktop fallback callback is also awaited. Scatter
replacement XML now writes both series fill and line color so sidecar reopens retain
the authored series color instead of substituting the theme palette.

Verification so far:

- Workbook-session suite: **128 tests / 30 files pass**.
- Targeted actual browser tests: original and session dialog add/reorder/rename,
  rejection/retry, pending Escape/cancel/edit guards, Save waiting, transpose and
  canceled deletion all pass. Existing format-pane context regression also passes.
- Sidecar save suite: **2 tests pass**, including all eight supported original
  chart families, retained and cross-sheet added series, colors, caches and two
  independent saves/reopens. Scatter categories follow existing numeric-X rules.
- Plugin build/schema/boundary checks and gateway/session/editor/desktop TypeScript
  checks pass. Full browser and native VS Code regressions are tracked below.

This is a scoped checkpoint. Contextual chart ribbons, visual arrangement/size/alt
text, broad plot/parameter edges, general original formula/cache freshness, actual
model/chat discovery, all-locale/appearance interaction and installed Windows/macOS
acceptance remain open. Earlier intermittent focus/cancellation observations remain
on the audit. No commit, push or VSIX packaging was performed.

Select Data checkpoint final regression results:

- Complete browser regression: **40 passing / 3.0 minutes**.
- Editor unit tests: **18 / 7 files pass**; browser presentation/search bridge
  boundaries: **11 / 2 files pass**. Changed-source lint, gateway TypeScript and
  extension/DOCX TypeScript pass; `git diff --check` passes.
- Native VS Code 1.109.5: the combined run passed shortcuts and all-nine chart
  insertion but failed the existing format-pane title-focus assertion before the
  added Select Data steps. It also emitted bare `Canceled` messages. The isolated
  floating-chart scenario subsequently passed in **22 seconds**, including actual
  Select Data add/rename/reorder, exact native Undo/Redo roots, Save, XML series
  count/names, distinct-session reopen with values/color, and final cascading
  deletion. This verifies the new native path but does **not** resolve the earlier
  intermittent focus/cancellation audit; do not report the combined run as clean.

- Final plugin rebuild passes (**29.49 seconds** for Vite), including generated
  schema and boundary checks. Shared desktop main/preload/renderer rebuild passes
  (**35.46 seconds** for renderer).
- Both Select Data browser scenarios pass again after adding concurrent-root
  rejection checks (**2 passing / 20.1 seconds**): a dialog opened before another
  transaction cannot publish its stale series selection and retains its draft.
- Final changed-source formatting and `git diff --check` pass.

## Chart Design contextual ribbon checkpoint (2026-10-26)

Selecting a floating source or session chart now exposes the actual desktop Chart
Design ribbon. It uses the shared SheetsRibbon controls, labels and state checks,
with the dynamic tab included in roving keyboard navigation. Deselection/deletion
removes the contextual tab and returns its active panel to Home. The chart echo
comes from the effective current session chart, including type, legend, grouping,
labels and axis titles.

The desktop's existing command mapping is extracted into `chartRibbonEdit` and
called by both applications. Quick Layout remains one transaction for legend and
labels; palettes target pie points or non-pie series using the same limits/colors
as desktop. Element menus, grouping, supported type conversions, Select Data,
Switch Row/Column, format-pane and deletion commands use the canonical host path.
Existing single-plot conversion limits and label-family restrictions remain.

The actual desktop ChartTextDialog is extracted from ExcelShell into a browser-only
component used by both applications. It awaits commits, disables input/cancel while
pending, guards repeated Enter/clicks and retains drafts/errors on rejection. The
plugin captures the opening chart/root/revision, rejects concurrent stale writes,
and tracks text edits during native Save flushing. Axis titles and grouping are
included in the plugin's supported canonical edit properties.

Initial verification:

- Editor and desktop TypeScript pass; extension/DOCX TypeScript pass.
- Original/session Chart Design browser scenarios pass: atomic layouts, state
  checkmarks, grouping, palette, title failure/retry/held Save, stale axis-title
  rejection, type conversion, dialog/pane entrances, deletion and tab fallback.
- Three further browser scenarios pass for pie/scatter/combo controls and palette
  semantics. Shared presentation/search boundaries: **12 tests / 2 files pass**,
  including the extracted title/axis-title dialog. Canonical ribbon candidate
  regression: **30 tests pass**.
- Plugin build/schema/boundary checks pass (**28.30 seconds** for Vite).
- Changed-source lint has no errors; ExcelShell retains the pre-existing
  `onListFunctions` useMemo dependency warning (also present at HEAD).

Full browser/native regressions and final shared desktop build are recorded below.
The first new native test reached the axis-title saved-file reopen successfully
but expected an explicit `dataLabels: 'none'`; the parser correctly represents
omitted labels as undefined. Its assertion now checks the semantic default.
Earlier focus/cancellation observations remain on the audit. This checkpoint does
not complete contextual Format/visual commands or the full menu-replication goal.
No commit, push or VSIX packaging was performed.

Chart Design checkpoint regression follow-up:

- Complete browser suite: **45 passing / 3.0 minutes** before the background
  backup refinement below.
- The final combined native run again passed shortcuts and nine-chart insertion
  but hit the earlier pane-title focus assertion before the new ribbon steps.
  Inspection identified a concrete interruption path: hot-exit backup requested
  the same `view.flush` that explicitly blurs chart pane drafts for Save.
- Background workbook backup now sends `commitChartDraft: false` to a validated
  flush request. It still waits for in-flight work and flushes cell edits, but
  captures the committed chart root without forcing an unsubmitted pane draft to
  blur. Explicit Save/history/tool flush defaults are unchanged.
- A browser regression now focuses a rejected pane draft, performs the background
  flush and checks unchanged focus, text and revision; the existing subsequent
  explicit-flush/held-commit checks still pass. The two Chart Design scenarios
  also pass after this change: **3 targeted browser tests / 26.0 seconds**.
- The native scenario now invokes the actual workbook backup function while a
  pane title is focused, checks unchanged chart root/focus/draft, then continues
  typing and explicit Save. Native final results are recorded after rebuild.

- Final rebuilt plugin/schema/boundary checks pass (**28.67 seconds** for Vite).
  Shared desktop main/preload/renderer rebuild passes (**38.48 seconds** for
  renderer). Editor unit suite passes **18 tests / 7 files**. Final changed-source
  formatting and `git diff --check` pass.

Final native verification after the background-backup fix:

- VS Code 1.109.5 combined suite: **3 passing / 45 seconds**, exit zero. This
  rebuild includes the actual Chart Design band and shared title/axis dialog,
  background backup flush option and current integration assertions.
- The real workbook backup call preserves the focused pane title draft and exact
  current root. Subsequent keyboard typing and native Save commit the complete
  title. The test uses an independent temporary backup slot and restores the
  document's original native backup identifier afterwards; the preceding test
  attempt correctly rejected reusing a host slot in a different storage directory.
- Quick Layout is one revision with exact native Undo/Redo roots. The shared axis
  title prompt writes the saved XML; a distinct-session reopen retains the axis
  title, bottom legend and semantic absence of data labels. Source series and
  drawing/history/save/removal checks continue to pass.
- No bare `Canceled` message appears in this final run. Earlier observations remain
  historical audit evidence; this result verifies the concrete backup focus fix,
  not every possible editor/window lifecycle condition.
- Final production-source lint checks pass except the already documented unchanged
  ExcelShell dependency warning. No commit, push or VSIX packaging was performed.

## Shared visual properties and layer menu checkpoint (2026-10-26)

The existing desktop Alt Text and Size and Properties dialogs are now mounted for
source and session charts. The plugin enables the actual chart context-menu
entries and both Bring/Send layer submenus. The shared VisualDialog now awaits
host commits, disables input/close while pending, guards repeated submissions and
preserves rejected drafts/errors. The plugin captures visual/root/revision at open
and tracks the complete operation for native Save flushing. The desktop chart
size dialog exposes placement modes; chart frame width/height remain controlled
by the existing resize handles, matching the shared desktop dialog's capabilities.

Canonical `edit_visual` accepts an editor visual ID with optional alternate text
(up to 2,000 characters), placement mode or absolute layer index. The host resolves
trusted source/session ownership and rejects unknown, protected, removed,
ambiguous, unsupported-placement and out-of-range targets. Standalone file CLI
execution explicitly refuses editor visual identities. Desktop AI dispatch uses
the existing undo-aware metadata/reordering paths; plugin agent transactions use
the canonical session engine. Tool schemas are regenerated.

Source drawing metadata patches now compose with previous geometry/properties;
subsequent moves no longer overwrite alternate text or placement. Deletion keeps
only the removal tombstone. Session chart metadata is written inside chart-add
properties, preserving it when serialized as new drawing frames. Source layer
changes expand into a consistent order for the surviving drawing part, using the
desktop's existing edited-slot permutation rule extracted as `projectVisualOrder`.
Session chart layer changes reorder the model array. The runtime applies metadata
and source order to current chart descriptors before rendering.

Verification so far:

- Full workbook-session suite: **131 tests / 30 files pass**. New coverage includes
  source/session metadata across geometry/layer changes, exact inverse roots,
  unchanged source metadata, deletion and atomic invalid/protected/removed cases.
- CLI XLSX regression: **12 tests pass**. Shared browser presentation/search bridge
  boundaries: **13 tests / 2 files pass**, including VisualDialog.
- Two property-dialog browser scenarios pass for source/session charts, including
  rejected draft retry, held submission, disabled Escape/cancel/edit, Save waiting,
  placement echo and concurrent stale submission rejection.
- Two source/session layer scenarios pass: actual submenu actions change the SVG
  presentation order for front/one-step-back commands with one revision each.
- Sidecar save suite: **2 tests pass**, covering source charts in eight families
  and session charts in all nine families through two independent saves/reopens.
  Added checks prove alternate text, placement and drawing order survive saving.
- Gateway/session/editor/desktop and extension/DOCX TypeScript checks pass;
  changed-source lint and `git diff --check` pass. Plugin/schema/boundary build
  passes (**35.29 seconds** for Vite).

Full browser/native regressions and final shared desktop build follow below.
Broad shape/image presentation, remaining arrange/rotate/size/link/clipboard
controls, general reference/formula freshness, other shared ribbon tabs, DOCX,
all-locale/appearance interaction and installed-platform acceptance remain open.
No commit, push or VSIX packaging was performed.

Final visual-properties checkpoint verification:

- Complete browser regression: **49 passing / 3.5 minutes**. Two further layer
  scenarios pass after adding actual-overlap checks (**2 passing / 16.8 seconds**):
  three charts share the same frame, `elementsFromPoint` identifies the front
  chart, a real mouse context-menu Send to Back changes the hit/occlusion target,
  and SVG order matches the committed source/session model. Geometry changes made
  to create the overlap preserve the existing layer/property journal.
- Real VS Code 1.109.5 combined regression: **3 passing / about 2 minutes**, exit
  zero. Actual source Alt Text/placement dialogs produce exact native Undo/Redo
  roots, escaped drawing `descr` and `editAs` attributes, unchanged chart-part XML
  and a distinct-session reopen retaining metadata and the same anchor. A direct
  agent-origin `edit_visual` transaction also undoes to its exact pre-edit root.
  Native shortcuts, nine-family insertion, previous chart/ribbon/history/backup
  and cascading deletion checks continue to pass. Bare `Canceled` messages remain
  observable during close/reopen; the general lifecycle audit remains open.
- Shared desktop main/preload/renderer build passes (**41.09 seconds** for renderer).
  Existing desktop visual arrangement/add-edit/schema/delete regression passes
  **41 tests / 4 files**. Final gateway and extension/DOCX TypeScript checks pass.
- Changed-source lint and formatting pass; `git diff --check` passes.

## Source shape/picture contextual formatting checkpoint (2026-10-26)

The plugin now installs the actual desktop visual renderer for source shapes and
pictures as well as charts. It preserves the chart-edit overlay when composing
that combined visual list. Selecting an editable source shape or picture exposes
the corresponding actual SheetsRibbon Shape Format/Picture Format band, with
shared translated controls, current paint state and contextual keyboard tab order.
Unsupported clipboard/hyperlink/embed context actions remain disabled.

The shared frame calculation is extracted as `frameChangeFromBox`; desktop and
plugin size/rotation commands use the same center, true-frame, rotated-AABB and
anchor rules. Shape fill/outline, rotate/flip and shared numeric size/properties
controls submit validated drawing properties. The host rejects paint/transforms
on unsupported families and a true-frame change without a preceding anchor in
the same transaction. Source drawing patches compose with prior properties.
Multi-selection align/distribute use the desktop installed-frame geometry helpers,
validate complete same-sheet selection coverage and commit one canonical batch.
Dialog operations remain frozen to the opening root/revision and participate in
Save flushing, pending-control guards and rejected-draft retention.

Floating media now comes through an identity-scoped `source.media` bridge. The
host validates a unique current source image/shape and a surviving logical sheet,
rejects removed/foreign/unsupported identities, verifies the fixed source session
and strictly validates image MIME/base64 output (bounded below the bridge transfer
budget). Reads overtaken by another root fail. No file path or source-session
argument is accepted from the browser. The shared desktop media hook/decoder
receives the bytes through its existing reader callback.

Authoritative verification:

- Complete browser regression: **51 passing / 3.5 minutes**. New scenarios verify
  source-picture decode, actual contextual tabs, split color controls, paint,
  rotation/flip, exact authored size units, rejected size drafts, held submission
  and Save waiting, plus physical multi-selection alignment and equal horizontal/
  vertical gaps. Existing chart, source/session layer, cell, find and menu tests
  continue to pass.
- Full session suite: **136 tests / 31 files pass**. Added tests cover scoped media
  ownership/response validation and stale reads, composing paint/frame/anchor/flip
  changes and atomic unsupported-target/frame rejection.
- Real VS Code 1.109.5 source-visual scenario: **1 passing / 28 seconds**, exit zero.
  It decodes an actual workbook picture, edits source shape paint and numeric
  frame/rotation, flips the picture, aligns multiple shapes, checks exact native
  Undo/Redo roots and single revisions, saves/reopens a distinct document session,
  and verifies authored frame dimensions/rotation/paint/flip plus unchanged media
  bytes and custom XML sentinel.
- Existing native shortcut/nine-chart/floating-chart regression: **3 passing /
  51 seconds**, exit zero, on the same rebuilt plugin. Bare `Canceled` messages
  remain observable on close/reopen; the general lifecycle audit remains open.
- Existing desktop visual regression: **41 tests / 4 files pass**; shared browser
  presentation/search boundaries: **13 tests / 2 files pass**; editor units:
  **18 tests / 7 files pass**.
- Plugin/schema/boundary build passes (**38.14 seconds** for Vite). Final gateway,
  workbook-session, editor, desktop and extension/DOCX TypeScript and changed-source
  lint checks pass. The shared desktop rebuild is recorded below.

This checkpoint adds source-object editing. New-shape/new-picture insertion,
clipboard/link/embed flows, other shared ribbon tabs, full DOCX menus, broader
family/parameter cases, all-locale/appearance interaction and installed Windows/
macOS acceptance remain unfinished. The full menu-replication goal stays open.
No commit, push or VSIX packaging was performed.

- Final shared desktop main/preload/renderer rebuild passes (**1 minute 19 seconds**
  for renderer). Final source formatting and `git diff --check` pass. All recorded
  test processes were inspected to completion; none is being left as an assumed
  background verification job.

## New session shapes and text boxes checkpoint

The actual shared Shapes gallery now enters the desktop draw gesture in the
plugin grid. Click inserts at the desktop 96-pixel default; dragging uses the
actual grid dimensions, Shift constrains a square, and Escape/window blur cancels.
Drawing over an existing shape does not move it. A frozen gesture overtaken by
another transaction rejects without adding an object. Text Box retains the
shared desktop cell-based size, white fill and initial text.

Session shapes and charts share one ordered canonical visual list. Shape text,
paint, transforms, frame size, placement and layer operations use canonical
transactions, including manual/AI history. Shared text editing retains rejected
drafts and disables pending edits; explicit Save awaits focused shape text.
Automatic backup retains the pre-existing draft-preserving behavior.

Evidence recorded for this checkpoint:

- Workbook-session suite: **138 tests / 31 files pass**, including mixed visual
  order, shape/text-box defaults, stable IDs, validation/protection and inverses.
- Package writer: **3 tests pass**, including two independent saves/reopens of
  **54 mixed objects**, authored text/no-fill/outline/transform/frame properties
  and exact order. Additions are written in bounded batches of 50; a regression
  caught and fixed missing required `edits: []` in subsequent batches.
- Real VS Code 1.109.5 combined suite: **4 scenarios pass / 2 minutes**, exit zero.
  The source-visual scenario additionally draws a new shape, checks exact native
  Undo/Redo, saves while its text editor still owns focus and reopens a distinct
  session with the exact anchor and authored text. Existing shortcut and chart
  scenarios also pass. Bare `Canceled` close/reopen messages remain observable;
  the general lifecycle audit stays open.
- Editor unit suite: **18 tests / 7 files pass**; presentation/search boundaries:
  **14 tests / 2 files pass**; desktop visual/geometry/transaction checks:
  **69 tests across 6 files pass** after updating two stale transaction assertions
  to match the explicit fill-clear override and shared chart label defaults.
- Gateway, session, editor, desktop and extension/DOCX TypeScript checks pass.
  Changed production-source lint and formatting checks pass. Plugin build,
  generated tool-schema and browser boundaries pass (Vite: **1 minute 23 seconds**).

The full browser suite passes: **53 tests / 5.2 minutes**, exit zero. It includes
the new shape/text-box scenarios, failed draft retry, pending Save, default click
size, Shift constraints, overlap drawing, cancellation and stale-gesture rejection.
Final shared desktop main/preload/renderer build passes, exit zero (renderer:
**36.77 seconds**). Final formatting and `git diff --check` pass. All verification
processes were inspected to completion.
New-picture insertion, graphical links/clipboard/embed, remaining shared ribbon
tabs, full DOCX menus and broad preset/locale/appearance/platform acceptance
remain open. No commit, push or VSIX packaging was performed.

## New floating pictures checkpoint

Insert Pictures now uses the actual shared desktop file picker/read/decode path:
PNG/JPEG/GIF, at most 20 MiB per file, natural-aspect cell sizing and the same
480-by-320 fallback when decoding fails. The plugin freezes the selected sheet,
anchor and root before selection. Cancellation and invalid files do not change
the document, and a selection overtaken by another edit rejects atomically.

The bounded `add_picture` operation validates media type, base64, byte length and
signature before appending a stable allocated picture identity. Pictures share
the ordered session visual union with charts/shapes and the actual desktop
Picture Format/menu/renderer. Existing canonical geometry, transforms, metadata,
layers and deletion are reused. The desktop operation executor also handles this
inline operation through its normal visual journal. Existing path/URL `add_image`
loading remains separate; clipboard/drop and full AI acquisition flows remain
acceptance work.

Picture-bearing roots execute the same candidate interpreter directly in the
extension host. This avoids sending repeated next/forward/inverse image bytes
through candidate transfers. The accepted root is projected normally; the existing
chunked undo and recovery stores retain its bytes without an external image cache.
General transaction/history limits still apply. The package writer preserves
mixed authored order and uses the existing bounded addition batches.

Evidence:

- Full actual browser regression: **54 passing / 4.2 minutes**, exit zero. The
  new picker scenario verifies actual accepted file types, invalid-file/cancel
  behavior, pending commit and Save, decoded pixels, desktop size defaults,
  Picture Format flip, and an overtaken selection. After final error-cause
  preservation, its targeted rerun also passes (**1 / 18.3 seconds**).
- Workbook-session: **140 tests / 31 files pass** with two workers. The first
  highly concurrent run timed out in the existing history-storage test; that
  exact test passes independently, and the complete final suite passes without
  changing its timeout. New coverage includes picture identities/transforms/
  order/inverses, invalid media/paint/protection and chunked backup restoration.
- Real VS Code 1.109.5 source-visual scenario: **1 passing / 49 seconds**, exit
  zero. It additionally opens the actual picture chooser, decodes the inserted
  image, verifies exact native Undo/Redo roots, saves and checks original picture
  bytes in XLSX media, then opens a distinct session and verifies both pictures
  and the inserted anchor. The earlier source shape/image and new-shape checks
  remain in the same scenario.
- Package writer: **3 tests pass**, including **55 mixed objects** in authored
  order through two independent saves/reopens; exact picture bytes, transforms,
  alternative text and frame dimensions survive.
- Host picture-candidate regression: **1 test passes** with a 20 MiB inline PNG
  payload, subsequent transform and deletion, asserting no candidate bridge
  transfer is invoked and inverse roots remain exact.
- Desktop picture/visual/transaction regressions: **55 tests / 4 files pass**.
  Shared presentation/search browser boundaries: **15 tests / 2 files pass**.

Final gateway, workbook-session, editor, desktop and extension/DOCX TypeScript
checks pass. Changed production-source lint, source formatting and `git diff --check`
pass. Final plugin/schema/boundary build passes (Vite: **52.17 seconds**); shared
desktop main/preload/renderer build passes (renderer: **55.46 seconds**).
The final plugin also passes the **3 existing native shortcut/nine-chart/floating-
chart scenarios**, exit zero. Startup responsiveness profiling and the pre-existing
bare `Canceled` close/reopen messages remain observable; the general lifecycle
audit remains open. All verification jobs were inspected to terminal results.
New clipboard/drop/embed and graphical-link flows, remaining shared Insert/base
ribbon commands, full DOCX menus, broad image/preset/locale/appearance and installed
platform acceptance remain unfinished. The full replication goal stays open.
No commit, push or VSIX packaging was performed.

## Shared clipboard transfers and external cell paste checkpoint

The plugin uses the actual desktop picture paste/drop installer, accepted image
rules and scene/scroll/scale hit test. The embedded focus host is scoped to its
own grid instead of the desktop container ID. Picture transfer processing freezes
the target/root before asynchronous reads, joins Save flushing and disposes its
listeners with the view. Cell inputs, shape text and other native fields retain
their editing paste behavior.

Hidden native toolbars left the default Univer clipboard hooks uninitialized.
`SheetClipboardController` is now initialized after workbook creation, when its
registration is available. Table/text clipboard payloads, including table preview
bitmaps, go to the desktop HTML/text parser without file items. Supported native
value/style/merge mutation plans are captured, stopped before native projection/
undo application and submitted as one canonical transaction. Hooks are restored
in a finally block; unknown mutation families and mixed rich text reject before
publishing a batch. Boolean cell types and authored computed CSS RGB colors are
preserved. The shared color fix applies to desktop serialization too.

Importing the neutral-style converter initially pulled desktop AI settings through
its constants import. The constants now come directly from the pure gateway
schema module. The browser dependency boundary remains unchanged, and a new
format-converter boundary regression prevents that indirect import from returning.

Evidence for the transfer/parser implementation:

- Full browser attempt: **56 of 57 pass / 4.9 minutes**. The picture-picker test
  timed out waiting for its first file chooser. That exact test subsequently
  passes **four independent repetitions / 54 seconds** without increasing its
  timeout. The full-run intermittency remains an open acceptance item; this is
  not recorded as a clean full-suite pass. Failed picker diagnostics now include
  editor, composition, projection and pending-work state.
- Actual Chromium clipboard Ctrl+V and file drop tests pass: selected-cell versus
  mouse-cell anchors, decoded pictures, pending commit/Save, mixed table previews,
  input focus, stale roots and rejected external cell paste retry are exercised.
  Styled HTML tables preserve font/fill colors, numeric values, italic/bold and
  row spans as one revision; rejected or pending writes leave visible cells alone.
- Real VS Code 1.109.5: **5 scenarios pass / 2 minutes**, exit zero. The new native
  transfer case covers pasted/dropped pictures, mixed preview/table input, exact
  Undo/Redo roots, authored table formatting, saved image bytes, unchanged custom
  XML, distinct-session reopen, exact anchors and successful decoding of each
  reopened image. Existing shortcuts/source objects/chart insertion/floating chart
  scenarios pass too. Bare close/reopen `Canceled` messages remain under the
  existing lifecycle audit.
- Shared desktop journal/picture/focus tests: **106 tests / 3 files pass**.
  Editor units: **20 tests / 8 files pass**, including reversible typed/style
  paste and atomic unsupported-content rejection. Shared presentation/search
  boundaries: **16 tests / 2 files pass**.
- Editor, desktop and extension/DOCX TypeScript checks pass. Focused production
  lint has no errors. The existing large App effect still reports three hook
  warnings. Source formatting and `git diff --check` pass.
- Transfer-only final plugin/schema/boundary build passes (Vite **35.02 seconds**),
  as does desktop main/preload/renderer (renderer **39.18 seconds**). Paste-button
  additions below have their own final verification.

The actual shared Home Paste button now reads the first ClipboardItem, applies
the same image-versus-table/text rule and invokes these same authoritative paths.
Its clipboard read participates in Save flushing and retains the target/root.
An unavailable read gives a localized native-shortcut hint; native Ctrl+V remains
usable. Chromium may re-encode PNG data when placing it on the OS clipboard: tests
compare insertion with bytes returned by the clipboard, rather than assuming its
encoding matches the bytes originally written there.

- New actual Paste-button browser scenarios: **2 pass / 17.4 seconds**, covering
  image and styled table commits, preview exclusion, delayed read/Save, stale root
  rejection and clipboard-read denial followed by successful native Ctrl+V.

The final Paste-button extension build/schema/boundary checks pass (Vite:
**1 minute 12 seconds**), as does the shared desktop main/preload/renderer rebuild
(renderer: **40.26 seconds**). Final editor, desktop and extension/DOCX TypeScript
checks pass; focused changed-source lint has no errors.

The native transfer scenario passes again on that build (**1 / 15 seconds**), now
also clicking the actual Home Paste button, checking its exact Undo/Redo roots,
saving three images plus a styled table, reopening a distinct session and decoding
all three pictures with their exact anchors. Its menu read uses injected known
ClipboardItem data to test the native menu transaction independently of platform
clipboard permission policy; actual Clipboard API read and denied-read/native
Ctrl+V recovery are covered by the Chromium browser scenarios above. This is not
installed Windows/macOS clipboard permission acceptance.

Source formatting and `git diff --check` pass. Verification jobs have terminal
results. The earlier full-run picker failure remains recorded and open.
Cut/internal-source freshness, clipboard-cache fallback, Paste Special, mixed rich
text, additional mutation families/automatic row sizing, full base/Insert/DOCX
ribbons, all-locale interaction and installed-platform acceptance remain open.
The full menu-replication goal stays active. No commit, push or VSIX packaging.

## Typed Copy snapshots and cross-sheet paste

The actual shared Home Copy button and native grid Ctrl/Cmd+C now read complete
selected source cells before invoking the desktop clipboard serializer. A detached
snapshot keyed by its native copy ID survives source edits and view reinstalls.
Same-editor Paste submits one bounded `paste_cells` transaction, retaining literal
formula-looking strings, errors, dates, replacement formatting, independent border
edges and merge geometry; copied formulas shift and drop stale caches.

Evidence on the current implementation:

- Targeted actual Chromium Copy/Paste scenarios: **2 pass / 16.3 seconds**. The
  menu copies cells outside the initial viewport and retains their original data
  after a source edit. Real grid Ctrl+C/Ctrl+V tiles a merged block, preserving
  anchors and blanking covered cells in one revision.
- Editor unit suite: **24 tests / 9 files pass**. Pure snapshot checks cover
  detachment, mixed references across sheets, merge translation/partial-source
  rejection, missing source cells, native copy identity, target merges and bounds.
  Native clipboard border tests cover distinct edges, explicit clears and CSS RGB.
- Session suite: **142 tests / 31 files pass / 22.34 seconds**. The first run had
  141 passes and a 20-second history-storage timeout while a Vite build and
  TypeScript jobs ran concurrently. That test passed independently in 8.93
  seconds and the full unchanged suite passed on the subsequent run; its timeout
  was not increased. New candidate checks cover typed values/styles/formula-cache
  invalidation, inverse roots, contained/partial merges, protection and date systems.
- XLSX gateway suite: **303 tests / 43 files pass**. Shared desktop journal,
  transactions and picture-paste units: **115 tests / 3 files pass**. CLI DSL
  regressions: **12 pass**. A CLI chart-visual guard now narrows optional chart
  state before conversion, matching the mixed visual union.
- Extension ribbon/isolation boundary suites: **18 tests / 2 files pass**.
- Native VS Code 1.109.5 Copy scenario: **1 pass / 9.67 seconds**. It clicks the
  actual Copy and Paste controls, switches sheets, changes source after Copy,
  checks exact native Undo/Redo roots, verifies literal-error XML, shifted formula
  XML and independent border styles, retains a custom package sentinel, then
  opens a distinct session and reads the saved literals/errors/dates/formulas.
  Its iframe uses an injected ClipboardItem store with complete read/write/readText
  interfaces; this isolates menu/history/persistence from OS clipboard policy.
  The first fixture lacked readText and invoked the native legacy-copy fallback;
  correcting the fixture made the scenario pass. Actual OS Clipboard API behavior
  is covered by the Chromium cases, not installed Windows/macOS acceptance.
- Plugin build/schema generation/boundaries pass (Vite **1 minute 5 seconds**).
  Desktop main/preload/renderer rebuild passes (renderer **39.07 seconds**).
  Focused changed-source lint and source formatting pass.

`paste_cells` is an ordinary bounded operation (at most 2000 cells), not a new
unlimited document capacity. Desktop execution explicitly rejects literal error
cells before applying a batch because the legacy desktop journal cannot preserve
their kind; VS Code canonical persistence covers them. Existing CLI supported-op
lists continue to reject the new operation rather than silently degrading it.

Cut, Paste Special, copied links/notes/validation/conditional formatting/rich text,
filtered/hidden-row parity, cross-editor typed snapshots, denied-copy reporting,
platform clipboard fallback and installed-platform acceptance remain open. Full
base/Insert/DOCX ribbon replication and locale/appearance gates also remain open.
The goal stays active; this checkpoint does not request commit/push/VSIX packaging.

Final verification for this Copy checkpoint:

- Complete Home/manual browser suite: **61 pass / 4.6 minutes**, including the
  original picture picker, native-key merged Copy, menu snapshot Copy, Paste read
  denial/recovery, Save flushing, graphical transfers and existing chart/shape
  controls. This clean run does not establish a root-cause fix for the previously
  recorded intermittent picker timeout; its audit task remains open.
- Final editor, session, gateway, CLI, desktop and extension/DOCX TypeScript checks
  pass. Source formatting and `git diff --check` pass. All verification jobs have
  terminal results.

## Atomic Cut checkpoint

The shared Home Cut button and native grid Ctrl/Cmd+X now preserve the current root
until Paste and route matching same-editor clipboard identities to canonical
`relocate_range`. This path reads full dependency material and moves captured typed
cells/formats, merge geometry, plain notes and hyperlinks with one history entry.
Contained A1 references, including `$` references, follow the moved cells. Moved
formulas retain outside dependencies on their original sheet. This is distinct
from legacy `move_range`, whose old copy/clear expansion is not used here.

Current evidence:

- New actual Chromium Cut cases: **3 pass / 24.1 seconds**. They exercise menu Cut,
  an overlapping destination, a dependent outside the viewport, one revision,
  consumed-cut rejection, real grid Ctrl+X/Ctrl+V with stale-source rejection,
  Escape cancellation without a root change and subsequent successful Copy.
- Canonical relocation/material tests: **8 pass** across two targeted files. These
  cover cross-sheet typed literals/errors/dates and independent borders, formula
  caches, outside dependencies, overlapping source capture, exact inverse roots,
  merges/notes/links, incomplete material, protected source/destination/dependents,
  partial merges, write budget and reading distant original formula cells.
- Formula move tests: **4 pass**, including mixed absolute/relative references,
  partial ranges, quoted sheet names, original-sheet outside dependencies, whole
  rows/columns, strings, table column names, external workbooks and short 3D sheet
  names. Partial-range behavior was checked against the installed desktop
  sheets-formula controller's MOVE_RANGE path, which only rewrites contained refs.
- Full workbook-session suite: **148 tests / 32 files pass / 32.39 seconds**.
  Gateway suite: **307 tests / 44 files pass / 13.81 seconds**. Editor units:
  **24 tests / 9 files pass**. Extension isolation/ribbon boundaries: **18 pass**.
- Native VS Code 1.109.5 Cut and existing Copy scenarios: **2 pass / 19 seconds**.
  Cut clicks the actual shared control, moves across sheets, updates a distant
  dependent, verifies exact Undo/Redo roots, saves and checks moved merges,
  hyperlink coordinates, note text, literal-error XML, escaped cross-sheet formula
  XML, independent borders and an untouched custom package entry. A distinct
  session reopens the saved literal/error/date/formula cells. ClipboardItem
  read/write/readText storage is injected inside the iframe, so these results
  verify menu/history/package behavior rather than installed OS clipboard policy.

Remaining scope is explicit: affected defined names and native chart references,
CF/validation/filter/table/pivot/sparkline metadata, rich notes, broad chart behavior,
cross-editor cut and clipboard denial/fallback still need implementation or wider
acceptance; interim atomic rejection is not feature completion. The 2000-write
limit includes both source clears and destination/dependent writes; full source
material retains the existing per-sheet range budget. Other ribbons, DOCX and
locale/appearance/installed Windows/macOS acceptance remain open. No commit, push
or VSIX packaging is requested by this checkpoint; the full goal remains active.

Final Cut validation:

- Full Home/manual browser regression: **64 pass / 4.7 minutes**. The previously
  intermittent picture picker passes again; its root-cause audit remains open.
- After the final error-presentation change, all four Cut browser cases pass
  (**4 / 29.3 seconds**), including an unsupported-metadata response that displays
  the existing localized no-change message, preserves the complete root and allows
  retry with the pending cut. This final change only maps that rejection prefix;
  the transaction and reference logic used by the 64-scenario run is unchanged.
  No new locale keys are needed because the message already exists in all shards.
- Native Cut passes again (**1 / 6.42 seconds**) with additional independent-session
  reads verifying every source cell is blank and the distant source-sheet formula
  has its persisted destination-sheet reference. This supplements the two-scenario
  native Copy/Cut run above.
- Shared desktop journal/transaction/picture units: **115 pass / 3 files**. CLI DSL
  regressions: **12 pass**. Final desktop main/preload/renderer build passes
  (renderer **38.35 seconds**); final plugin schema/build/boundaries pass (Vite
  **36.15 seconds**). Editor/session/gateway/desktop/CLI TypeScript checks pass.

The earlier Copy-only discussion of Cut as entirely unimplemented is superseded
by this scoped implementation and its explicit remaining feature-transform gaps.
This checkpoint does not complete the full desktop-menu replication goal.

The final extension/DOCX TypeScript run also passes after error-presentation
polishing. Focused lint, final source/document formatting and `git diff --check`
pass. Every verification job for this checkpoint has a terminal result.

## Shared Paste Special checkpoint

The plugin uses the actual desktop Paste Special dialog and shared availability
model, plus the extracted desktop transform. Five quick entries and grid
Ctrl/Cmd+Alt+V are connected to authoritative canonical transactions. Optional
`paste_cells` fields preserve target values/formulas for format-only work, preserve
target styles for values/formulas, replace number formats independently, retain
borders and skip individual cells without scalar-operation expansion.

Verified behavior so far:

- Original desktop pure transform tests: **16 pass** after extraction with API
  re-exports preserved. New transform-to-canonical unit cases cover arithmetic,
  transpose/skip, links, literals/errors/divide-by-zero, column widths and pending
  formula combinations that do not need computed values.
- Canonical field tests: **3 pass**, covering retained target fields, number-format
  resets, independent borders, skipped cells, inverse roots and a styled 101-cell
  block in one operation. Full session suite: **151 tests / 33 files pass / 32.63
  seconds**. Gateway suite: **307 tests / 44 files pass / 10.98 seconds**.
- Final targeted browser cases: **8 pass / 50.4 seconds**, including all five quick
  entries, exact dialog defaults/disabled choices, arithmetic + transpose + skip
  in one revision, links/shortcut, current source column widths, failed retry,
  delayed Save, stale dialogs, external HTML/text and desktop source freshness.
  Plain HTML without a boolean type marker retains native parser text semantics;
  shared text transposition converts TRUE to a boolean as the desktop does.
- Source freshness was checked against the installed native clipboard service and
  desktop own-layer implementation. Plain cached values differ from current
  formatting, current column widths and advanced current source values; the
  browser cases explicitly exercise those distinctions. An earlier full run was
  deliberately interrupted to correct those semantics and is not a passing gate.
- Initial native VS Code Paste Special case: **1 pass / 24 seconds**. It exercises
  the actual cross-sheet dialog, exact native Undo/Redo roots, arithmetic/transpose/
  skip, values-plus-number-formats retaining target paint/borders, column widths,
  custom package sentinel preservation and distinct-session saved cell reads.
  ClipboardItem read/write/readText is injected inside the iframe; OS clipboard
  policy remains covered separately by Chromium and is not installed Windows/macOS
  clipboard acceptance. The final native rerun additionally checks saved column
  width numbers by converting them through the workbook's Normal-font digit width.

The final complete browser/native/type/build results are recorded below when
terminal. Remaining formula computation freshness, metadata/CF/rich content,
filtered/merged target parity, source deletion/structural edits, cross-editor
ownership, installed platforms and other ribbon/DOCX groups remain explicit work.
Comments/Validation are disabled by the shared desktop dialog's own availability
model; they have not been independently implemented here. The full menu goal is
not complete. No commit, push or VSIX packaging is requested for this checkpoint.

Final Paste Special results:

- Complete final Home/manual browser suite: **73 pass / 5.6 minutes**. This is the
  run after desktop cached-value/current-format/live-source/column-width semantics
  were aligned; earlier intentionally interrupted runs are not counted as passes.
  The previously intermittent picture-picker test passes, with its cause still
  unproven and its audit task kept open.
- Final native Copy/Cut/Paste Special regression: **3 pass / about 1 minute** on
  VS Code 1.109.5. The strengthened Paste Special case verifies saved numeric column
  widths against the workbook Normal-font maximum digit width. Startup emitted a
  transient extension-host unresponsive warning and later recovered. Bare
  `Canceled` messages also appeared during lifecycle transitions. These remain
  performance/lifecycle audit evidence; successful scenarios do not resolve them.
- Editor units: **29 tests / 10 files pass**. Shared desktop transform/journal/
  transaction/picture units: **131 / 4 files pass**. CLI DSL: **12 pass**. Extension
  ribbon/isolation boundaries: **18 / 2 files pass**.
- Final plugin schema/build/boundaries pass (Vite **33.68 seconds**). Desktop
  main/preload/renderer rebuild passes (renderer **35.88 seconds**). Extension/DOCX,
  editor, session, gateway, desktop and CLI TypeScript checks pass. The final
  static fixes explicitly supply the fixture's existing default column count and
  accept schema-produced optional `undefined` in the field-composition type; they
  do not change application runtime behavior.
- Focused changed-source lint, formatting and `git diff --check` pass. Verification
  processes have terminal results.

The goal remains active. This does not claim complete clipboard metadata,
calculation freshness, all menu groups, DOCX menus, locales/appearances or installed
Windows/macOS acceptance. No commit, push or VSIX packaging was performed.

## Shared Page Layout checkpoint (2026-10-27)

The actual desktop Page Layout band replaces the plugin's simplified controls.
Desktop and plugin share the validated page-command decoder. Connected paths
include margins/orientation/paper, print area/titles, scale/fit, imported page
breaks, and separate screen/print grid and heading flags. Themes/Colors/Fonts
remain visibly unavailable and are not accepted as complete.

- Decoder units: **2 pass**. Source metadata, title-row cap, fit defaults,
  normalized area addresses, retained breaks and history-free empty break actions
  are covered. Full editor units: **31 pass / 11 files**.
- Shared desktop print-settings, print-paginate and edit-journal: **146 pass**.
  Extension boundary/manual-chart regressions: **18 pass**.
- Native VS Code Page Layout scenario: **1 pass / 19 seconds**. Checks source
  echoes, actual shared controls, ordinary clicks after choosing fit width,
  exact Undo/Redo roots, unchanged cell XML, saved setup/print names/breaks,
  unchanged package sentinel and a distinct-session reopen. An initial native
  run exposed a fit popup that swallowed the next Print click. Replacing the
  shared label wrappers with layout divs fixed that path; no forced clicks were
  used. A narrow browser-only reproducer had passed before the native fix and
  is not evidence that the native issue was absent.
- Page-only cache regression initially failed: an existing chart value of 12.5
  became zero when its source formula was pending. Print-only transactions now
  skip session/native chart-cache refresh and unrelated chart-cell hydration.
  Protection indexing is retained. Focused candidate/material tests: **46 pass**,
  including preserved session/native chart caches and missing unrelated chart
  references that do not block print metadata.
- Gateway full suite: **307 pass / 44 files / 21.40 seconds**. Desktop build:
  **pass**, including main/preload/renderer. Desktop/editor TypeScript checks
  passed after the shared-menu changes. Broader final checks are recorded below.

Full menu replication remains open: document themes, remaining base/Insert/DOCX
commands, broad locale/appearance/installed-platform and actual print/export
acceptance are not established by this checkpoint. Existing picker/lifecycle/
performance audit items remain open.

Final checkpoint evidence:

- Full Home/manual browser regression before the final container-style correction:
  **77 pass / 6.9 minutes**, including all four new Page Layout scenarios. The
  desktop build emitted Vite reload notices during this run; no cases failed.
- Full canonical session suite after the chart-cache correction and its added
  regressions: **154 pass / 33 files / 20.48 seconds**. A preceding loaded run
  had 153 passes and one 20-second history-storage timeout. Its unchanged
  standalone repeat passed in **3.62 seconds**; the timeout threshold was not
  raised. Retain this observation in the performance audit.
- Full VS Code TypeScript chain (workbook session, Sheets editor, DOCX editor,
  DOCX session, extension including DOCX config): **pass**. New test-fixture type
  errors found during verification were corrected before this final pass.
- Changed-code lint, theme-color check and all 21 manual locale structure checks:
  **pass**. Shared new action-conflict text exists in all 21 app locale shards.

After adding Page Layout to the shared desktop panel container (preventing the
legacy toolbar rules from overriding shared controls), the final targeted
browser run passed **5 cases / 36.7 seconds**, including an assertion that this
panel uses the shared container and no legacy toolbar class. The rebuilt native
VS Code Page Layout scenario passed again: **1 case / 20 seconds**, with exact
history, unchanged cell XML, saved print settings and independent-session reopen.
It logged a recovered extension-host unresponsive event and three bare `Canceled`
messages; retain the existing lifecycle/performance audit.

Final plugin build/schema/boundary checks: **pass / 32.98-second Vite build**.
Final container TypeScript/lint, changed-code Prettier and `git diff --check`:
**pass**. See `sheets-page-layout-ledger.md` for connected controls and remaining
Page Layout scope. No commit, push, VSIX package or Windows installed acceptance
was performed in this checkpoint.

## Shared Function Library checkpoint (2026-10-27)

The desktop and plugin now share the Function Library group, live catalog and
actual two-step Function Arguments flow. This is not full Formulas-tab parity:
Defined Names, Formula Auditing and Calculation settings/menu groups remain open
in `sheets-formulas-ledger.md`.

Canonical writes freeze the owning cell/sheet/root, survive pointer range picks,
retain failed parameters, disable pending actions and participate in Save
flushing. Shared AutoSum geometry emits a formula below each selected column in
one plugin transaction. Function preview first reads bounded static precedents
from the host, including original off-page and cross-sheet values; it declines
unverified formula/name/external/dynamic previews instead of substituting zeros.
Preview projection hydration does not change session cells or history.

- Initial catalog/arguments browser run: **5 pass / 44.4 seconds**, covering
  frozen targets, current inputs, failed retry, Save waiting, stale-root reject,
  category reopening and atomic multi-column AutoSum.
- Final targeted browser run after lifecycle fixes: **10 pass / 1.3 minutes**,
  adding original off-page reads, blank unverified formula previews, grid-only
  Shift+F3, real pointer range picks, startup/native-formatting regression and
  the existing seven-tab scenario.
- Editor unit suite: **34 pass / 12 files / 11.60 seconds**; shared function
  arguments/catalog/closure units: **50 pass / 4 files**. New pure tests cover
  qualified references, whole axes, budgets, unresolved names/external/dynamic
  references, per-column AutoSum and the grid boundary.
- Editor/desktop TypeScript and changed-code lint/Prettier: **pass**. Extension
  TypeScript including DOCX and boundary/manual-chart units: **18 pass**.
- Initial plugin build correctly rejected the function catalog's broad UI barrel
  because it pulled in the AI provider module. Switching the shared dialog to
  the existing pure `@genoffice/ui/chrome` entry fixed the boundary, with no
  weakening of the checker. Final plugin build/schema/boundary checks: **pass /
  35.02-second Vite build**. Final desktop main/preload/renderer build: **pass /
  38.19-second renderer build**.
- An initial broad browser attempt and native startup attempt exposed an actual
  lifecycle error: fx was subscribed before the formula editor service existed.
  Move the subscription to each ready workbook projection and dispose it with
  that projection. Explicitly pre-optimize newly used Univer UI/formula modules
  in both browser harnesses to prevent dependency-triggered page reloads. The
  original interrupted/failed attempt is not counted as a passing full run.

Broader final browser/native results follow below. Existing lifecycle/performance
and picker audits, full caret/cancel/computation behavior, all locales/appearances
and installed-platform acceptance remain open. No commit, push or VSIX packaging
was performed in this checkpoint.

Native diagnostics also exposed a shared fx caret bug: the editor may report
caret zero for `=`, so insertion previously produced `SUM(...)=`, rejected by the
canonical formula schema. The shared `finishedFormula` now clamps insertion past
the equals sign. A new pure regression and actual browser fx write pass; shared
function units now total **51 passes**. Native fx preview, insertion, exact
Undo/Redo, save and independent reopen subsequently executed, with a final test
assertion exposing a stale compiled whitespace expectation (`SUM(2, 3)` is the
shared desktop output). Recompile that corrected expectation before accepting
the native test as green. Failure cleanup logged a closed-lease revert error;
retain the lifecycle audit.

Function writes now flush an active grid draft before checking the frozen root,
track the entire write preparation for Save, and reject if that draft changed
the document. Opening also freezes the root before asynchronous target capture,
and repeated opens are single-instance. Preview effects ignore cancelled async
responses and clear failed results rather than retaining an old result.

Final native evidence: **1 pass / 21 seconds** after the final write-flush guard,
caret correction and test recompilation. Actual formula-bar fx, authoritative
cross-sheet 19.5 preview, canonical formula insertion, exact native Undo/Redo,
saved formula XML/package sentinel, a distinct-session reopen, parameter preload
and editing/saving the existing call all passed. This final run logged no closed
lease error; prior failure-cleanup/lifecycle observations remain in the audit.

Final submission-focused browser repeat: **4 pass / 34.5 seconds**, covering
retry/pending/Save, native fx caret-zero insertion, single-instance opening with
Save waiting for the source target read, and preserving an unfinished grid draft
while rejecting the overtaken formula target. Final plugin build/schema/boundary
checks passed after this guard (**41.02-second Vite build**). Latest shared desktop
build passed after the caret change (**53.00-second renderer build**). Editor
TypeScript and changed runtime/test lint pass after the final guard.

The completed full Function Library/Home/manual browser run passed **87 cases /
6.2 minutes**, including all ten new library/arguments/fx/range-pick/opening/draft
scenarios and the existing Home/Insert/Page Layout/chart/clipboard regressions.
Together with the final native pass above this closes the current Function
Library implementation checkpoint, not the overall desktop-menu goal.

## Canonical defined-name foundation (2026-10-27)

Existing add/delete name operations now support authoritative logical sheet scope
and participate in immutable history/save. This is backend and projection work;
it does not establish Name Manager UI, semantic rename/reference rewriting or
complete Formulas-tab parity. A delete/add definition replacement is atomic but
is not claimed to perform a semantic rename of dependent references.

- Focused canonical/source/save tests: **52 pass / 3 files**. Cover global/local
  same-name coexistence, atomic replacement/failure rollback, protection, invalid
  names/scope, coordinate-change guards and pending original formula caches.
  A real sidecar save/reopen verifies both scopes, hidden names and Print_Area
  retained verbatim, unchanged source worksheet XML and custom package sentinel.
- Full session suite: **158 pass / 34 files / 18.70 seconds** with two workers.
  A preceding higher-load run timed out the unchanged 20-second history-storage
  case (157 passed); no timeout threshold was raised. Keep the performance audit.
- Gateway suite including preservation/ambiguity/scope tests: **309 pass / 45
  files / 11.08 seconds**. Shared declarative writer: **9 pass**, including
  hidden-name and case-insensitive collision rejection with independent scopes.
- Browser navigation/library regression: **4 pass / 27.3 seconds**. Canonical
  add/delete names update Go To immediately and leave cell state unchanged;
  existing source-name navigation and shared function library/shortcut cases
  pass. An initial test asserted the wrong selection-context shape; correcting
  the assertion was necessary, with no production change for that failure.
- Native VS Code name scenario: **1 pass / 12 seconds**. An `agent`-origin
  authoritative transaction updates Go To, preserves cell overlays, supports
  exact native Undo/Redo and saves both scopes; a distinct-session reopen retains
  navigation. This exercises the host transaction path, not a real Chat/model.
- Native tool schema regeneration/plugin boundary/build: **pass / 48.78-second
  Vite build**. Desktop main/preload/renderer build: **pass / 62-second renderer
  build**. Full VS Code type chain and CLI type check pass. Gateway's stricter
  exact-optional check exposed an input type mismatch, corrected without changing
  runtime behavior; its final result and broader browser checks follow below.

Hidden-name collisions are blocked at save; earlier edit preflight still needs
source metadata support. Shared Name Manager UI, reference-aware rename, scope
calculation, Create from Selection and Use in Formula remain open. No commit,
push or VSIX packaging was performed in this checkpoint.

Final scoped-definition checks: gateway exact-optional TypeScript **pass** after
the type-only correction; standalone CLI type check and **14 DSL tests pass**,
including explicit scoped-operation rejection. Native agent tools/boundary:
**17 pass / 2 files**. Changed code lint/Prettier and `git diff --check`: **pass**.

Inspection of the pinned desktop engine confirms Name Manager rename/removal
are reference operations: its update-formula controller rewrites matching name
tokens, with removal producing `#REF!`. The direct definition primitives added
here are not claimed to implement that UI behavior. Connecting the dialog must
include those authoritative reference plans; delete/add alone is insufficient.

Completed browser regression after the name projection/backend changes:
**88 pass / 7.3 minutes**, including the new canonical-name navigation scenario
and all prior Home/Insert/Page Layout/chart/clipboard/function scenarios. Together
with 158 session tests, 309 gateway tests, 14 CLI tests, 17 native-tool/boundary
units and the real native name save/history/reopen pass, this verifies the current
definition foundation. It does not complete the remaining Name Manager reference
operations, its shared UI, other Formulas groups or the overall desktop-menu goal.

## Shared Defined Names and reference replay (2026-10-27)

The actual desktop Defined Names group and Name Manager are now shared with the
plugin. This checkpoint adds semantic rename/removal, source reference replay,
bulk Create from Selection and Use in Formula. It does not establish complete
Formulas, calculation, locale or installed-platform parity.

- Focused browser acceptance: **7 pass / 47.4 seconds**, including rename/remove,
  rejected drafts, pending controls, Save waiting, distinct scopes, stale roots,
  formatted date labels, over-100-name atomic creation and top/left label modes.
- Full session tests: **160 pass / 35 files / 22.45 seconds**, with two workers.
  Real source/save/reopen covers far-away original formulas, action-time overlays,
  later formula creation, name reuse and visible/hidden local shadows. Source
  formulas using CDATA and foreign cell/worksheet extension payloads are covered;
  the final targeted source-save repeat passes after the self-closing-cell guard.
- Gateway tests: **312 pass / 46 files / 25.84 seconds**. Editor units: **34 pass**.
  Shared definition serialization: **10 pass**, including empty-source definitions.
- Native reader reservation test: **1 pass**. Hidden and empty bindings carry
  name/scope metadata without disclosing hidden formula content; release rebuilt.
- Native VS Code actual Name Manager: **1 pass / 30 seconds**. UI rejects a hidden
  binding collision before history changes, renames/removes original off-page
  references, supports exact native Undo/Redo, saves while preserving hidden and
  print names/package sentinel, then reopens in a distinct session with Go To.
- Plugin build and workbook/editor/plugin type checks pass. Focused lint,
  Prettier and Rust formatting checks pass. Build retains existing large-bundle
  warnings; the overall performance audit remains open.

Final acceptance after the structural/reference-log guard and cache-preservation fixes:

- Complete browser run: **98 pass / 1 interrupted startup / 7.9 minutes**. The
  interrupted isolated-candidate startup and all seven name cases subsequently
  pass in an **8-case / 34.5-second** repeat after builds finish. This covers all
  99 cases across the full run and repeat, not a single clean 99-case run.
- Final native actual Name Manager repeat: **1 pass / 17 seconds**, using the
  final plugin build, with exact native history and distinct-session reopen.
- Final session suite: **161 pass / 35 files / 23.99 seconds**, two workers under
  light load. The preceding concurrent-build run passed 160 cases but hit the
  existing 20-second history-storage timeout; the threshold remains unchanged.
- Final gateway suite: **313 pass / 46 files / 21.44 seconds**. Final shared
  serialization/CLI/boundary selection: **39 pass / 3 files**. Additional binding
  tests reject incomplete metadata and hidden-local destination shadowing without
  mutating the input root.
- Desktop and plugin final builds both pass (about 72 seconds for each renderer),
  as do final gateway/session types, focused host lint/format, theme-color check
  and diff whitespace checks. Cache invalidation now preserves foreign cell
  extensions and only visits sheetData cells. Source literal cells with a foreign
  formula/cache payload remain unchanged.

The initial sandbox browser attempt could not bind its loopback test server; the
repeat used local-server access. The first full run's isolated startup was
interrupted by concurrent build HTML reloads. These observations and the history
performance timeout remain recorded; no failure is silently counted as passing.
Structural changes with pending name-reference logs or reserved file-only names
currently reject before mutation, including after the final visible name is
removed. Their complete transforms, destination-shadow qualification, scoped
computation and broader locale/platform/appearance/menu acceptance remain open.

## Shared Formula Auditing (2026-10-27)

This checkpoint connects the actual desktop Formula Auditing group, shared
reference parser/float arrows, formula-view interceptor and Watch Window. It does
not complete Calculation or the full desktop-menu goal.

- Focused actual browser UI: **6 pass / 40.4 seconds**. Covers precedent boxes and
  arrows, complete-source off-page/cross-sheet dependents, removal without authored
  edits, imported and per-sheet formula view, Ctrl+Backquote, cyclic error
  navigation, watched values/formulas and pending values across tab changes,
  rejected view writes/retry, Save waiting and overtaken queries.
- Session suite: **165 pass / 36 files / 27.12 seconds**, two workers before heavy
  builds. Gateway suite: **313 pass / 46 files / 8.26 seconds**. Focused shared
  parser/search/audit units: **44 pass / 4 files**. New audit source doubles reject
  range reads, proving dependency/error scans use the complete index rather than
  viewport hydration; tests include source overlays/clears, stale roots and
  renamed original formula searches.
- Native VS Code actual Formula Auditing: **1 pass / 25 seconds**. Exercises the
  original C100001 precedent/dependent, source formatted value in Watch Window,
  D100001 error selection, unchanged authored roots for read-only commands, exact
  native formula-view Undo/Redo, saved showFormulas attribute, retained formulas
  and package sentinel, and a distinct-session reopen with formula view enabled.
  The extension host briefly reported unresponsive/responsive and Canceled logs
  during the run; the scenario and process exit succeeded. Retain the existing
  lifecycle/performance audit rather than treating these logs as resolved.
- CLI/native-tool/boundary selection: **32 pass / 3 files**, including the added
  real XLSX serialization of formula/grid/headings/zoom settings. Editor units:
  **34 pass / 12 files**. Desktop/plugin builds pass (about 85/81 seconds for their
  renderers); gateway, plugin including DOCX, editor, desktop and CLI types pass.
  Focused lint, Prettier, theme-color and whitespace checks pass.

The first browser attempts exposed pre-workbook service installation: obtaining
SheetInterceptorService before workbook plugin registration crashed initialization.
Imports now share the preset service, and the display interceptor installs after
workbook creation with per-projection disposal. The second attempt was stopped
once its identical page error was confirmed, rather than repeating every timeout.
A later three-pass/one-fail run exposed a test locator collision between the Data
ribbon tab and the Data worksheet tab; the scoped locator correction is included
in the six-pass run above. No timeout threshold was increased.

The complete 105-case browser run finished with **104 pass / 1 shape-draw failure /
5.8 minutes**. All six auditing cases passed within that run. The existing default
click-size/Shift-square/text-box shape case failed to produce its first shape;
three independent repeats of the unchanged test subsequently passed (**3 pass /
34 seconds**, traces retained). This does not establish a resolved root cause or
a single clean full-suite pass; retain the intermittent drawing/picker audit.

Current source inspection also confirms the generic candidate writable/protected
guard still rejects canonical view-setting writes on protected sheets. Desktop
Show Formulas changes its view without that guard. Aligning view-only permissions
while retaining content/protection restrictions remains required; the enabled
menu is not accepted as fully equivalent under every protection state.

Formula-result freshness, scope/destination qualification, larger-file query/render
performance, keyboard/locale/appearance/installed-platform acceptance, Calculation
and all remaining XLSX/DOCX menu groups remain open. Pending saved formula values
are already disclosed in all 21 locales; this notice does not prove calculation
freshness. The full desktop-menu goal remains active.

## Protected worksheet view grant (2026-10-27)

The prior protected-sheet Show Formulas gap is now addressed with an explicit
view-only host grant, without converting protected sheets into content-writable
sheets. Eligible pure view patches contain only showFormulas, showGridlines,
showHeadings and zoomScale. Normal readonly/content/mixed-print checks remain;
broader print/layout and readonly presentation parity is still open.

- Source/kernel/protected-save selection: **5 pass / 2 files / 1.99 seconds**.
  Tests wait for complete indexed protection, reject missing grants, readonly
  input and mixed content/print batches atomically, preserve input/inverse roots,
  and verify real XLSX save/reopen with unchanged protection/password, sheetData
  and custom package sentinel.
- Full session: **170 pass / 38 files / 25.74 seconds** with two workers before
  builds. Gateway: **313 pass / 46 files / 13.73 seconds**. Editor: **34 pass / 12
  files / 10.66 seconds**. Gateway/session/editor/plugin (including DOCX) types,
  changed lint/format and whitespace checks pass. Final plugin renderer build
  passes in **29.40 seconds**, with the existing bundle-size warning.
- Targeted actual browser interactions: **7 pass / 30.9 seconds**, including
  protected formula-view toggle on/off and retained rejection of formula writes.
- Final native VS Code: **2 pass / 31 seconds**. Ordinary and protected-source
  scenarios both exercise audit arrows, source error selection, formatted watch
  values, exact native view Undo/Redo, saved showFormulas/protection and a
  distinct-session reopen. The first two diagnostic runs failed because the
  strict candidate-request material schema had not yet accepted the new grant;
  both failed even for the ordinary sheet. The bridge now explicitly validates
  the field. Diagnostics exposed that actual error rather than raising a timeout.
- Final complete browser regression: **106 pass / 8.6 minutes**, one worker with
  retained failure traces enabled. Includes actual sidecar/Node/isolated-browser
  candidate commits and failed-projection recovery, every existing Home/Insert/
  chart/clipboard/Page Layout/Function Library/name/auditing case, protected view
  and the previously intermittent default-shape case. A clean run is not proof
  of the old shape/picker issue's root cause; that audit remains open.

Calculation integration remains open. Headless Node experiments with the same
Univer engine successfully computed numeric and cross-sheet values, boolean,
text/error-like values and a dynamic-array anchor. They also exposed missing
spill children in ordinary result data and ambiguous raw string/error types.
See calculation-notes.md; these experiments are not a production calculation
service or completed Calculation menu. The full XLSX/DOCX menu goal stays active.

## Normalized-snapshot calculation adapter (2026-10-27)

A host-only adapter now computes a supplied authoritative, hydrated snapshot using
Univer core/sheets/formula plugins, without a DOM. This is a calculation primitive,
not completed Calculation-menu or document-cache integration. No new ribbon
command is enabled on this path yet.

- Final real bundled Node/CommonJS adapter selection: **13 pass / 1 file / 1.32
  seconds**. Tests cover numeric/boolean/error/error-shaped text, cross-sheet
  references, complete dynamic-array/reference spills and per-element errors,
  blocked versus literal-SPILL results, scalar-empty versus spilled-blank
  semantics, date-format hints, hyperlinks and IMAGE metadata, immutable input,
  output quota, empty jobs, cancellation before/during execution and at START,
  subsequent reuse, concurrent reused caller IDs and Calculate Sheet behavior.
- Final full workbook-session suite: **183 pass / 39 files / 28.88 seconds**, two
  workers. Package typecheck/build, changed-file lint/Prettier and whitespace
  checks pass. Added direct calculation-plugin and esbuild test dependencies
  already exist in the workspace lock graph; lock-only install used offline mode
  with scripts disabled. This is not an external dependency version upgrade.
- CommonJS test artifacts are built through esbuild and loaded in a real Node
  process environment without document/window services. The test does not rely
  on stubs for formula results or a browser display calculation.
- Raw worksheet/result values could not distinguish true errors from error-like
  text, and omitted spilled children. The adapter observes the public runtime
  writer, forwards the original method and retains rich values during the same
  evaluation. The attempted public dependency-manager tree API is unimplemented
  in this version, so it is not used. No error-type probe re-evaluates formulas.
- Reused logical unit IDs reproduced incorrect SUMIF values in concurrent jobs;
  assigning each engine invocation a private UUID fixed that case while keeping
  caller identity in the result. Each result declares workbook/sheet scope and
  has stable sheet/coordinate order. A sheet request uses the desktop dirty-range
  and dirty-name semantics, updates off-sheet dependents and leaves unrelated
  cached formulas outside the result. It does not certify global freshness.
- Empty-reference inspection found that normalizing a reference to an array can
  change its single-blank behavior: the runtime writer publishes numeric zero
  while that array exposes a null object. The collector now uses the same public
  getFirstCell/iterator operations as the writer; a dedicated test covers both
  single-reference zero and spilled blanks. Blocked spills are distinguished
  through the runtime's accepted clear-array data, not the first value's text.

The service restores hooks, stops failed/cancelled jobs and disposes the engine;
its output limit is not a preemptive memory/CPU bound. Complete source hydration,
hidden/scoped names, tables/external resources, worker ownership and limits,
frozen-root validation, authoritative derived-cache publication, 1900/1904 dates,
locale/time-zone context and exact source/array OOXML save semantics remain open.
The current Calculation ribbon, auto/manual mode, F9/Shift+F9 and cache/read/search/
watch/chart/preview/history/save/reopen integration still require implementation
and native/platform verification. See calculation-notes.md. The full goal remains
active; these tests do not establish completed desktop-menu parity.

## Complete calculation-source loading (2026-10-27)

- New loader tests: **19 pass / 2 files**, including two actual XLSX/native
  protocol cases. One has styled blanks, exact empty strings/errors, arrays and
  formulas beyond row 100000 with understated metadata; another crosses the
  native 10000-result page boundary and verifies inserted rows and overrides.
- Full workbook-session regression: **202 pass / 41 files / 48.29 seconds**,
  two workers. Full gateway regression: **313 pass / 46 files / 22.71 seconds**.
- Full native Rust regression: **276 pass**, plus empty main/doc-test targets.
  Release sidecar rebuilt with the new internal enumeration mode. Eight native
  Find tests include one-cell pagination of styled blanks and array metadata.
- Workbook-session typecheck/build, changed-file ESLint and whitespace checks
  pass. No renderer, plugin UI, installed VSIX or Windows acceptance is claimed
  for this host-only source-loader increment.

Loading rejects incomplete/foreign/non-advancing/duplicate enumeration, missing
range records, stale ownership, cancellation and exceeded operation budgets.
Source features/arrays remain explicitly in original coordinates; logical data
and journals are separate. Private hidden/scoped definitions, table/external
resources, array transformation, full date/default/function context, worker
bounds and authoritative derived caches are still incomplete. Calculation menu,
auto/manual mode and history/save/reopen/platform acceptance remain open; this
checkpoint does not establish completed desktop-menu parity.

## Private workbook calculation context (2026-10-27)

- Actual native source-context/loading tests: **22 pass / 3 files**. New real
  XLSX scenario verifies global/local/hidden/empty/built-in names, calcPr settings,
  stable source scope after reordering/renaming, and absence of hidden formula
  text from ordinary open metadata and canonical state. Closed/foreign sessions
  are rejected. Source tests verify context budgets and stale/cancelled ownership.
- Full workbook-session regression: **205 pass / 42 files / 29.05 seconds**,
  two workers. Full native regression: **279 pass / 14.01 seconds**, plus empty
  main/doc targets; release sidecar rebuilt successfully.
- Three new native tests cover private definitions, prefixed namespaces with
  CDATA/entities, malformed scopes/nested markup/relationship namespaces,
  cancellation and closed reads. Existing public metadata remains unchanged.
- Workbook-session typecheck/build, changed TypeScript ESLint/Prettier and
  whitespace checks pass. No new UI or installed VSIX/Windows result is claimed.

The private resource command reads the owning session's file path, never a path
provided by the renderer. Private formulas are not added to public metadata or
state. The loader retains raw original scope/text/settings and unresolved external
IDs separately from logical edit journals. This does not establish calculation
binding or mode behavior: installed Univer name lookup caches omit worksheet
scope, so merely registering localSheetId is insufficient evidence for same-name
local/global resolution. Resource transformation, tables/arrays/external/date/
default semantics, worker limits and canonical calculation/cache/menu/history/
save/reopen/platform work remain open. The full desktop-menu goal stays active.

## Normalized calculation name binding (2026-10-27)

- Real bundled Node/CommonJS calculation suite: **23 pass**. Ten added cases
  cover same-spelling local/global and qualified names, Unicode/global fallback,
  nested expression precedence, lexical LET/LAMBDA scope/named calls, direct
  dynamic INDIRECT names, quoted/workbook qualifiers, local range arrays,
  concurrent jobs/caller-definition mutation, unary/percent/array LET behavior,
  shared follower provenance and invalid/circular/ambiguous bindings.
- Full workbook-session regression: **215 pass / 42 files / 34.32 seconds**,
  two workers. Workbook-session typecheck/build, changed-source ESLint/Prettier
  and whitespace checks pass. No native sidecar or renderer files changed in
  this increment; no new installed VSIX, UI or Windows acceptance is claimed.
- A tokenizer inspection confirms that -Rate is a single FUNCTION-type token,
  whereas -(Rate) separates the operator/operand. The actual negative-name test
  failed with a raw negative alias and passed after explicit operand grouping.
- A named LAMBDA calculation returned VALUE with ordinary named-expression
  grouping and passed using the engine's inline LAMBDA(...)(args) form. Shared
  follower tests verify original relative formula text and absence of private
  aliases in returned values, while leaving the input SI/master records intact.

These checks prove normalized logical definition binding for the tested cases,
not complete source/canonical formula authority. Source definitions/scope IDs,
structural/reference/definition journals still need a verified composition step.
Recursive/iterative names and arbitrary dynamic named expressions remain open;
circular static bindings currently fail explicitly. Table/source-array/external/
date/default context, worker limits and canonical cache/menu/mode/history/save/
reopen/platform acceptance remain required. The complete desktop-menu goal stays
active. Scope reference and implementation boundaries are in calculation-notes.md.

## Source definition/journal composition (2026-10-28)

- Real bundled Node/CommonJS calculation suite: **24 pass**; the added case
  covers Excel 16.0 global-definition lexical scope. Name-binding focused suite:
  **8 pass / 2 files**. Source-definition composition tests: **3 pass**, plus one
  real native sidecar/save/reopen integration case; the targeted composition
  run reports **4 pass / 2 files**.
- The integration fixture carries source formulas beyond row 100000, hidden
  dependent definitions, a hidden CDATA string containing the literal text
  `&amp;`, a same-spelling local name, sequential rename/recreate actions, and a
  later action-time formula. The normalized engine returns `[67.5,
"&amp; Revenue", 26.5]` both before and after the real package writer saves and
  the native sidecar reopens the workbook. The test also checks inverse-state
  fidelity and that hidden formula bodies do not enter canonical state.
- Actual Windows Excel 16.0 COM output confirms that the tested global nested
  name is serialized as `=Rate*2`, while a definition bound to the active
  worksheet's local name is serialized as `=Other!Rate*2`; the observed results
  are 20 and 40 respectively. A separate saved-fixture Excel acceptance attempt
  is unverified because the Windows SSH endpoint refused the connection on
  2026-10-28. This is not a successful desktop Excel open/save check.
- Direct workspace TypeScript checks pass for both workbook-session and
  xlsx-gateway after the source-scope fallback correction. `git diff --check`
  also passes. Full regression suites and package builds have not been rerun in
  this checkpoint; the pnpm launcher failed before execution because its SQLite
  store could not be opened, so the installed `tsc` binaries were invoked
  directly.

Source name composition is now implemented for tested immutable-source,
reference-journal and action-time Name Manager cases. The structural-operation
guard remains in place for name-reference bindings; the pure structural
transform case does not prove those operations are accepted. Legacy state scope
recovery, prefixed writer XML edge cases, tables/arrays/external/date/default
resources, worker isolation and canonical Calculation cache/menu/history/save
integration remain incomplete. The overall desktop-menu goal stays active.

## Structured-table resource binding slice (2026-10-28)

The private source model now extracts table ranges/names/column maps, and the
Univer adapter registers representable tables through `ISuperTableService`.
Workbook-session TypeScript compilation passes after this change. No test suite
or native Excel table fixture was run in this slice. The implementation rejects
totals rows, non-single-header tables and `#Totals` formula references because
the installed engine's public table model cannot represent those semantics.
Actual structured-reference evaluation, calculated columns and table
transformations remain unverified and open.

The same slice adds a normalized snapshot builder and a source evaluation
orchestrator. Workbook-session TypeScript compilation passes with these files.
The builder maps formula text, scalar input values, forced-string literals and
number-format hints, and carries the logical name/table resources to the engine.
It fails closed for 1904 workbooks, legacy array ranges, literal error cells and
external-workbook formulas. No tests were run for the snapshot/orchestrator
change; row/column defaults, full style/date semantics and canonical result
publication remain unimplemented.

## Plugin Calculation ribbon checkpoint (2026-10-28)

The plugin's Formulas tab now includes the shared Calculation Options menu,
Calculate Now and Calculate Sheet commands. Runtime dispatch uses the shared
desktop automatic/manual calculation hook and sheet dirty-range/name request;
F9 and Shift+F9 use the same paths and are scoped away from cell editing,
inputs and dialogs. The `sheets-editor` and desktop `apps/sheets` TypeScript
checks pass; changed-file Prettier and whitespace checks pass. No runtime,
browser or native acceptance was run in this checkpoint.

The mode is runtime-local and projection-only. It is not persisted in source
`calcPr`, and the buttons do not yet call the source snapshot/evaluation service
or publish canonical formula caches. Menu rendering and command wiring
therefore do not establish document-level Calculation parity; native history,
save/reopen and real formula-result freshness remain open.

## Shared desktop Sort & Filter group and filter lifecycle commands

The actual desktop Data → Sort & Filter group is extracted inside SheetsRibbon
and rendered by both desktop and plugin. The shared Sort dropdown, glyphs, option
order, Filter tool, Clear/Reapply/Advanced stack and translated shortcut detail
are preserved; a common LargeRibbonMenu keeps the other desktop menu helpers on
the same rendering path. The plugin removes its duplicate simple-sort/header
controls. Simple A→Z/Z→A now sort the explicitly selected rows without an assumed
header, matching the desktop range.sort({column:0}) command; the Custom Sort
dialog retains its own default checked header option. The temporary value-list
editor/removal command remains under additional controls until native header
filter popups are integrated. The rest of Data is still a partial panel.

The canonical clear_filter_criteria operation removes all column conditions and
restores visibility while retaining the range. reapply_filter preserves every
condition and recomputes hidden rows from current source/overlay values. Both
flow through the same session material/validation, desktop AI execution, native
history, original-based package writer and generated AI tool schema. Reapply
retains formula-cache preflight; clear-criteria does not need formula evaluation.
The CLI can clear imported conditions and reapply conditions created in its
current batch; it retains the refusal to reconstruct unknown imported criteria.
It also rejects imported filter operations mixed with coordinate changes rather
than silently reusing an old range. Existing remove-filter behavior is distinct.

Filter toggle reads current metadata at the captured root, removes an existing
filter or uses the exact desktop expandToContinuousRange algorithm to choose a
new range. The browser adapter supplies complete source halos, not viewport-only
cells; each read stays within 2,000 cells and captures revision/cancellation.
Single-cell selections expand in all directions, header-only selections downward,
and explicit multi-row ranges remain explicit. Merged-anchor geometry is provided
through the same expansion algorithm. Pending source reads/planning/writes share
a single-flight work barrier, participate in Save, and cannot commit after an
overtaking edit. Clear without conditions or filter is a no-op. Advanced on a
sheet without a filter first creates the smart range, as the desktop does.

Ctrl/Meta+Shift+L uses the same toggle handler from the workbook view; native
fields, active cell editing, composition and open dialogs keep their keys. Linux
VS Code proves actual Ctrl+Shift+L; macOS secondary aliases and full host-focus
precedence remain pending. The native test also exposed a shared Dropdown issue:
a wrapping label could activate the trigger again after picking and removing its
option. Preventing default label activation on the option click keeps the selected
list closed. The native test asserts aria-expanded=false before using the next
condition, and the existing 15 Dropdown tests pass.

Four source-region tests pass (offscreen row 50001, blank separators, header-only
width, explicit range, merged anchors, stale read and budget rejection). All 49
session ribbon tests pass, including unchanged definitions after value edits and
one-step reapply/clear; CLI filter create/clear/reapply and structural refusal
checks pass. Six browser regressions pass for cross-tab edits, shared sort/menu
choices, Custom Sort retry, Advanced AND/OR, shortcut/dialog guards, smart toggle,
clear/no-op, current-value reapply and stale read rejection during Save. A final
four-case run after the Dropdown fix passes as well.

Final native VS Code 1.109.5 acceptance passes all three filter controls, Custom
Sort and Advanced Filter scenarios (30 seconds). A single-cell shortcut creates
A1:B4 while preserving a separate table at A80; an AI B3 edit uses the same
history, Reapply reveals its now-matching row while retaining conditions, native
Undo/Redo restores exact roots, and Clear keeps the range. Filter off/on history,
repeated saves and a distinct clean-session reopen preserve the AI value, source
style and distant source text; an imported cleared filter can then be removed
and saved. Custom Sort and imported Advanced criteria native checks remain green.
The final run emits no bare Canceled diagnostic.

Desktop/plugin/CLI TypeScript, changed-file lint, full extension build/browser
boundaries, source inventory and whitespace checks pass. Current inventory is
20 modules, 1,126 syntax evidence entries and 76 dynamic dispatch/option entries;
those are not completion counts. Full header popups, complete Data band, larger
filter ranges beyond 2,000 cells, broad formula/empty/merge/date/filter/protection/
locale/appearance/lifecycle cases and installed Windows/macOS acceptance remain
open. This checkpoint does not close overall desktop-menu parity.

## Shared Advanced Filter and canonical custom criteria

The actual desktop AdvancedFilterDialog is mounted in the plugin Data panel.
It retains one-column, six-operator, optional second-condition and AND/OR
parameters, empty-first-value validation and the existing 21-locale labels.
The controlled plugin locale and chrome-only Dropdown import avoid the desktop
AI UI barrel. Async Apply disables editing/dismissal and keeps inputs and inline
errors on rejection so the user can correct or retry the same captured context.

Opening flushes edits, reads the complete current filter metadata from the host,
then reads the first 26 header cells as the desktop does. It uses the active
worksheet even without a prior cell click. Source-read metadata includes original
or current canonical filter definitions; the dialog captures the worksheet/root/
revision/filter range and cannot apply after that version is overtaken. Pending
apply work participates in Save flushing, including before the queued write.
The original desktop dialog still supports its synchronous callback.

set_filter_criteria now accepts exactly one of legacy values (including null to
clear a column) or validated one/two-condition customs. Session candidate,
desktop execution and CLI serialization support that shape. A host-neutral
predicate is shared by desktop streamed filtering, canonical sessions and CLI
new-filter writes: raw numeric comparisons, default-equality wildcard text,
boolean conversion, literal/wildcard notEqual, and AND/OR composition. Explicit
equality normalizes to the OOXML default just as the desktop dialog does.
Other columns' existing value/custom conditions are preserved and combined.
Candidate validation refuses filtered-column formulas without ready/error caches
or content changes earlier in the same batch that would invalidate their results.
The CLI retains its existing refusal to edit imported criteria without recreating
the filter; this checkpoint does not claim full imported-filter CLI parity.

Current range-row and complete-outline reads now compose canonical filter hiding
before explicit row-visibility overrides, matching grid construction and writer
ordering. Complete source outlines therefore reflect filtered detail even outside
the hydrated viewport. Existing source/size/cell properties remain unchanged.

Verification: three gateway predicate/schema tests, 48 session ribbon tests,
15 desktop advanced/streamed-filter tests, two CLI filter tests, and 13 current-
range/outline tests pass. Two browser tests cover empty-value validation, numeric
AND, wildcard OR, retained rejected criteria, pending Apply/Cancel/Escape guards,
Save waiting, exact row visibility, cancel and stale-dialog rejection. Native
VS Code 1.109.5 acceptance applies >=10 AND <20 to raw numbers (formatted 12.50
matches while literal text "12" does not), proves one revision and exact native
Undo/Redo, verifies saved customFilters/hidden-row XML, and independently reopens
into a distinct clean session. The original values, numeric type, format/fill,
other worksheet and XML sentinel read back unchanged. Adding a condition to
another column after reopen preserves the imported first-column conditions;
second-stage native history and saved XML are verified as well.

The final combined native run passes all three Advanced Filter/Outline scenarios
(41 seconds). The Outline scenario still emits bare Canceled diagnostics during
view replacement; the outstanding lifecycle audit remains open. Extension and
desktop typechecks, changed-file lint, inventory consistency, whitespace and full
extension build/browser boundaries pass. Inventory now covers 20 source modules,
1,124 syntax evidence entries and 76 dynamic dispatch/option entries requiring
review; those counts are not completed commands.

The 2,000-cell filter limit, broader formula/date/error/hidden/protection/appearance/
locale/platform variants, native header filter popup, Filter toggle/Clear Criteria/
Reapply adapters and complete desktop Data ribbon remain open. This is the
reference application's one-column custom AutoFilter dialog, not a claim of the
separate full Excel criteria-range/copy-to workflow.

## Shared Remove Duplicates dialog and native acceptance

The desktop RemoveDuplicatesDialog is extracted from ExcelShell and used by both
applications. It retains the desktop description and default checked header flag,
with a controlled locale for the plugin. Async apply uses a synchronous pending
lock, disables the controls, blocks dismissal during submission and keeps both
the chosen header option and inline error on failure so the user can retry.
The desktop reference compares every selected column and offers only a header
checkbox; it does not have a column picker.

The plugin captures the worksheet, range, root and revision after flushing when
the dialog opens. Apply reads that frozen selection rather than the latest cursor
position, rejects overtaken reads or dialog contexts, shares the desktop
case-insensitive typed-row comparison and submits changed scalar values as one
canonical transaction. Single-flight work participates in Save flushing before
the source read as well as during submission. Rejected dialog transactions permit
retry. No duplicate rows means no transaction, with the desktop result message.

Four browser cases pass: retained no-header choice after rejection and retry,
typed boolean/text distinction, preservation of a stationary formula and target
style, disabled pending Apply/Cancel and blocked Escape, Save waiting for writes
and source reads, cancel/no-op without revisions, stale dialog and concurrent
read rejection, and merged-selection rejection before any submission.

Native VS Code 1.109.5 acceptance passes through the actual shared dialog with
source row data and known ready formula caches. Two duplicates compact into one
revision; exact native Undo/Redo restores both roots. An unchanged row keeps its
cross-sheet formula; a moved FALSE() formula becomes a boolean false as on the
desktop; trailing source cells clear. Independent save/close/reopen in a new,
clean revision-zero session reads the header, compacted values/types, preserved
target fill, unchanged formula, unrelated date serial/date kind/format, other
worksheet source value and custom XML sentinel. The fixture explicitly seeds
known caches (Other!A1 = 7 and FALSE() = false); the pending-cache guard is not
weakened to accept an unevaluated formula.

Plugin and desktop TypeScript, changed-source lint, the full extension build and
browser boundary checks pass. The inventory generator now follows the extracted
RemoveDuplicatesDialog as a separate source (19 modules, 1,110 syntax evidence
entries; these are not counts of implemented commands). The 2,000-cell selection
limit, broader sparse/date/error/protection/filtered-row variants, all appearance/
locale and installed Windows acceptance remain open. The complete shared Data
ribbon, Get Data/Pivot, Subtotal and the other menu families are still unfinished.

## Plugin Flash Fill ribbon checkpoint (2026-10-28)

The plugin Data tab now exposes Flash Fill. The action reuses the desktop
`inferFlashFillTemplate` / `applyFlashFillTemplate` algorithm, reads up to six
preceding columns and the selected target column from the host using a frozen
state/revision, and writes inferred strings as one canonical transaction. A
single-cell selection follows the contiguous source rows to its left, matching
the desktop command; larger reads are chunked and the output is limited to 2000
rows. Missing/pending/error formula results, merged cells and incomplete reads
fail closed. Typecheck and changed-file lint pass; no tests or UI acceptance
were run. Pattern coverage remains the shared common subset (concatenation
templates), and interactive undo/save/reopen behavior still needs acceptance.

## Plugin Text to Columns ribbon checkpoint (2026-10-28)

The plugin Data tab now exposes the desktop delimiter presets for comma,
semicolon, space and tab. The command reads one selected source column at a
frozen revision, applies the same literal delimiter splitting behavior as
Univer's desktop `splitRangeText` helper, detects merged output ranges, inserts
worksheet columns when the split reaches the current grid edge, and writes the
result in one canonical transaction. It fails closed on formulas without ready
results, error cells, selections/output above 2000 cells, or more than 100 new
columns. Quoted CSV fields are not parsed by the desktop helper either. No tests
or UI acceptance were run; undo, save/reopen and overwrite behavior still need
interactive acceptance.

## Plugin Data Validation ribbon checkpoint (2026-10-28)

The Data tab now provides a plugin-owned editor for list, same-sheet list
reference, number-between, date-between, checkbox and custom-formula rules, plus
clear validation. Applying a rule uses the canonical `set_data_validation`
operation against the selection captured when the dialog opened; a revision
change makes the action fail and asks the user to reopen it. The free Univer
data-validation preset is installed, and newly created rules are applied to the
current grid after the host transaction commits. TypeScript and changed-file
lint pass; no tests or UI acceptance were run.

This is a supported-rule subset, not full desktop Data Validation parity. The
plugin does not yet hydrate existing XLSX validation rules into its live grid or
editor, nor expose input/error prompts, all Excel criteria, or the full native
rules panel. Existing document validation preservation and save/reopen behavior
remain unverified in the VS Code editor.

## Data Validation session persistence follow-up (2026-10-28)

Source review found that `set_data_validation` was accepted by the shared DSL
schema but had no branch in the session candidate interpreter, so the first
plugin dialog implementation could not actually commit its rule. The session
candidate now journals the translated Univer wire rule and exact range as
`dataValidationEdits`; the session package writer passes those edits to the
gateway's existing append/remove DV writer. Existing unrelated rules are
preserved by append mode, while replacing/clearing the same exact area is
represented in the remove list. CLI and session use the same
`dvRuleToWireRule` mapping now. Gateway, workbook-session, editor and CLI
TypeScript checks plus changed-file ESLint and whitespace checks pass. No tests
or actual XLSX save/reopen acceptance were run.

## Data Validation overlap-preserving patch (2026-10-28)

The gateway's append/remove writer now subtracts the union of replacement and
removal rectangles from each existing `sqref` rectangle. It keeps untouched
rules verbatim, emits each surviving fragment with its own `sqref`, and offsets
relative formula references from the original area's top-left to each fragment's
new top-left. Absolute references remain pinned through the shared formula
offset routine. Gateway TypeScript, changed-file ESLint and `git diff --check`
pass. No tests were run; multi-area formula semantics and native Excel
save/reopen still require acceptance.

## Column-header filter popup (2026-10-28)

The plugin now mounts the desktop grid's Univer filter preset and CSS instead of
using a separate column-number/value-textarea editor. Imported and canonical
filters restore their range and absolute model columns without recalculating a
partially hydrated grid. Filter-hidden rows belong to the filter model; explicit
manual/outline visibility remains in row hd flags, so clearing criteria can
reveal filtered rows.

The command adapter is installed after Univer reaches Rendered. Installing only
after the workbook projection's first animation frame was too early: the filter
UI had not yet registered its commands. The lifecycle-bound adapter freezes
root/revision/sheet/column, reads the complete bounded table before opening, and
refreshes authoritative values when switching back from conditions. Formula
references remain present in the projection while their host results supply the
value list. A pending formula, overtaken read or table above 2,000 cells refuses
opening. Native value/blank/custom predicates and column clearing submit through
the host session, and Save waits for reading/submission. Rejection preserves the
popup draft; retries clear the previous error and repeated pending clicks cannot
submit a second transaction.

Evidence so far:

- `filter-popup.test.ts` and `filter-region.test.ts`: 7 unit tests pass for offset
  conversion, blanks, custom predicates, unsupported criteria, source expansion,
  stale reads and range budgets.
- Six filter browser tests pass in `/tmp/genoffice-filter-popup-browser-final.log`:
  the existing Advanced Filter/Sort & Filter/Reapply checks plus native popup
  value application, failed retry, pending double-click/Save and complete
  offscreen/stale-read behavior.
- The formula/condition browser case passes in
  `/tmp/genoffice-filter-popup-browser-formula-final.log`, including a formula
  referring outside the loaded viewport and switching conditions → values →
  conditions before applying Greater Than.
- Final production extension build and browser-boundary validation pass in
  `/tmp/genoffice-filter-popup-build-complete.log`.

- The cross-feature seven-tab browser regression passes after replacing the
  legacy textarea interactions with native header checkbox selection:
  `/tmp/genoffice-filter-popup-browser-crossfeature.log`. Final popup value,
  offscreen/stale and formula/conditions cases also pass (3 cases in
  `/tmp/genoffice-filter-popup-browser-complete.log`; that run exposed the old
  cross-feature test's assumed value ordering, which the isolated regression
  corrected by selecting the checkbox by its actual text).
- Real VS Code 1.109.5 header checkbox/history/save/reopen acceptance passes:
  `/tmp/genoffice-filter-popup-native-viewport.log` (1 passing, 10s). The CDP
  workbench viewport is 1200×900. This checks one revision per confirmed filter,
  exact native Undo/Redo states, saved filter XML, unrelated custom XML/style
  preservation and independent clean-session reopen. The final run has no bare
  Canceled diagnostic; this does not close the separate global lifecycle audit.
- Editor/extension type checks, changed-file lint/formatting, source-inventory
  consistency and `git diff --check` pass.

This is partial filter parity: canonical color criteria/persistence,
large tables, all locales/appearance/protection/imported-table variants and
installed Windows acceptance remain open. No new VSIX or Windows build is claimed.

## Desktop filter-panel sorting and authoritative visible rows (2026-10-28)

The editor now mounts the actual Univer Sheets Sort preset used by the desktop
app. Its `EmbedSortBtn` is registered at the filter panel embed point; the
Ascending/Descending buttons keep the desktop header-excluding range and the
clicked absolute key column. `SortRangeCommand` is adapted to one frozen
canonical `sort_range` submission. The stock component closes immediately after
starting a sort, so the command gate prevents closing while the host is pending;
a rejection retains the popup and allows retry. Save includes that pending work.

The previous `active_filter_sort` rejection has been replaced by desktop visible
slot semantics: the host materializes complete current row visibility for every
sort span, sends the bounded row list through the strict candidate schema, and
the shared value and style permutations use the same selected row indices.
Filtered and manually hidden slots are not rewritten. Filter edits in earlier
operations of the batch and explicit visibility journals are replayed when
selecting rows. Creating/editing a filter for the first time also journals
source-manual hides not owned by live original filter criteria, preserving them
through filtering, sorting and original-based serialization.

Verification:

- Session candidate/material tests: **55 passing** in
  `/tmp/genoffice-popup-sort-session-accepted.log`, including noncontiguous visible
  row movement, typed/format/style preservation, source hidden-row reads and
  source-manual hide preservation at first filter creation, and clearing criteria
  before sorting in the same transaction despite an older host visibility cache. Sort domain
  regressions: **5 passing** in `/tmp/genoffice-popup-sort-domain-tests.log`.
- The new browser test verifies numeric Ascending/Descending with active
  criteria, manual hiding, host rejection/retry, pending Save and moved styles.
  `/tmp/genoffice-popup-sort-browser-regression.log` has seven passing cases and
  exposed a shortcut test that sent keys before Escape restored dropdown focus;
  waiting for actual restored focus verifies that behavior. The corrected
  shortcut and popup-sort cases both pass in
  `/tmp/genoffice-popup-sort-shortcut-final.log`. Existing header/condition,
  formula-cache, Advanced Filter, Reapply and seven-tab cases are covered in
  the regression run.
- Real VS Code 1.109.5: **1 passing (8s)** in
  `/tmp/genoffice-popup-sort-native-accepted.log`. A source row is hidden in the
  original XLSX before opening. The test applies filter criteria through the
  header, sorts by the other column, verifies filtered/manual stationary slots,
  exact native Undo/Redo, saved filter/hidden-row XML, unrelated custom XML,
  numeric formats and styles, and an independent reopened session with revision
  zero and a clean document. Untouched hidden cells are read through the
  authoritative snapshot rather than assumed to exist in the sparse overlay.
  The final focused run has no bare Canceled diagnostic; the earlier combined
  run did emit one, so the wider lifecycle audit remains open.
- Existing native Custom Sort and header filter tests pass in
  `/tmp/genoffice-popup-sort-native-final.log` against the same production
  implementation; only the new test's sparse-overlay assertion needed repair.
- Full production extension build and browser boundary checks pass in
  `/tmp/genoffice-popup-sort-build-final.log`; type checks, changed-file lint,
  formatting, source inventory consistency and whitespace checks pass.

Remaining: color criteria/persistence, formula movement, merged/formula/error
sorts, notes/hyperlinks/validation movement, additional sort types and hidden-row
provenance in imported mixed filter/manual cases, all appearance/locale and
installed Linux/Windows variants, and larger data budgets. This evidence accepts
the implemented popup sorting workflow, not full Data/menu parity or a new VSIX.

## Bounded desktop color-filter workflow (2026-10-29)

The actual desktop By Colors panel now commits fill or font criteria instead of
refusing them. The shared schema accepts exactly one of values, custom conditions
or colors; one color criterion selects a fill palette (including null for no
fill) or a font palette. Native RGB strings normalize to hex for the canonical
state, and restoration converts them back to native RGB. Default font black and
no-fill/explicit-white differences are retained. Other columns still combine
with AND, and Reapply reads current styles. The color predicate does not require
formula values to be ready, although opening the shared panel retains its existing
source/value readiness guard and is not yet accepted for every pending-formula
UI workflow.

Single-color persistence uses the standard colorFilter dxfId and cellColor
attributes described by [Microsoft's ColorFilter documentation](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.colorfilter?view=openxml-3.0.1).
The complete desktop selection is stored under the supported filter-column
[extLst extensibility point](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.filtercolumn?view=openxml-3.0.1)
with URI `urn:genoffice:filter-colors:v1`. A multicolor criterion is preserved in
that extension with its cached hidden rows; it is not converted to a misleading
single-color condition. External Excel interpretation/reapplication of this
GenOffice extension has **not** been verified. The native reader understands
standard DXFs and this extension, preserves column offsets, and caps palettes to
the canonical 2,000-color limit.

The shared gateway stylesheet editor interns color DXFs, preserving prior DXF
indices. A real reopen test exposed solid fills being read from bgColor instead
of fgColor; the native DXF reader now prefers foreground for solid patterns and
retains background fallback for ordinary differential fills. Colors also travel
through desktop save/restore and AI-operation handling, CLI serialization, and
regenerated VS Code native tool schemas. Schema/matching logic stays independent
of the Univer presentation dependency; only native RGB conversion imports the
renderer color utility.

Evidence:

- `color-filter.test.ts`: 2 passing tests for native/canonical conversion,
  multicolor/no-fill/default-black/theme resolution, standard DXF serialization,
  multicolor extension serialization and preserved cell XML.
- `filter-popup.test.ts`: 4 passing tests, including offset-column fill/font
  conversion and the existing value/custom boundary checks.
- `color-filter-save.test.ts`: 1 real-sidecar integration case with five
  independently saved/reopened configurations: red fill, red/blue/no-fill,
  no-fill alone, black/blue fonts and red font. For single selections the test
  removes the GenOffice extension before reopening, proving standard DXF import.
  It verifies cached hidden rows, unchanged cells/styles and custom XML.
  `/tmp/genoffice-color-save-final.log` and
  `/tmp/genoffice-color-session-final.log` (51 passing across save/candidate
  suites) contain this evidence.
- A further cross-column/style-edit/pending-formula predicate case passes with
  the candidate suite (51 passing) in `/tmp/genoffice-color-mixed-test.log`.
- CLI fill/no-fill and font criteria, lowercase normalization and saved XML pass
  in `/tmp/genoffice-color-cli-test.log`; core conversion/serialization tests pass
  in `/tmp/genoffice-color-unit-final.log`.
- Native Rust standard/extended filter import passes in
  `/tmp/genoffice-color-rust-test-final.log`; three DXF/conditional/table-style
  regressions pass in `/tmp/genoffice-color-dxf-regression.log`. The release
  sidecar builds in `/tmp/genoffice-color-rust-build-accepted.log`.
- Four browser cases pass in `/tmp/genoffice-color-browser-final.log`, covering
  value filtering, authoritative formula/condition filtering, visible-slot
  sorting and color fill/font/multiple/no-fill/retry/style-Reapply/clear.
- Real VS Code 1.109.5: 2 passing (24s) in `/tmp/genoffice-color-native.log`.
  The new case checks one canonical revision, exact native Undo/Redo, standard
  color/DXF XML, independent clean-session reopens, restored checked swatches,
  editing an imported single color into multiple colors, restored multiple
  choices after another reopen, clearing, preserved source values/styles and
  unrelated custom XML. Startup briefly reported an unresponsive extension host
  and then recovered; this run does not establish a startup-performance budget.

- After separating the Univer RGB adapter from pure protocol/matching code,
  the final production build/boundary checks pass in
  `/tmp/genoffice-color-build-final.log`; the native imported-color/edit/clear
  case passes again (1 passing, 17s) in `/tmp/genoffice-color-native-final.log`.
  This focused run has no unresponsive-host diagnostic, but is not a performance
  benchmark. Native tool schemas explicitly expose the bounded fill/font arrays
  and hex color pattern; final type/lint/format/inventory/whitespace checks pass.

Remaining color acceptance: external Excel multicolor interoperability,
rich-text runs, conditional/table/gradient/pattern/effective-theme colors,
pending-formula panel opening and error/invalid-DXF variants, geometry,
appearance/all locales, larger regions and installed Linux/Windows VSIX. The
existing 2,000-cell filter budget remains. This does not complete Data/menu parity.

### Shared Data Tools controls and Text to Columns (2026-10-29)

The desktop and VS Code now mount the same `DataToolsRibbonGroup`. It contains
the original Text to Columns dropdown (comma, semicolon, space, tab), Flash Fill,
Remove Duplicates, Data Validation and Consolidate. Desktop has no independent
Text to Columns preview wizard. The plugin's existing dialogs remain adapters;
sharing this group does not accept full dialog or Data-tab parity.

Text to Columns fixes sparse-source blank handling and removes structural
`insert_cols` based on the used extent. Canonical cell writes grow used dimensions
without shifting formulas referring to previously blank columns. The operation
retains absolute row offsets and the existing native inclusive trailing-clear
column. A tracked promise includes both source reads and submission in Save,
rejects duplicate requests and permits retry after an atomic host rejection.
The selection and output retain their 2,000-cell limits.

Evidence:

- `/tmp/genoffice-data-tools-browser.log`: Save waits for held source reads and
  held submission, duplicate calls are rejected and one revision results; shared
  Remove Duplicates pending/retry regression passes. The first menu test failed
  due to an ambiguous Data-tab test selector (ribbon and sheet tab); corrected.
- `/tmp/genoffice-data-tools-menu-browser.log`: all four actual dropdown options
  split B4:B6, including the omitted blank B5, produce B/C outputs and leave
  `=D4` and the structural journal unchanged.
- `/tmp/genoffice-data-tools-retry-browser-final.log`: rejected host transaction
  leaves the original root unchanged, pending Save returns false and a retry
  succeeds. This exposed a sticky pending-error state, fixed by using the same
  recoverable submission mode as other semantic ribbon commands.
- `/tmp/genoffice-data-tools-native.log`: actual shared Remove Duplicates button
  and dialog pass native history/save/independent-reopen regression.
- `/tmp/genoffice-text-to-columns-native.log`: real VS Code 1.109.5 selects the
  shared dropdown option on B4:B6, performs exact Undo/Redo, saves the original
  formula reference and custom XML sentinel, and verifies C4/C6 and original
  A2 style/value through a new sidecar/session after reopen. Temporary extension
  host unresponsiveness was reported during startup and recovered; this is not
  performance acceptance.
- Editor, extension and desktop TypeScript checks pass; production extension
  build/browser boundary checks, focused lint/format and refreshed inventory
  checks pass (20 modules, 1,127 syntax evidence entries, 76 dynamic entries).

Still open: Flash Fill pending-work coverage, larger/rich-text/formula/error/
protection/merge variants, native trailing-clear semantics versus Excel wizard
expectations, full Data Validation/Consolidate desktop dialogs, Forecast/Outline/
Get Data/Pivot/Subtotal controls, all locale/appearance and installed Linux/Windows
acceptance. No Windows VSIX was generated or verified by this follow-up.

### Shared Forecast/Outline menus, small-view popups and Save results (2026-10-29)

Desktop and VS Code now reuse `ForecastRibbonGroup` and `OutlineRibbonGroup`.
The plugin's flat Goal Seek and outline buttons/summary checkboxes have been
replaced by the real What-If Analysis → Goal Seek and Group/Ungroup menus. Their
icons, descriptions, labels and canonical summary check marks come from the
same desktop components. Desktop still appends Subtotal, whose plugin workflow
is missing; the full Outline/Data ribbon is not accepted as complete.

A new failure-path test exposed a Save timing bug: after an already-enqueued
recoverable transaction was rejected, cell-editor flushing could consume its
pending promise and let its semantic-work slot clear before Save inspected it.
`flushEditor` now snapshots active semantic promises before awaiting the cell
editor, retains their results and checks subsequent work until settled. A failed
in-flight operation makes that Save return false while allowing the action to
retry. Existing Show Formulas, advanced filter, duplicate removal and Text to
Columns Save tests also pass.

Native small-window acceptance exposed another actual issue: the Group menu's
last summary option extended outside the webview and VS Code's status bar
intercepted clicks. Shared `MenuSelect` now measures its own viewport, selects
available space above/below, clamps horizontal bounds and maximum height, and
updates position on resize/scroll. Long menus scroll inside those bounds. The
native test retains its original small viewport; no forced click or enlarged
window is used to bypass this failure.

Evidence:

- `/tmp/genoffice-forecast-outline-browser.log`: nested row/column Group, Hide/
  Show Detail and Clear Outline dropdown commands retain canonical visibility.
- `/tmp/genoffice-forecast-outline-save-browser.log`: five existing Save-path
  regressions pass after the promise-result fix. The new summary test initially
  timed out on a wrong test label, Cancel instead of the actual Goal Seek Close
  button; corrected.
- `/tmp/genoffice-forecast-outline-browser-final.log`: four browser tests pass:
  591×365 outline gutters/depth and shared-menu commands, smart Sort & Filter,
  all four Text to Columns dropdown options, and rejected summary-setting Save/
  retained mark/retry/focus plus the Forecast → Goal Seek open/Close workflow.
- `/tmp/genoffice-forecast-outline-native.log`: the first native run identified
  real pointer interception by the VS Code status bar on the last Group option;
  embedded outline detail and Goal Seek tests passed in that run.
- `/tmp/genoffice-forecast-outline-native-final.log`: all three real VS Code
  1.109.5 tests pass after correction (44 seconds). They cover nested row/column
  grouping, both summary flags, exact Undo/Redo, source-cell preservation, saved
  outline XML and custom XML, independent new-session reopen and both restored
  menu marks; nested folding/depth/clear behavior; and cross-sheet Goal Seek with
  one revision/history/save/reopen. Startup reported temporary extension-host
  unresponsiveness and recovered (profile `/tmp/exthost-18a435.cpuprofile`); this
  does not establish startup/performance acceptance.
- `/tmp/genoffice-forecast-outline-build-final.log`: production extension build
  and browser boundary checks pass. Editor/desktop/extension TypeScript checks,
  focused lint/format and `git diff --check` pass. Refreshed source inventory:
  20 modules, 1,129 syntax evidence entries and 76 dynamic entries.

Still open: Subtotal dialog/transactions/reference transforms, Get Data/Pivot,
full Data Validation/Consolidate desktop dialogs, Flash Fill pending work,
complete keyboard/locale/appearance/lifecycle variants, large-workbook and
startup performance, and installed Linux/Windows verification. No Windows VSIX
was generated or verified by this follow-up.

### Shared Subtotal dialog and atomic structural/content workflow (2026-10-29)

The original desktop Subtotal button now appears in the shared Outline group
in VS Code. `SubtotalDialog` is the actual desktop component, retaining its
At each change in / Use function / Add subtotal to selectors, defaults and
Sum/Count/Average options. It now supports a controlled locale and async apply,
disables inputs/dismissal while pending, keeps rejected choices and translates
known planner errors. Headers and selection/root/revision are frozen by the
plugin adapter; an overtaken draft cannot create rows.

Both hosts use `buildSubtotalPlan`: adjacent equal labels form groups, later
repeated labels stay separate, blanks are retained, group and grand-total labels
keep desktop wording, and SUBTOTAL functions are 9/3/1 respectively (Count is
COUNTA). Totals use absolute column references and final row positions. The
desktop's generated totals now also select summary-below when a source requested
summary-above, matching their actual position and the canonical adapter.

A new standalone `subtotal` session operation materializes complete structural
source data and runs private insertion, content and outline stages. It returns
one final changeset with the original root as inverse; no intermediate root is
installed or added to history. Content stages are bounded, private faults and
cancellation do not mutate source state, and ordinary mixed structural/content
transactions remain rejected. Source formulas and typed/style payloads move
through the existing structural interpreter. The selected source is bounded to
2,000 cells and 200 groups; complete material retains the existing 200,000-cell
per-sheet budget. Existing source outlines and the generic unsupported
structural-feature cases remain guarded until their transforms are accepted.
Native session tools expose the operation. Standalone CLI, legacy desktop AI
execution and generic in-memory execution explicitly require a canonical session
for this compound operation; the normal desktop ribbon remains supported.

Evidence:

- `/tmp/genoffice-subtotal-plan-tests.log`: two pure planner tests cover blank/
  later repeated groups, absolute non-A/nonzero-row positions, group limit,
  invalid columns/labels and worksheet overflow.
- `/tmp/genoffice-subtotal-candidate-final.log`: 57 candidate/ribbon tests pass,
  including one-root inverse, typed/style preservation, cross-sheet formula
  movement exactly once, Sum/Count/Average numbers, incomplete/protected/source-
  outline guards, faults in private structural/content stages and cancellation.
- `/tmp/genoffice-subtotal-browser-final.log`: two browser tests exercise the
  real three-selector dialog, Average selection, host rejection with retained
  options, disabled pending Cancel/Escape, Save waiting, one accepted revision,
  generated formula/outline/gutter, no-write Cancel, localized invalid-column
  message and a draft overtaken by a concurrent edit.
- `/tmp/genoffice-subtotal-native.log`: initial real VS Code 1.109.5 one-revision
  history/save/reopen pass. Additional saved-outline checks initially asserted
  literal summaryBelow="1" and a metadata boolean; OOXML omits that true default,
  and the optional metadata field may likewise be absent. Those assertions were
  corrected to validate effective default semantics and the actual reopened UI.
- `/tmp/genoffice-subtotal-native-verified.log`: final real VS Code test passes
  (7 seconds). It verifies B3:C6 source offsets, two subtotal rows and grand total,
  exact native Undo/Redo, A8→A11 and Other!A1's shifted reference, original value/style
  preservation, saved formulas, outline levels 2/1 and custom XML, independent
  new-session reopen, the restored gutter and checked summary-below menu option.
- `/tmp/genoffice-subtotal-cli-guard-test.log`: session-only Subtotal is rejected
  explicitly before standalone file parsing/mutation.
- `/tmp/genoffice-subtotal-build-accepted-final.log`: production extension build,
  regenerated native tool schemas and browser boundary checks pass. The shared
  dialog initially used the broad UI barrel; switching to the existing chrome
  entry preserves the native-chat boundary. Editor/session/gateway/desktop/
  extension/CLI TypeScript checks, focused lint/format and diff checks pass.
  Gateway checking also exposed and corrected existing exact-optional-property
  type mismatches in shared filter predicate inputs; color-filter tests pass.
- Inventory includes SubtotalDialog: 21 source modules, 1,141 syntax evidence
  entries and 78 dynamic entries. Counts are not parity acceptance.

Still open: imported/existing outline, visibility/filter, merge/name/visual and
other guarded reference transforms; large sources/group-count performance;
computed totals and formula-cache freshness; every locale/appearance/lifecycle
variant; installed Linux/Windows behavior. Get Data/Pivot and other full-ribbon
and DOCX gaps remain. This follow-up does not establish complete Subtotal or
complete desktop-menu parity, and does not produce a Windows VSIX.

### Actual shared Consolidate dialog and sparse labels (2026-10-29)

The plugin now imports the actual desktop ConsolidateDialog and its chrome-only
Dropdown dependency. The previous plugin dialog replica is removed. Layout,
function selection, reference list and Add/Enter/Delete actions, typed-but-unadded
reference on OK, deduplication, target descriptions and left-label behavior come
from that same component. The existing desktop synchronous onCreate callback
continues to work; the shared component also accepts an async callback and a
controlled locale. A synchronous pending ref prevents duplicate applies before
a React render; pending fields, Cancel, Escape and backdrop close are guarded.
A rejected host result keeps every draft choice for retry.

The existing canonical Consolidate adapter remains authoritative. Its label-mode
reader incorrectly treated omitted sparse cells as incomplete data; such cells
now contribute blank labels and are skipped by the shared matrix algorithm.
Formula labels must be ready and have an authoritative cached result; error or
pending formulas are refused before output. Bounds/overlap/merge guards, frozen
root/revision checks and source/commit Save tracking remain in place.

Evidence:

- `/tmp/genoffice-shared-consolidate-browser.log`: three browser tests pass:
  original rejected label-mode retry and stale-draft behavior; all five shared
  function options with generated SUM/COUNTA/AVERAGE/MAX/MIN formulas, label-mode
  reset/removal of Max/Min, invalid reference, Add/Enter/dedup/Delete and no-write
  Cancel; sparse labels, pending source reads and submission, disabled controls/
  Escape, Save waiting and one final output revision.
- `/tmp/genoffice-shared-consolidate-formula-browser.log`: pending formula labels
  are refused, the entire original root stays unchanged and Cancel is restored.
- `/tmp/genoffice-shared-consolidate-native.log`: actual VS Code 1.109.5 shared
  Dropdown shows five options and selects Max then Sum; source references are
  added and accepted while the last typed source is still unadded. One final
  formula revision, exact native Undo/Redo, saved original custom XML and new
  formula, independent new-session reopen, unchanged original values on both
  sheets and the restored formula all pass (9 seconds). Startup reported
  temporary extension-host unresponsiveness and recovered; no performance
  acceptance is inferred.
- `/tmp/genoffice-shared-consolidate-build.log`: production extension build and
  browser boundary checks pass; no desktop AI UI dependency is imported.
  Editor/desktop/extension TypeScript, focused lint/format and diff checks pass.
- Inventory now includes ConsolidateDialog: 22 source modules, 1,155 syntax
  evidence entries and 78 dynamic entries. These counts are not completion.

Still open: complete formula/error/cache variants, imported rules and source
features, larger outputs/references, all locale/appearance/geometry/lifecycle
variants, performance, installed Linux/Windows, Data Validation's full desktop
pane and remaining Get Data/Pivot/other ribbons and DOCX workflows. No Windows
VSIX is generated or verified here; complete menu parity remains unproven.

### Shared Get Data group and From Text/CSV (2026-10-29)

Evidence (scratchpad logs of this session):

- Native VS Code 1.109.5 `Desktop Get Data UI`: Refresh All shows the desktop
  no-pivot message without a revision; a cancelled pick writes nothing; a BOM,
  semicolon-delimited CSV with a quoted delimiter imports at E3 as one revision
  with numeric typing and blank trailing fields; exact native Undo/Redo, saved
  custom XML sentinel, and independent new-session reopen pass.
- The existing native Data suite (Text to Columns, Subtotal, Remove Duplicates,
  Outline x2, Goal Seek, Custom Sort, Consolidate, Advanced Filter) passes with
  the new group mounted: 10 passing.
- Unit: `csv-import-model` (2), xlsx-gateway 331/331, workbook-session 250/250
  (one intermittent protected-view-save failure observed in one of four full
  runs), sheets-editor 42/42, vscode-office 61/61, desktop data/csv/ribbon 99/99.
  Desktop/editor/extension typechecks, lint, theme-color check and production
  extension build with boundary checks pass. Inventory regenerated: 22 modules,
  1,156 evidence entries, 78 dynamic entries.

Still open: 50,000-cell parity, Merge Workbooks, pivot refresh, locales/
appearances and installed Linux/Windows VSIX. Not complete Data-tab parity.

### Staged bulk transactions, 50,000-cell CSV and Merge Workbooks (2026-10-29)

- workbook-session `staged-candidate` (4): above-ceiling stages publish one
  changeset; add-sheet then content stages; >50,000 cells refused; engine commit
  is one revision/one history entry, idempotent retry and exact undo.
- workbook-session `merge-workbooks` (1, real sidecar): deduped "Data (2)"/
  "Other", values, cached formula value without formula, bold/fill/number
  format, C1:D1 merge, 21.5-character column width and 30pt row height.
- Native VS Code 1.109.5 `Desktop Get Data bulk UI` (2): an 8,400-cell CSV is
  one revision with exact Undo/Redo and saved/reopened values; Merge Workbooks
  through the host open dialog (simple-dialog mode) adds two deduped sheets in
  one revision with exact Undo/Redo, saved and reopened values/styles.
- Unit: xlsx-gateway 331, workbook-session 255, sheets-editor 42,
  vscode-office 61, desktop merge/csv 7; typechecks, lint and production build
  with boundary checks pass.

### Session PivotTable create (2026-10-29)

- workbook-session `pivot-save` (real sidecar): baked Region/Sum of Amount/Grand
  Total output, pending pivot blocks insert_rows and overlapping pivots, saved
  pivotTable/pivotCacheDefinition parts, custom XML preserved, reopened file
  reports the pivot range and value 32.5.
- Native VS Code 1.109.5 `Desktop PivotTable create UI`: Refresh no-pivot
  message; Insert → PivotTable opens the desktop dialog with C3:E7; an
  overlapping destination is refused in the dialog without a revision; G3
  creates one revision; exact Undo/Redo; saved pivot parts and sentinel; an
  independent reopen reports the native pivot range and baked values.
- Data/Insert native regression set: 18 passing; Insert Link fails identically
  at HEAD. Unit: gateway 331, session 256, editor 42, extension 61, desktop
  pivot/ribbon 194; lint/theme/typechecks/build and regenerated tool schemas.

### File PivotTable refresh (2026-10-29)

- workbook-session `pivot-refresh` (real sidecar): file pivot definitions load at
  open; an edited source value recomputes in place; a new member inside the
  source grows the output by one row; growth into occupied cells is refused;
  saved cache has refreshOnLoad, the pivot location is E2:F7 and a reopened
  file reports the grown range and grand total.
- Native VS Code 1.109.5 `Desktop PivotTable create UI` extended: after reopen,
  Data → Refresh shows "Refreshed 1 PivotTable(s)", one revision with recomputed
  North 23 and a new West row (G3:H8, grand total 40), exact Undo/Redo, saved
  refreshOnLoad and location, and a second independent reopen with the grown
  pivot range.
- Data/Insert native set: 19 passing. Unit: gateway 331, session 257, editor 42,
  extension 61, desktop pivot/ribbon 194. CLI: only the 4 failures present at HEAD.

### Edit PivotTable / Fields pane (2026-10-29)

- workbook-session `pivot-save`: session pivot relayout to Fruit shrinks E2:F6 to
  E2:F5, clears the old last row, and saves/reopens with the new range.
- workbook-session `pivot-refresh`: a grown file pivot relayouts to Fruit, refresh
  then refuses until save, and the saved definition location is E2:F5 on reopen.
- Native VS Code 1.109.5 `Desktop PivotTable Fields pane UI`: create through the
  dialog; the pane follows the cursor; checking Fruit nests rows (one revision);
  unchecking Region shrinks and clears leftovers; deferred mode waits for Update;
  exact Undo/Redo; save/reopen; the pane opens on the file pivot with Fruit
  checked, rejects an empty-rows layout in-pane without a revision, relayouts to
  Region and saves location G3:H7, confirmed by another reopen.
- Data/Insert/Pivot native set: 19 passing plus 1 Go To flake (Advanced Filter
  passes alone). Unit: gateway 331, session 257, editor 42, extension 61, desktop
  pivot/ribbon 194; CLI only the 4 failures present at HEAD.
