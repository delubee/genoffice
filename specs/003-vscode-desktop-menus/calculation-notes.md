# Calculation integration evidence and remaining design

The goal remains full desktop-menu fidelity. These notes do not establish a
working Calculation group or reduce formula/computation requirements.

## Current evidence (2026-10-27)

- Desktop `calc-options.ts` vetoes automatic calculation-start mutations in
  manual mode; forced Calculate Now and scoped Calculate Sheet requests remain
  permitted. Desktop App handles Automatic/Manual, F9 and Shift+F9 commands.
- Plugin `WorkbookRuntime.recalculate()` currently invokes only its display
  engine. Session source/formula records can still report pending results, and
  package-writer deliberately omits unverified caches. A visible mode switch
  alone would leave a different result/history/save behavior and is insufficient.
- Existing `XlsxSidecarClient.recalcCells` can evaluate a staged file through
  IronCalc, but uses a different function engine and its wire result currently
  lacks a general exact scalar/type field. It cannot silently become proof of
  desktop-equivalent fresh results for every formula.
- A real headless Node experiment successfully registered UniverSheetsPlugin,
  UniverFormulaEnginePlugin and UniverSheetsFormulaPlugin, created a sparse
  workbook with A1=2/B1=`=A1*2`, observed the calculation-result mutation and read
  B1 as `{f:'=A1*2',v:4,t:2}`. The engine disposed and process exited normally.
  CommonJS loading works; direct unbundled Node ESM loading fails on the
  transitive opentype.js named-export boundary. Existing host builds use
  esbuild/CommonJS; packaging still needs explicit verification.

A second real Node experiment returned numeric 4 and cross-sheet 9, boolean
`{v:1,t:3}`, text and an error string, plus a SEQUENCE anchor. Two important
limitations were observed: the raw worksheet result represents both text and
`#DIV/0!` with `t:1`, so the scalar matrix alone does not establish error provenance;
and a spilled child was absent from getCellRaw and the ordinary result unitData.
These observations motivated the rich-runtime adapter below. Its public runtime
observer now retains error provenance and complete matrices; canonical document
cache/source integration is still required. The observations are not a reason to
silently drop dynamic arrays or narrow formula coverage.

## Implemented normalized-snapshot adapter

`packages/workbook-session/src/calculation/univer.ts` now calculates an already
hydrated authoritative Univer snapshot without a browser/DOM. It uses the same
core/sheets/formula plugins, disables initialization calculation, and makes one
explicit workbook or desktop-style sheet calculation request. Its observer wraps
the public `IFormulaRuntimeService.setRuntimeData`, forwards the original runtime
write, and collects rich values from that evaluation. It does not re-evaluate a
cell to guess whether a string is an error. The dependency manager's getAllTree
API was tested and found unimplemented in this installed version; no private
AST/cache inspection or second evaluation is used as a workaround.

The result preserves caller workbook identity, sheet/row/column/formula, explicit
requested scope, exact scalar types, full array values, computed number-format
hints, hyperlink and IMAGE metadata. Engine single-empty-reference and spilled
blank semantics differ and are preserved through the runtime's own getFirstCell
and iterator APIs. Blocked spills publish an error-only result; accepted arrays
whose first text value happens to spell `#SPILL!` remain text arrays. Authored
input is cloned and never changed. Returned formatting is a hint, not permission
to replace authored document styles or execute privileged URL/image actions.

Each invocation uses a private engine unit UUID. Tests reproduced wrong SUMIF
results for concurrent/reused caller IDs before this isolation, then passed with
private IDs. Results have stable sheet/row/column order. Cancellation, including
an abort at START followed by the engine's internal reset, and output-budget
failures return no partial successful result. Engine hooks and subscriptions are
restored/disposed. The output-cell budget does not make an individual expensive
function preemptible or enforce a heap bound; worker/resource isolation remains
required before wiring production document commands.

The test builds and loads a real Node22/CommonJS artifact through esbuild, matching
the host bundling boundary. It covers scalars, error-shaped text, arrays and their
error elements, references/blank references, spill rejection, same-ID concurrent
jobs, cancellation/reuse, formatting/link/image metadata and Calculate Sheet.
The scoped request uses the actual desktop dirtyRanges/dirtyNameMap geometry;
an off-sheet dependent updates while an unrelated cached formula remains absent
from the result. The adapter is not yet called from the ribbon or a canonical
calculation/cache controller. It does not prove full-source, scoped-name, table,
external-link, date-system/time-zone, rich-resource, worker or saved-file parity.

## Required implementation and verification

1. Share the actual desktop Calculation group, options/default echoes and
   keyboard commands. Do not implement a decorative/manual-only UI switch.
2. Establish a trusted host calculation service using the desktop engine and
   exact typed results, including errors/text/booleans, named scopes, cross-sheet
   references and unsupported/function coverage. Hydrate authoritative source
   precedents and current overlays, not only the visible grid. Match desktop
   Calculate Sheet dependency semantics rather than just filtering result cells.
3. Apply results through the canonical transaction/cache model, with consistent
   read/search/watch/chart/preview/save behavior, source-formula cache support and
   explicit pending/stale status. Preserve authored formulas and unrelated OOXML.
4. Avoid independent browser/host calculation of volatile functions when comparing
   candidate roots: NOW/RAND and other nondeterministic values cannot be validated
   by expecting two separate evaluations to match byte-for-byte. The host must
   establish and validate the authoritative result once.
5. Define automatic/manual state, projection rebuilding, file calcPr persistence,
   native Undo/Redo and Save waiting together. A manual-mode projection must not
   accidentally run forced calculation from ordinary loadPage/install calls or
   discard previously displayed results merely because the projection rebuilt.
6. Verify actual menu actions, source/off-page/cross-sheet computation, typing and
   AI-origin changes, fresh/pending reads, atomic failures/stale work, native
   history, saved caches/calcPr and distinct-session reopen. Keep the full menu,
   large-file, locale/appearance and installed-platform requirements open.

Next integration must bind complete source/context data and result artifacts to a
frozen authored root. It must preserve hidden-name privacy, array ownership and
explicit freshness, and keep derived cached values distinct from authored cells.
A sheet-scoped result is not a statement that every formula in the workbook is
fresh. Preserve exact empty-string/boolean/error caches and spill OOXML on
save/reopen; review calcPr, 1900/1904 dates and UI-versus-host time-zone semantics.
Do not enable the new Calculation menu on a display-only or viewport-only path.

## Implemented complete stored-cell loader

`calculation/source.ts` now loads a frozen logical root from the fully indexed
immutable source, rather than from the projection or only static precedents.
The native internal `stored_cells` enumeration includes formatted blanks and
empty strings, retains source array anchors/ranges, and follows completed index
chunks even when a producer understates worksheet dimensions. Ordinary user Find
modes keep their existing text matching and do not expose array metadata.

The loader validates source scope, grid coordinates, ascending unique cells and
advancing cursors; it retries a partially indexed page at the same position.
It maps cells through insertion/deletion logs, excludes deleted originals, merges
explicit clears/styles/formulas and new sheets, and reads sparse bounded bands
through the existing typed/formula/name-rewriting current-range reader. It raises
only its private source bounds when indexed cells exceed metadata dimensions.
An enumerated source cell missing from a range response causes failure rather
than becoming a fabricated blank. Root ownership/cancellation is checked before
and after asynchronous reads, including style metadata, with no successful
partial result. Cell/byte budgets are explicit operation resource budgets, not a
new claimed document-size limit or a preemptive process memory bound.

The host-private result separates logical features from source-space geometry,
retains source array ownership even for overwritten/deleted anchors, and keeps
current definition/reference journals and structural context. This is logical
data plus resource context, not an engine-ready resource binding or a verified
calculation result. Hidden/scoped definitions now have the private package context described below;
source names/arrays/tables/external resources must still be bound and transformed
before this data may be calculated as document authority. Column/row defaults,
dates/time zones and full function semantics remain part of that work.

Real-sidecar tests exercise an understated-dimension XLSX with data beyond row
100000, styled blanks, exact empty strings/errors, formula caches, overridden
array anchors and another sheet. A second fixture crosses the 10000-result native
page boundary with 10005 dense cells plus a far formatted blank, then loads the
current inserted-row geometry and explicit overrides. Unit tests also cover
stale/cancelled reads, private class receivers, incomplete/foreign/duplicate
pages, missing records and operation budgets. See verification.md for counts.

## Host-private workbook calculation context

The native read_calculation_context command is session-scoped and reads only
that session's immutable source path. WorkbookSource exposes it as an optional
host-only capability; calculation loading requires it and rejects absence,
stale/cancelled ownership, a foreign source session or exceeded context budgets.
SidecarWorkbookSource validates the response and every local-name scope against
its owned source metadata. Closing a source prevents subsequent context reads,
and the client includes these pure reads in its queued-read cancellation sweep.

The result retains every original definition, including hidden, empty, built-in
and duplicate-spelling global/local names. Scope uses stable source sheet IDs,
not the current workbook order. Name text supports numeric/general XML entities,
CDATA and prefixed SpreadsheetML; nested markup, invalid scopes and foreign
relationship namespaces fail. XML namespace and parent-path checks keep extension
lookalikes out of bindings. All unqualified calcPr attributes are retained, as
are ordered external relationship IDs, without fetching external dependencies.
The calculation options are raw source settings, not already implemented manual,
iteration or precision behavior. DTD contexts are rejected.

Hidden definition bodies remain absent from ordinary open metadata, canonical
state and the existing public renderer/AI read routes. Only the host-private
loader result receives the private context. Tests verify this with a real native
session while sheets are reordered and renamed; source definition text and
scope remain separate from the logical transformation journals. Nothing from
this result is written into public metadata or authored workbook cells.

A local inspection of the installed Univer DefinedNamesService implementation
found that getValueByName caches by unit ID and lower-cased name, without a
worksheet-scope key. The types expose localSheetId, but that alone does not prove
correct resolution for same-spelling local/global names. Do not simply register
all definitions and certify fresh document results: binding and a real evaluation
regression must resolve this case, together with name-reference replay, array/
table/external resources and date/default context. The actual Calculation group
and canonical cache/worker/save integration remain incomplete.

## Normalized calculation name binding and shared formula provenance

The normalized Univer adapter now accepts logical definedNames. Its private
binding pass uses the real engine tokenizer and creates isolated hidden aliases
for each resolved binding/evaluation sheet. Same-spelling local/global names no
longer share the engine's spelling-only cache entry. Qualified sheet names,
quoted apostrophes, Unicode/case-insensitive lookup, current-workbook-qualified
global references and local fallback are covered by real calculations. Local
scope takes precedence for an unqualified name, matching Microsoft's
[Names in formulas](https://support.microsoft.com/en-us/excel/names-in-formulas)
reference. This is not permission to reinterpret external workbook qualifiers.

Binding changes only the private computation snapshot. Authored formulas and
caller definitions are retained separately, and progress callbacks cannot mutate
the bound resources through the caller's objects. Nested definitions retain
expression grouping. LET declarations/initializers/sequential scope and LAMBDA
parameters are handled separately from document names; ordinary built-in calls,
string literals and table/external tokens retain their lexical role. Named
LAMBDA calls use the engine's actual inline-call syntax, which rejects the extra
grouping used for ordinary named expressions. A negative name also requires an
explicit grouped operand: the tokenizer exposes -Rate as one token, and a raw
negative private alias would otherwise yield NAME instead of the negative value.

Address functions resolve dynamic name strings through their public
getDefinedName method at execution time using the runtime's sheet, rather than
looking up a mutable process-global spelling map. The adapter restores these
per-job methods on cleanup. Actual INDIRECT tests cover direct named ranges,
case-insensitive strings and qualified off-sheet names; this is not evidence of
complete arbitrary named-expression/recursive/iteration support for INDIRECT.
No numerical probe or second evaluation is used to infer a name result.

Shared masters/followers are expanded with the engine's public reference-offset
operation before binding and before capturing authored provenance. Missing or
ambiguous masters are rejected. Returned formulas for followers are logical
original formulas, with correct relative offsets and no private alias leakage;
source SI/master records stay unchanged. Circular static name bindings fail
explicitly, not with a successful partial result. This is an implementation
boundary, not completion of legitimate recursive LAMBDA or iterative calculation
requirements. The private alias count is bounded as an operation resource check,
not as a claimed workbook-size limit or worker heap/CPU guarantee.

The host source loader now composes private source definitions with modeled
Name Manager edits. Original local scopes map through stable source sheet IDs to
live logical sheet IDs, so sheet reorder does not change scope and deleted local
definitions are omitted. Source-only bodies receive the logical coordinate and
name-reference transforms; action-time journal bodies are appended without
replaying earlier renames over newly authored text. Hidden, empty, built-in and
explicitly preserved source definitions remain available privately. Hidden
bodies remain outside canonical state. This is verified with a real native
source session and a synthetic save/reopen calculation fixture.

Actual Excel 16.0 COM inspection on Windows showed that a global definition
created while another sheet is active stores an unqualified nested name as a
workbook-level binding (`=Rate*2`); a definition explicitly created against a
local name stores the sheet-qualified reference (`=Other!Rate*2`). The normalized
binder now uses this distinction for unqualified names inside global definitions
while retaining explicit local qualifiers. This is a concrete Excel observation,
not a claim that all nested-name, legacy-document or dynamic-name behavior is
covered. A later attempt to open the saved synthetic fixture could not reach the
Windows host (`127.0.0.1:2222 Connection refused`), so independent Excel
acceptance of that saved file is still pending.

Coordinate-changing operations remain guarded when name-reference bindings are
present; the pure source-definition transform test does not lift that guard.
Legacy recovered states without stable `sourceSheetIds` retain a compatibility
fallback and need dedicated migration/ambiguity handling. Prefixed writer XML
with CDATA/comments also needs broader preservation coverage. Tables, source
array ownership, external resources, date/default/time-zone semantics,
recursive/iterative names and dynamic expression coverage remain required.
Canonical cache/worker/menu/mode/history/save/reopen integration and the full
desktop-menu goal remain open.

## Structured-table calculation binding work in progress (2026-10-28)

The private source result now carries each source table's logical sheet ID,
range, name and column-title map. The normalized engine adapter registers these
tables through the public `ISuperTableService`, which is the formula engine's
structured-reference lookup. It validates workbook-grid bounds, unique
non-empty column names and workbook-wide table-name uniqueness. The installed
engine's `ISuperTable` shape has no totals-row count and assumes a single header
row; source tables with totals rows or a different header count fail closed.
Structured `#Totals` references also fail closed because the engine maps that
section to the last table row regardless of workbook metadata.

This is code-level binding only. It has not yet been exercised with actual
structured-reference formulas or a native XLSX table fixture, and calculation
results are not consumed by the plugin editor. Table calculated-column formulas,
table edits/structural transforms, external links, arrays, date/default context,
worker bounds and authoritative cache/menu integration remain open.

The new source snapshot builder connects loaded cells, logical formula strings,
inline number-format patterns, logical names and table resources to one engine
request. It preserves formula-looking text as forced strings. It rejects 1904
workbooks, legacy source array ranges, literal source error cells and formulas
with external-workbook references until their engine resources are represented.
`calculateSourceWorkbook` now composes this snapshot with the normalized engine
adapter, but no editor command calls it yet. Row/column defaults, full style
semantics, exact 1904 dates and authoritative result publication remain open;
this adapter path is not yet document-authoritative.

The plugin Formulas tab now renders the shared desktop Calculation Options,
Calculate Now and Calculate Sheet controls. The editor runtime reuses the
desktop automatic/manual start-mutation veto and active-sheet dirty-range/name
request; F9 and Shift+F9 route to those same operations, excluding active cell
editing, inputs and dialogs. Mode state belongs to the current runtime and is
not persisted into the source `calcPr`. The current controls still operate on
Univer's editor projection, not `calculateSourceWorkbook`; they do not publish
canonical fresh caches or guarantee Save/reopen equivalence. The visual/runtime
control parity is now present, while the document calculation contract remains
open and must not be considered complete.
