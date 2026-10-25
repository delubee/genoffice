# Verification: shared desktop ribbons and DOCX dialogs

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
