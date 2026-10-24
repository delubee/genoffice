# Existing merged cells — 0.1.2

## Fix

The VS Code grid previously discarded source merge geometry. The source feature
cache also marked an A1-only range response complete even though its merges were
limited to that range. A separate read-only `read_merge_ranges` sidecar request
now returns sheet-wide merge geometry with its indexing-complete flag, without
reading cell chunks. Completed geometry is cached for the immutable session.

Presentation reads include merge geometry and the current value of any visible
merge whose anchor lies outside the page. Ordinary model-facing cell reads keep
their existing shape and cell limits. The projection installs horizontal and
vertical merges again after commits, undo/redo, reload and sheet changes.

Manual commands coalesce repeated merged-selection input into its top-left
anchor. Conflicting pasted values are rejected. The shared candidate interpreter
and Python artifact imports reject covered-cell writes; original merge XML is
preserved on save. Copy/fill involving merges and coordinate-changing edits in
workbooks with merges remain unsupported and fail before commit. This change
does not add merge/unmerge commands.

## Verification

- Linux and Windows: three workspace TypeScript checks passed.
- Linux and Windows: session 51, editor 9, extension 31 unit tests passed.
- Linux and Windows: 273 Rust tests passed, including sheet-wide offscreen merges.
- Chromium/Univer: 13 browser tests passed. A real keyboard edit in the product
  React layout updates only the merged anchor and preserves the covered value.
  Horizontal/vertical merge geometry survives projection replacement.
- Linux / VS Code 1.109.5: actual Webview edits through manual and agent candidates,
  rejects a covered-cell write without state changes, saves, closes and reopens.
  All three original merge ranges and edited anchor values survive. The XLSX
  contains no newly introduced B1 value, and its custom XML sentinel is unchanged.
- Windows 11 / VS Code 1.109.5: the same merged-cell scenario and the appearance
  regression passed (2 tests). Final packaged installation evidence is recorded
  alongside the artifact audit.
- ESLint and `git diff --check` passed for the changed implementation and tests.

A browser screenshot is generated at
`packages/sheets-editor/test-results/browser/manual-intents-typing-into-93599-on-forwards-only-the-anchor/merged-heading.png`.
Real-model Chat invocation, macOS, and actual Remote SSH acceptance are outside
this bug fix's verification.

## Final Windows artifact

`genoffice-workbook-0.1.2-win32-x64.vsix`: 6,770,875 bytes.
SHA-256: `60dd8945da5bfa24e3967cb8d077b4ec9f042fb1b0837887aca616f192684dfb`.

The final static-CRT sidecar passed the actual Windows merged-cell test in
9,717 ms. Package audit confirms AMD64 PE and no VC++ redistributable imports.
A separate ordinary VS Code profile installed this exact VSIX and successfully
edited, saved and reopened a workbook, read it through the native tool, and
installed the optional product skill with Cargo removed from the runtime PATH.
The install proof hash matches the audited local artifact; see
`merged-cells-installed-win32-x64-proof.json` and
`merged-cells-package-audit.json`.
