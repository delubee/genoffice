# Implementation status — 2026-10-24

This is a **preview, not a release candidate**. All six stories have implemented code and local deterministic evidence. Real model/Chat acceptance, macOS 14/SSH environments, and the exact performance reference environment remain unverified. Windows 11 x64 build and installation are now locally verified; Windows current-stable/human/crash acceptance remains pending. `specs/001-vscode-office-editor/tasks.md` retains unchecked release gates.

## Implemented and locally exercised

- Shared browser editor + Node session authority, independent of the Electron application. XLSX custom editor, values/formulas, common formats, typed copy/fill, worksheets/rows/columns and basic session charts.
- Immutable original package, typed sparse overlay, host-verified atomic changesets, versioned bindings, alias locks, disk-conflict handling, idempotent receipts and bounded durable undo/redo. Save does not clear history. Unsupported/protected operations fail closed; original native objects are preserved.
- Real Linux VS Code native save/undo/redo, hot exit and SIGKILL recovery. A first-ever crashed window may need **GenOffice: Recover Unsaved Workbook**. Recovery creates a new session without the previous undo timeline.
- Eight stable native LM tools, current unsaved reads, explicit selection binding, aggregates and ordinary atomic writes. Native invocation harness checks target/focus/stale-revision behavior; it is not a model conversation.
- Typed streamed workspace export, immutable validated private result staging, 30-minute import plans, at most 1,000 cells per undoable batch, strict cursor/revision chains and honest cancellation progress.
- Actual 100k input / Python standard-library summary / new Summary and chart / manual continuation / save/reopen. Separate full 100k output correctness test completes 100 batches in about 11 minutes and verifies saved/reopened typed hashes, cancellation and all receipts.
- Optional workspace skill installation/status with explicit updates, user-content detection, differences and backup; visible standard-library Python example. No model credentials, alternate chat or process runner is introduced.
- Linux x64 platform VSIX with packaged sidecar and legal notices; architecture/permission/resource checks and SHA-256 artifact. Three-platform CI is configured; exact Windows 11/macOS 14 arm64 runners must be provisioned.
- Actual local UI samples: opening 20/20 <=10 s and 1k-cell edit 20/20 <=2 s. Host has about 16 GiB RAM, not the required 8 GiB reference environment. See `test-results/performance.md`; reference acceptance is pending.

## Boundaries and outstanding acceptance

Original native charts and advanced workbook objects are not generally editable; the chart pane manages session-created basic charts. Formula caches can remain pending. Chinese and English editor labels are provided; other locale shards use the English fallback. Stable registration APIs cannot report whether a particular Chat conversation has enabled a tool.

Actual macOS/SSH installation, logged-in agent read/edit/Python conversations and tool-confirmation rejection are unverified. Local tests do not replace those environments. No PR or Marketplace publication has been created.

## Entry points

See [README](README.md), [troubleshooting](docs/troubleshooting.md), and `test-results/` for commands, limits, proofs and remaining gates. Development: `npm run vscode:build`, `npm run vscode:typecheck`, `npm run vscode:test`. Platform package: `npm run native:build -w apps/vscode-office`, `npm run notices`, `npm run package -w apps/vscode-office`. Tests include browser E2E, native integration, installed-VSIX smoke, recovery, capacity and performance as separate explicit commands.

## Windows update — 2026-10-23

Windows x64 VSIX was built on Windows 11 in `E:\github\genoffice` and installed in a clean VS Code 1.109.5 profile. Windows unit tests (46 session / 7 editor / 31 extension), 272 Rust tests, eight native integration tests, and final installed-package smoke passed. Windows recovery flush/drive-case/retired-slot and test portability fixes are included in this update. See [Windows build evidence](test-results/windows-build.md) and [installed package proof](test-results/package-win32-x64-proof.json). Current-stable Windows, human input/crash and real-model acceptance are separate pending gates.

## Preview 0.1.2 — appearance and merged cells

Excel-style light chrome is now the default, with a persisted option to follow
VS Code. Existing horizontal/vertical merges render correctly; manual input
writes the anchor, and agent/Python writes to covered cells are rejected.
Creating/unmerging cells and copying/filling merged ranges remain unsupported.

Linux and Windows checks passed: three workspace typechecks, 51 session / 9
editor / 31 extension unit tests, and 273 Rust tests. Chromium passed 13 browser
tests including actual keyboard input into a merged selection. Actual VS Code
1.109.5 on Windows passed merged-cell save/reopen and appearance tests; the final
static-CRT 0.1.2 VSIX also passed independent normal-profile installation.
See [appearance evidence](test-results/appearance.md),
[merged-cell evidence](test-results/merged-cells.md), and
[final installed package proof](test-results/merged-cells-installed-win32-x64-proof.json).
Remaining real-model, macOS, Remote SSH and broader release gates are unchanged.

## Preview 0.1.3 — ribbon and worksheet sizing

Seven ribbon tabs now expose common Home, Insert, Page Layout, Formulas, Data,
Review and View operations. Header dragging and numeric row/column sizes use
shared transactions, native undo/redo and XLSX persistence. Original source
sizes are projected, including blank rows with custom height. Page settings,
frozen panes, value filters and text notes are persisted through existing
shared OOXML writers. The product skill and native tool schemas include these
operations. The native VS Code chat workflow is unchanged.

This is an initial migration of common ribbon controls, not full desktop
feature parity. See README's ribbon matrix for the available controls and
limits. Advanced dialogs, original objects, print preview, and transformations
that cannot preserve reference-bearing features remain outside this increment.

0.1.3 verification: Linux/Windows unit suites 59 session / 10 editor / 31
extension; shared gateway 252; desktop compatibility 70 and desktop typecheck;
Chromium 15 browser tests. Windows actual extension-host tests and final
normal-profile VSIX installation passed. See
[ribbon and dimensions evidence](test-results/ribbon-and-dimensions.md) and
[0.1.3 installed proof](test-results/ribbon-installed-win32-x64-proof.json).

## Preview 0.1.4 — merge editing

Home includes merge/unmerge commands through the shared manual/agent transaction
path. Merging refuses nonempty covered cells and unsupported feature overlaps;
undo/redo, original-based saving and reopening preserve geometry and contents.
Windows verification passed: 104 workspace unit tests, six real VS Code tests,
and actual VSIX installation with editing, saving and native tool reads.
See [merge editing evidence](test-results/merge-editing.md) and
[0.1.4 installed proof](test-results/merge-installed-win32-x64-proof.json).

## DOCX 0.2.0 verification

The same plugin now includes a lazy DOCX host bundle and separate Word webview.
Implemented: manual formatting, headings and genuine numbering definitions,
rectangular tables, original-based saves, native tool editing, scoped selection
bindings, hash/version-bound Python exchange, private recovery and shared native
history. Unsupported/private blocks are protected and omitted from tool reads.

Final verification: all five workspace type checks, 133 unit tests and 17 affected
desktop style tests passed. Browser controls, actual VS Code native tools/Python,
backup restoration and independent XLSX/DOCX sessions passed. The final Windows
fixture (1,113,986 bytes, 200 paragraphs, 10 tables) opened in 1,977 ms; a
100-character native tool edit projected in 132 ms.

Both Linux and Windows 0.2.0 VSIX files passed actual isolated production
installation tests, including real Word webview input, save/reopen, undo/redo,
native reads and shipped skill installation. See
[feature verification](../../specs/002-vscode-docx/verification.md) and
`test-results/docx-*-proof.json` for exact hashes/evidence.

Advanced/private content and complex tables remain protected. Native history
survives saves; recovery starts a fresh session/history. Real-model chat/skill
discovery, macOS, actual Remote SSH and forced-crash/human acceptance remain
unverified. This remains a preview.
