# DOCX 0.2.0 verification — 2026-10-24

The existing extension ID remains `genoffice.genoffice-workbook`. XLSX and DOCX
ship in one VSIX with separate editors, lazy host bundles and independent state.
Build verification predates the feature commit. No Marketplace publication was performed.

## Gates

- Windows final source: all five workspace type checks and **133 unit tests**
  passed (70 workbook-session, 10 sheets-editor, 19 docx-session, 34 extension).
- Linux: the same package checks, 17 affected desktop style tests, and built-page
  manual input, complete font/size entry, headings, table rows/columns, shared
  history and theme preservation passed.
- Actual VS Code 1.109.5: native DOCX tools, Python export/preview/import,
  save/undo/redo, backup restoration with fresh identity, and independent
  XLSX/DOCX documents passed. Windows also passed the junction/UNKNOWN path
  regression: 3 final native tests passed.
- Actual installed production packages on Linux and Windows: edited XLSX value
  731 survives reopen; manual Word input survives save/reopen; Word native
  undo/redo, registered native reads and skill installation passed. The fixture
  controller edits the real isolated webview through Chromium; production tool
  confirmation settings are preserved.
- Dependency graph checks keep DOCX outside the eager extension bundle and
  disallow Node/Electron/AI runtimes in the browser. Only the audited pure
  agent-core redaction subpaths are allowed.
- Formatting and `git diff --check` passed.
- Commit-time lint: changed JavaScript/TypeScript files pass. Repository-wide
  lint also reports six pre-existing errors in three Sheets test fixtures.

## Windows performance

VS Code 1.109.5, win32-x64, Node 22.17.0, isolated build on `windows-work`:

- DOCX fixture: 1,113,986 bytes, 200 paragraphs, 10 tables.
- Cold open through mounted editor readiness: **1977 ms**.
- 100-character native tool edit through UI projection: **132 ms**.

These are measured fixtures, not a promise of exact Word pagination or arbitrary
large-document performance.

## Artifacts

- Windows: `apps/vscode-office/artifacts/genoffice-workbook-0.2.0-win32-x64.vsix`
  (9,151,037 bytes).
  SHA-256: `3b1e13dea539a72c69530270bd6e6b54d5e4cafe545bdf78a798aa31a4395268`.
- Linux: `apps/vscode-office/artifacts/genoffice-workbook-0.2.0-linux-x64.vsix`.
  SHA-256: `b0fc21fd0f5d215e5f2835c76b6526b24a35828406d14d2af8930e2fbb450ca6`.
- The Windows binary is AMD64 with static CRT; no VCRUNTIME/MSVCP DLL dependency
  or runtime Rust/Cargo installation is required.
- Windows delivery copy:
  `E:\github\genoffice\apps\vscode-office\artifacts\genoffice-workbook-0.2.0-win32-x64.vsix`.
- Build isolation:
  `E:\github\genoffice\.genoffice\work\j015-20261024`.
  The original Windows project sources were preserved.
- Source baseline: `c726850b`; the 80-file verified source delta has archive
  SHA-256 `b4620493928cf201e64784f653d7bc00ee23ec6ad96ffb9394216f405106cca4`. Post-verification documentation updates are
  outside that snapshot. Subsequent scoped lint annotations change comments
  only; runtime behavior matches the verified artifacts.

Evidence lives in `apps/vscode-office/test-results/`:
`docx-installed-win32-x64-proof.json`, `docx-package-win32-x64-audit.json`,
`docx-win32-x64-performance.json`, and `package-linux-x64-proof.json`.

## Implementation boundaries

Native history remains intact across saves. Recovery restores the unsaved OOXML
snapshot and original disk baseline into a fresh session/history; the native
process-local undo stack is not serialized across a restart. Existing disk
changes still block overwrite after recovery.

The XLSX type program retains its existing indexed-access checks. The separate
DOCX program uses the upstream Docs strict TypeScript settings so both are
checked without changing upstream source-wide compiler requirements.

Advanced/private objects and merged/complex tables are protected. Header/footer
editing, tracked-change/comment editing and precise Word pagination are deferred.
Real-model chat/automatic skill discovery, macOS, actual VS Code Remote SSH
acceptance and forced-process-crash/human acceptance remain unverified.
