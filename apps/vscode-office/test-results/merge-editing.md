# Merge editing validation — 0.1.4 source

Date: 2026-10-23. Development branch: `feature/vscode-office`.

Home now exposes Merge cells and Unmerge cells. Both operations use the shared
session candidate interpreter, native undo/redo and original-based XLSX writer.
The AI tool schema and bundled workbook skill include the operations.

Merging refuses nonempty covered cells, including zero, false, empty strings and
formulas with blank results. Values and styles are never automatically discarded.
Partial merge overlaps and table, pivot or filter intersections fail before
commit. Unmerge accepts any intersecting cell and preserves cell contents.
The current geometry also constrains subsequent operations in the same batch.
Coordinate-changing operations after session merge edits remain restricted.

## Verification

- Three workspace typechecks passed; extension host/webview build and generated
  tool-schema consistency check passed.
- Workspace unit regression: 62 session, 10 editor and 31 extension tests passed.
  A further targeted four-test merge suite verifies content preservation,
  overlap/conflict rejection, inverse roots and same-batch merge/write guards.
- Shared XLSX gateway regression: 252 tests passed.
- Browser regression: 15 of 16 passed in the full run. The remaining seven-tab
  keyboard test was interrupted by Vite reloads from a simultaneous production
  build; it passed when rerun alone after the build finished. The new merge
  button test checks real projected geometry and refusal to hide existing data.
- Real VS Code 1.109.5 on Linux: the merge integration test passed. It unmerges
  an original range, writes a formerly covered cell, creates a new merge,
  checks undo/redo, inspects saved OOXML, preserves an unrelated custom XML part,
  and verifies geometry and values after closing and reopening the workbook.
- Changed-file ESLint, theme-color gate and `git diff --check` passed.

## Windows 0.1.4 verification — 2026-10-24

The restored `windows-work` connection was used to build in
`E:\github\genoffice`. Incoming source files were backed up before applying
this source update. Windows typechecks, 63 session tests, 10 editor tests and
31 extension tests passed; the extension host/webview build passed.

Six real VS Code 1.109.5 integration tests passed: merge editing, existing
merged-cell preservation, row/column dimensions, page/view/notes/filter settings,
Chinese sorting, and Excel/VS Code appearance switching. All relevant saves and
reopen checks succeeded.

The actual Windows x64 VSIX was installed into an isolated ordinary VS Code
profile. Editing, saving/reopening, native workbook tool reads and bundled skill
installation passed with Cargo/Rust paths removed from the runtime PATH.
The installed-package proof matches the downloaded artifact SHA-256.

- Artifact: `artifacts/genoffice-workbook-0.1.4-win32-x64.vsix`
- Bytes: 6,788,106
- SHA-256: `20874cea23d520246dd829785935e1e76eb6413e32aaffba3d29d82c16f58b4d`
- Package audit: [merge-package-audit.json](merge-package-audit.json)
- Installed proof: [merge-installed-win32-x64-proof.json](merge-installed-win32-x64-proof.json)
- Native component: AMD64 PE with static CRT; no VCRUNTIME/MSVCP DLL dependency.
- Remote logs: `.genoffice/work/merge-20261024/{typecheck,unit,build,native,package,installed}.log`

Conditional formatting, data-validation controls and other advanced desktop
features remain outside this increment. No real-model Chat test was performed.
