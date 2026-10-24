# Upstream synchronization — 2026-10-24

The local fork was integrated with upstream `79ecbc97` while preserving the
VS Code workbook feature and its published history. The feature checkpoint is
`5f5dc5df` (0.1.4); `e98b8aec` merges it into `dev/genoffice`.
`main` is an exact upstream mirror. The final upstream merge is recorded on
`dev/genoffice`; `feature/vscode-office` remains at the feature checkpoint.
No branch was rebased, reset or force-pushed. This operation does not publish
local branches to `origin`.

## Integration decisions

Three files required conflict resolution:

- `apps/sheets/src/shared/desktop-api.ts`: retain the extracted protocol
  re-export, move upstream's redaction schemas into the shared protocol, and
  retain the new desktop redaction/skills APIs.
- `packages/xlsx-gateway/src/gateway/xlsx-gateway.ts`: retain the upstream
  redaction parameter at its positional slot; append formula-cache invalidation
  after it so older desktop/CLI callers keep their argument meanings.
- `packages/xlsx-gateway/src/gateway/xlsx-package-io.ts`: forward both save
  options in the corrected order.

A real-sidecar regression test saves redaction marks, including their previous
fill, while invalidating a stale formula cache in the same transaction.

The new Docs spreadsheet dependency is bundled explicitly in its standalone
main process, following the repository's workspace dependency rule. Built Docs
and Shell main bundles were checked for unbundled xlsx-gateway imports.
Two optional-property type mismatches in fork sorting/merge code were corrected
for the shared gateway's strict TypeScript settings.

## Release scope

The integration rebuilds local development outputs. The existing Windows
0.1.4 VSIX and its installed-package proof identify the pre-sync feature build
`5f5dc5df`; that artifact is not replaced or relabeled by this synchronization.
A release containing the upstream merge needs a fresh platform build and its
own installed-package verification.

## Verification

| Check | Result |
| --- | --- |
| TypeScript | VS Code session/editor/extension, shared gateway, Docs, Sheets and Shell passed |
| Rust XLSX | 275 tests passed; release sidecar rebuilt |
| Workbook session | 63 passed in the full run; the remaining history-storage test passed alone |
| Editor / extension units | 10 / 31 passed |
| Shared gateway | 302 passed |
| Sheets desktop | 3,665 passed in the full run, one skipped; the timed-out table-design file passed all 15 tests on isolated rerun |
| Docs desktop | 4,432 passed in the full run, four skipped; the Zotero TCP file passed all five tests with local networking enabled |
| Shell | All 1,107 tests passed with local networking enabled and two workers |
| Browser | 15 passed in the full run; the cold-start candidate test passed on isolated rerun |
| Real VS Code 1.109.5 / Linux | All six appearance, merge, dimensions, settings and Chinese-sort tests passed, including save/reopen |
| Builds | VS Code workspace, standalone Docs, and Shell builds passed |
| Static checks | Changed-file ESLint, theme-color gate and staged diff whitespace checks passed |

The initial broad test commands exposed infrastructure limits, which were
reproduced and resolved without changing application behavior or assertions:

- Concurrent compilation caused the history-storage case to exceed 20 seconds;
  it passed alone in 4.31 seconds with the original limit.
- The table-design case imports the desktop executor dynamically. Its five-second
  limit expired during that import; isolated verification with a 20-second CLI
  allowance passed the complete file. The test source and default limit were
  left unchanged.
- Updating the lockfile caused Vite's first dependency optimization to exceed the
  browser case's 30-second limit. The same case passed in 3.1 seconds once the
  optimizer cache was populated; the test limit was left unchanged.
- The default execution sandbox forbids TCP/Unix-socket listeners and some child
  launches. Zotero and Shell networking/process tests were rerun with those
  capabilities enabled. The final Shell full-suite command passed.

The passing cross-save regression is
`packages/workbook-session/tests/upstream-save-compat.test.ts`.
Detailed local command logs use `/tmp/genoffice-upstream-*.log`.
