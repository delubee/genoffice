# US1 manual editing evidence — 2026-10-22

Scope: the basic manual editing story on this Linux x64 host, VS Code 1.109.5 and Playwright Chromium 151.0.7922.34. No model account, chat conversation or Python interpreter was used. This is not release or cross-platform acceptance.

## Verified

- Real product Webview opening and first-page installation; manual-origin transactions use the browser candidate and independent host replay, update the visible projection and native dirty state, and save successfully.
- Save As contains unsaved edits, preserves source bytes and its saved marker, and opens under a new session. Reopening the same URI reuses its document. Closing releases the session. A physically read-only workbook rejects edits and normal save without moving its root.
- Visible Univer commands leave the previous value/style until host acknowledgment. IME composition blocks flush. Direct sheet mutations are rejected. Basic formula `=B2*2` displays 240, then 400 after its input changes.
- Toolbar font, bold/italic, colors, fill, number format, alignment, wrap and borders use format intents. Copy/fill copy typed values and relative formulas and retain destination formatting. An independent host test fills 2,001 cells and rejects overlapping copies.
- Sheet add/rename/delete and row/column insert/delete preserve logical identities. Real staged saves replay row insertion, rename and add/delete tombstones from the original package twice, preserving shifted values, styles and cross-sheet references without applying structural operations twice.
- Column/bar/line/pie session charts serialize and reopen with current unsaved source data, and do not duplicate on a second save. The toolbar creates, edits and deletes charts; chart caches follow supported cell edits. Charts are currently displayed in a side pane; their stored worksheet anchors are retained for XLSX output.
- Damaged files are rejected with private snapshots cleaned up. Injected ENOSPC and cancellation leave original/target bytes intact and clean incomplete stages. A destination changed before publish is preserved together with the staged package.
- Pending formula results are not serialized as valid caches. Nonidentity session saves invalidate unverified original formula caches too, including untouched dependent formulas on another worksheet; explicitly supplied ready results can then be applied. Reopen reports missing caches as pending. Browser calculation results are projection data and do not create a second history or claim authoritative ready caches.
- VS Code light/dark changes update Univer chrome without changing canonical workbook data. UI CSS uses VS Code semantic tokens. All 21 locale shards compile with matching keys; Chinese and English labels are available, and other locale shards currently use English fallback.

## Commands and actual results

| Check | Result |
| --- | --- |
| `npm run vscode:build` | Three workspace builds passed; backend rebundled after the cache-invalidation change |
| `npm run vscode:typecheck` and gateway typecheck | Passed |
| `npm test -w @genoffice/workbook-session` | 20 passed; subsequent protection-index regression also passed |
| `npm test -w @genoffice/sheets-editor` | 5 passed |
| `npm run test:e2e -w @genoffice/sheets-editor` | 10 passed |
| `npm test -w apps/vscode-office` | 7 passed, 4 files |
| `test:integration --label minimum --fail-zero` with Xvfb | 3 passed; includes product lifecycle/save/readonly checks |
| Existing Sheets sidecar/streaming-save/structure Vitest files | 105 passed |
| `npm run build -w @genoffice/shell` | Actual main, preload and renderer builds passed |
| `npm run check:theme-colors` | Passed; its repository scope covers existing renderer CSS, while the new Webview CSS was also inspected directly |

Fixture SHA-256: `8cdb24baa6557abf1638ae667ead65639d1fd6b08da69ee4d635e7496a783844`. ZIP timestamps are fixed. The fixture has two sheets, a cross-sheet formula, literal identifier/empty string, numeric format/fill and an unknown OOXML preservation sentinel.

## Implementation locations

Structure/chart intents are emitted by `WorkbookEditor.tsx`; pure execution and reference validation live in gateway `session-candidate.ts` / `session-chart-references.ts`, with host materialization in `candidate-material.ts`. Saving reuses `xlsx-package-io.ts` / `xlsx-gateway.ts`. Electron retains compatible client/schema/helper re-exports. These are the actual locations replacing the proposed structure/chart/save adapter filenames in the task list; no large App.tsx copy or empty DesktopApi was introduced.

## Remaining gates

US4 conflict/alias/binding handling, US5 product unified undo/redo and durable recovery, product native tools, Python exchange, skill installation, platform packaging, current VS Code, Windows/macOS/SSH, real chat and performance acceptance remain unfinished. The separate API-probe tool/undo test is not product tool or product history evidence. Original native charts and advanced workbook objects retain their original package parts; the new editing UI currently manages session-created basic charts. No VSIX, commit, PR or publication was produced.
