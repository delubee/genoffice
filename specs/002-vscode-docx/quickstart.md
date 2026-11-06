# DOCX validation quickstart

## Development gates

Use Node 22 and the repository's installed dependencies. Run new schema/session
unit suites, `npm run vscode:typecheck`, `npm run vscode:test`, affected Docs/UI
regressions and `npm run vscode:build`. Browser tests need a permitted localhost
server; native VS Code tests need a display on Linux. Use the existing Windows
isolated build/cache workflow for Windows packaging.

## Acceptance scenarios

1. Install one VSIX and open one DOCX and one XLSX through their own Open With
   choices. Confirm native workbook tools/settings still work and Word assets
   are not initialized in the XLSX Webview.
2. Edit plain text, headings, fonts/alignment, lists and a basic table. Save,
   close and reopen; inspect actual DOCX XML and displayed values.
3. Alternate at least 20 manual/tool edits and cross a save. Undo/redo each step;
   stale and duplicate tool submissions must have documented outcomes.
4. Edit a plain paragraph beside pictures, header/footer and protected content.
   Check untouched ZIP entry contents and rejected object/attribute mutations.
5. Export unsaved content, run a visible workspace Python transform, preview and
   apply results. Change the document or result file before apply and verify
   stale/hash/path protections.
6. Change theme without resetting selection/history; test external file changes,
   Save As, backup/restore, wrong session messages and close/reload.
7. Run real Windows custom-editor integration and actual installed-VSIX checks
   for both formats. Record performance on the 1 MiB/200 paragraph/10 table fixture.

Record commands, versions, actual results and remaining gates in a feature
verification report. Do not mark real-model or untested platform acceptance
complete based solely on unit/renderer tests.
