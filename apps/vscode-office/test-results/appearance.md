# Workbook appearance — 0.1.1

The default is `excel`: light chrome, Segoe UI/Calibri UI fonts, a green title bar and green selection accents. `vscode` follows the IDE's light/dark/high contrast mode and chrome variables. The toolbar selector persists `genoffice.workbook.appearance` in the document's workspace folder (user settings for standalone files). Settings changes update already-open views without recreating their workbook runtime.

Theme changes use CSS chrome tokens and Univer's theme service. No workbook transaction, source write, save, history move or authored cell/font/fill/chart-color change is performed.

## Actual verification

- TypeScript checks passed for all three workspaces on Linux and Windows.
- Local unit suites: session 46, editor 7, extension 31 passed.
- Real Chromium/Univer: 11 browser tests passed. The new test switches Excel -> VS Code dark -> VS Code light -> VS Code dark -> Excel and verifies unchanged unsaved canonical content. Final title-bar screenshots and targeted rerun passed.
- Windows 11 / VS Code 1.109.5: actual custom editor defaults to Excel, remains light, applies persisted settings immediately and preserves the exact unsaved workbook state. Final native test passed in 7,033 ms.
- Final Windows packaging and clean-profile installation evidence is recorded alongside the artifact audit.

The shared editor defaults retain their prior host behavior; the VS Code adapter explicitly defaults to Excel. Product skill version remains unchanged. The extension version is 0.1.1 so an existing 0.1.0 installation can be upgraded.

Final Windows artifact: `genoffice-workbook-0.1.1-win32-x64.vsix` (6,750,718 bytes), SHA-256 `9292dc1d12580a4af4323c9802c21c67104957f9d4b924419f6d0fe23b37b7db`. The actual clean-profile install proof matches this exact package.
