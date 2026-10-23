# Setup verification

Date: 2026-10-21. Tasks T001–T006 only; document editing has not been implemented.

- Node v22.21.0, npm 11.14.1, rustc/cargo 1.96.0.
- `npm install --ignore-scripts`: successful. No previously locked package version changed. Electron install scripts were not run.
- `npm run vscode:build`: all three new workspaces built successfully.
- `npm run vscode:typecheck`: all three new workspaces passed.
- Actual VS Code 1.109.5 downloaded by @vscode/test-cli. `test:integration --label minimum --fail-zero`: one host activation test passed, exit code 0.
- Playwright Chromium 151.0.7922.34 launched and rendered a page successfully outside the process-restricted execution sandbox.
- Stable/current VS Code, actual workbook operations, native chat, Python integration, Windows, macOS, Remote SSH and performance gates have not yet been verified.

The extension uses manifest name `genoffice-workbook`; npm scripts address it as `-w apps/vscode-office` because a VS Code extension name cannot be an npm scoped name. Shared libraries retain their scoped workspace names.

The server had no display. Xvfb and its shared libraries were extracted under `/tmp/genoffice-xvfb`; the two packages xorg-x11-xkb-utils and libxkbfile were installed for its required keyboard compiler. The test display runs with TCP listening disabled. These are development environment dependencies, not extension runtime dependencies.

The test host and browser required execution outside the session sandbox because that sandbox denies their local IPC. This is distinct from a failed extension test; the actual host activation result above was obtained after supplying the required test environment.
