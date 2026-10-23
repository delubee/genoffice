# Regression gate — 2026-10-22

Actual local Linux x64 results (Node 22.21.0, Rust/Cargo 1.96.0; native VS Code embeds its own Node):

| Command / gate | Actual result |
| --- | --- |
| `npm run vscode:build` | Three workspaces built; browser boundary checks passed |
| `npm run vscode:typecheck` | Session, editor, extension passed |
| `npm run vscode:test` | Session 46, editor 7, extension 31 tests passed |
| `npm run test -w @genoffice/xlsx-gateway` | 252 passed |
| gateway, sheets, shell typecheck | Passed |
| `npm run test -w @genoffice/sheets` | Rust 272 passed; JS 3,453 passed, 1 skipped |
| `npm run test -w @genoffice/shell` | 1,012 passed outside IPC-restricted sandbox |
| `npm run test -w @genoffice/i18n` | 26 passed |
| editor Playwright E2E | 10 passed |
| Native VS Code minimum 1.109.5 | 8 passed; separate opt-in benchmark skipped |
| Native VS Code current stable 1.143.0 | 8 passed; separate opt-in benchmark skipped |
| Actual Sheets and Shell Electron builds | Main, preload and renderer builds passed |
| Targeted ESLint across new/affected code | Passed |
| `check:theme-colors` | Passed |
| `check:help-langs` | All 21 languages structurally match English |
| `check:english-comments` | Passed |
| `npm run licenses` | npm production license allowlist passed |
| `cargo deny ... check licenses` | Passed |
| `npm run notices` | Generated notices for 572 npm packages and 144 crates; packaged with VSIX |
| `git diff --check` | Passed |

The initial ordinary sandbox Shell run could not bind loopback/IPC and could not prepare Electron; its failures are not recorded as product passes. Re-running the complete suite with the required local IPC/filesystem access passed. Rust license checking likewise needed permission to unpack Cargo's cross-platform dependency cache. No approval bypass is added to product code.

Native tests invoke registered product tools through VS Code's actual LM API using a trusted harness. They do not establish real model behavior or approval dialog rejection. Browser tests use actual Univer; the 100k output capacity test uses actual Node/Rust and is separately documented. The packaged smoke tests use ordinary profiles and installed VSIX files, without extensionDevelopmentPath, and remove Cargo paths from the child environment.

The exact reference performance device, Windows/macOS and SSH are not established by these local regression results. The application intentionally retains English fallback for editor locale shards other than Chinese/English; the 21-language manual gate is not a claim of full editor translation.
