# Acceptance status — 2026-10-22

## Local installed artifact

`artifacts/genoffice-workbook-0.1.0-linux-x64.vsix` was installed with the official VS Code CLI in separate ordinary, clean profiles on **1.109.5** and **1.143.0**. The product is loaded from the installed extension directory, with no extensionDevelopmentPath and no standalone GenOffice app. The test launcher removes Cargo/rustup paths. Actual editing, save/reopen, registered native tool read of saved value 731, and workspace skill installation pass.

Evidence: `package-linux-x64-minimum-proof.json`, `package-linux-x64-stable-proof.json`, and `artifacts/package-proof.json` plus VSIX SHA-256 sidecar. A separate small test helper is installed through the CLI; it is not included in the product VSIX. The first smoke attempt timed out because a helper copied into the extension directory was absent from VS Code's installed-extension registry; installing the helper properly as its own VSIX corrected the harness.

## Quickstart coverage

| Scenario | Local evidence | Remaining verification |
| --- | --- | --- |
| V1 manual editing/history | Browser E2E, native product editor/history, installed smoke, recovery proofs | Full human UI close/save/discard/cancel matrix on all target OSs |
| V2 native chat/tools | Eight native product tools; bindings, current unsaved reads, edits and focus isolation | Logged-in model conversations, actual tool-confirmation rejection |
| V3 100k Python summary | Native tools, actual standard-library script, independent regional totals, chart, save/reopen, typed source digest | Native agent terminal approval with a real model |
| V4 complete result/failures | Actual 100k output / 100 batches, cancellation after three, private-copy mutation, typed import tests | Other target OSs |
| V5 conflicts/epochs/history | Native and deterministic tests, alias ownership, stale plans, disk conflict, save failure, hot-exit/SIGKILL | Target OS/SSH repetition |
| V6 skill/degradation | Actual native install/status and modification detection; pure keep/update/backup and missing-tool tests | Real agent skill discovery and read/edit/Python conversations |

## Unavailable environments

Windows 11 x64, macOS 14 arm64 and an actual Ubuntu 22.04 x64 SSH extension host were not available. No platform/remote pass is inferred from header fixtures or local host-policy tests. CI has the required platform labels and minimum/stable commands, but has not been remotely executed. No Marketplace publication, commit or PR is part of this work.
