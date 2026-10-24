# Windows x64 package verification — 2026-10-23

## Build environment

Built through `ssh windows-work` in `E:\github\genoffice` on Windows 11 x64, starting at commit `3b058c934be6398f74cca18b580d69e58c04e867` on `feature/vscode-office`. Node 22.17.0, npm 10.9.2, Rust/Cargo 1.96.0 and existing Python 3.14.4. Dependencies were installed from the committed lockfile with `npm ci --ignore-scripts`; no standalone GenOffice runtime is required.

The artifact additionally contains the working-tree Windows compatibility fixes listed below; it is not a byte-for-byte build of the original commit. These fixes are retained locally and in the remote build checkout, and have not been committed or pushed by this packaging task.

## Corrections found by actual Windows execution

- Recovery copied bases need a writable handle for Windows `FlushFileBuffers`. The private copied base is made writable, flushed through `r+`, closed and made read-only again before acknowledgment. A read-only source regression test covers this.
- VS Code normalizes drive letters differently from filesystem realpath. Recovery slot reuse accepts equivalent native paths while retaining canonical parent/target boundary checks. Listing uses canonical storage. The tests exercise drive-letter case variation on Windows and reject outside paths.
- VS Code removes a clean document's previous backup; the next dirty edit must recreate its stable slot. Missing slots are allowed only at the verified canonical parent, and the save/retire/recreate regression is covered.
- Python output is checked as parsed JSON rather than requiring Unix line endings. The native skill installation test compares canonical filesystem paths.

## Actual verification

| Gate | Result |
| --- | --- |
| Windows three-workspace TypeScript check | Passed |
| Windows session tests | 46 passed |
| Windows browser-editor unit tests | 7 passed |
| Windows extension unit tests | 31 passed |
| Windows Rust engine | 272 passed |
| VS Code 1.109.5 full native integration | 8 passed; separate opt-in benchmark pending |
| Final backup regression | 4 recovery/slot tests passed on Windows and Linux |
| Native Python summary/save/undo after final slot fix | Passed: 81,818 ms; backup error count 0 |
| Installed final VSIX / clean normal profile | See `package-win32-x64-proof.json` |
| Final ZIP/target/PE/DLL/SHA validation | `windows-package-audit.json` |

The installed-package proof runs the actual VSIX without `extensionDevelopmentPath`, edits a workbook, saves/reopens it, reads value 731 through the registered native tool, and installs the bundled workspace skill. Its launcher removes Cargo/rustup paths.

## Deliverables

Remote: `E:\github\genoffice\apps\vscode-office\artifacts\genoffice-workbook-0.1.0-win32-x64.vsix`.

Local: `apps/vscode-office/artifacts/genoffice-workbook-0.1.0-win32-x64.vsix` and the adjacent `.sha256` file. SHA-256: `3223b0b641bb541dd698f027bc5b481f612f50d38d77b16167e230c756a8c7c5`.

The package has a Windows x64 target and a real PE x64 sidecar. Its import table contains Windows system DLLs and no VCRUNTIME/MSVCP/CONCRT redistribution DLL. The sidecar and package were built on the Windows host.

## Remaining acceptance

Current-stable VS Code on Windows, human keyboard/IME/close-dialog coverage, Windows crash-recovery repetition and real model/approval conversations are not established by these tests. macOS and actual Linux Remote SSH acceptance remain pending. This remains a preview, with no Marketplace publication or implicit commit/push.
