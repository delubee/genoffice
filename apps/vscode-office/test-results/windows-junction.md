# Windows junction path compatibility — 0.1.5

Date: 2026-10-24. Source branch: `fix/windows-rclone-paths`, based on integrated
upstream/feature checkpoint `5f6fe655`.

Opening a workbook previously required `fs.realpath` to succeed before copying
its bytes. Save baseline capture imposed the same requirement. The document-only
resolver now handles Windows `UNKNOWN`, `ENOSYS`, `ENOTSUP` and `EINVAL` from that
operation: it verifies the user path with stat, follows readable junctions with
bounded cycle detection, and checks the resulting accessible path. Missing files,
permission failures and errors on other platforms retain their existing behavior.
Workspace exchange, recovery and skill containment checks keep strict realpath.
Save hash/version checks and atomic publication behavior are unchanged.

## Verification

- Linux: 70 workbook-session, 10 editor and 31 extension tests passed; three
  workspace typechecks and changed-file ESLint passed.
- Windows: the same 111 unit tests and three typechecks passed in an isolated
  source tree; the static-CRT sidecar and host/webview build passed.
- Windows VS Code 1.109.5: seven integration tests passed. The new test creates a
  real local NTFS junction, injects the reported UNKNOWN error for realpath on
  its document paths, then verifies opening, editing, saving and reopening.
  File reads, junction traversal, snapshot copying, sidecar and atomic save are
  real. Six existing appearance/merge/dimensions/settings/Chinese-sort regressions
  also passed.
- The actual VSIX was installed into a fresh ordinary profile. Editing,
  save/reopen, native tool reads and bundled skill installation passed with
  Rust/Cargo removed from the runtime PATH.
- Artifact SHA-256 and installed proof match. The AMD64 native component has
  no VCRUNTIME/MSVCP dependency.

No failing rclone file path was supplied. The test reproduces the filesystem
capability error using a local junction; actual rclone mount behavior remains
unverified. This release handles path resolution and does not add unsupported
filesystem operations or relax workspace path-containment rules.

## Build provenance

Windows DNS could not resolve npm registries. The build used a separate source
copy under `E:\github\genoffice\.genoffice\work\j015-20261024`. Existing dependency
files were copied with workspace junctions excluded; 12 updated public dependency
tarballs were checked against the lockfile SHA-512 and imported from the Linux
npm cache. Offline npm installation succeeded, and all 30 workspace links were
verified to point into the isolated source tree. The original remote source
worktree was retained. Logs are under `.genoffice/work/junction-20261024`.

Artifact: `artifacts/genoffice-workbook-0.1.5-win32-x64.vsix` (6,790,383 bytes).
SHA-256: `9803121b3080084907396277b98d2f237b15b8438e9a82069c75967d0a91680e`.
See [package audit](junction-package-audit.json) and
[installed proof](junction-installed-win32-x64-proof.json).
