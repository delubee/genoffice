# US5 — unified history and recovery

Local verification: 2026-10-22; Linux x64, Node 22.21.0, VS Code 1.109.5,
Chromium 151. This is development evidence, not cross-platform release acceptance.

## Verified behavior

- Mixed manual/agent transactions produce one native edit event per batch. Undo/redo
  restores the earlier state ID while increasing revision. New edits truncate redo;
  saving retains the original base, full overlay, structural log and timeline.
- Structural/chart undo across multiple saves serializes the original identity
  exactly. Native dirty state becomes clean when redo returns to the saved root.
- View toolbar history uses the same native history. Requests for a background
  workbook are rejected: VS Code's global undo command does not honor a URI target.
  This prevents undo from affecting the focused workbook accidentally.
- Complete base/state/feature backups are written before their manifests are
  published. Interrupted replacement preserves the previous acknowledged copy.
  Recovery creates a new session and restores content, not the old undo timeline.
- Stable backup slots are necessary: after SIGKILL, VS Code can provide its older
  serialized backup ID. The slot atomically points to the newest complete copy;
  deleting an obsolete generation cannot delete its replacement.
- Before-write cancellation, pending cell input, a read-only target, injected
  ENOSPC during root staging and failed saved-marker publication retain content
  and history. A file-written/marker-failed result explicitly reports fileWritten.
  A detached view can save canonical committed content.
- Actual serialized forward/inverse byte budgets are checked before commit.
  Saving never silently evicts undo entries to make room.

## Commands and results

```sh
npm run vscode:typecheck
npm run test -w @genoffice/workbook-session
npm run test -w apps/vscode-office
npm run test:e2e -w @genoffice/sheets-editor
npm run test:integration -w apps/vscode-office -- --label minimum --fail-zero
npm run test:recovery -w apps/vscode-office
npm run test:recovery -w apps/vscode-office -- --fresh-crash
```

Browser suite: 10 passed. Native integration suite: 5 passed. The native recovery
proof installs copies into a private extensions directory and uses an isolated
profile/workspace. It runs without extensionDevelopmentPath: development windows
skip part of VS Code's native backup registration and cannot prove this behavior.

The ordinary-window proof first hot-exits with unsaved A2=99, restores it, changes
A2 to 101, waits for acknowledgment, SIGKILLs only its own test application, then
restarts. The recovered value is 101, with a new process/session and empty undo
timeline. The original XLSX SHA-256 remains
`8cdb24baa6557abf1638ae667ead65639d1fd6b08da69ee4d635e7496a783844`
until explicit native Save. Save clears both dirty markers and Close disposes the
document. Machine-readable evidence: [recovery-proof.json](recovery-proof.json).

The fresh-window proof SIGKILLs the first window after acknowledgment, without a
previous normal shutdown. VS Code did not automatically reopen this first-ever
editor in the tested environment. The **GenOffice: Recover Unsaved Workbook**
command restores its acknowledged copy and correctly marks the native document
dirty. Save/Close then succeed. Evidence:
[recovery-fresh-proof.json](recovery-fresh-proof.json). Recovery is restricted to
the current workspace's extension storage and never overwrites an already open
workbook. Full hashes and original URI are checked on selection.

## Limits and pending acceptance

- Automatic restoration of a first-ever crashed window is not guaranteed; use the
  explicit recovery command if VS Code opens an empty window.
- Backup acknowledgment is the durability boundary; edits newer than the last
  acknowledged backup are not promised. Cross-process undo history is not restored.
- Windows, macOS, current VS Code, Remote SSH, actual full-volume exhaustion and
  interactive Save As/close-dialog cancellation remain unverified here. ENOSPC
  above is controlled fault injection, not a filled system disk.
- Original advanced native chart editing and the other release gates remain outside
  this local history/recovery verification. The overall feature remains incomplete.
