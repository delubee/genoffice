# US4 bindings, conflicts and failure isolation — 2026-10-22

Verified on the current Linux x64 host only. Windows/macOS/current VS Code/SSH verification remains in the packaging and acceptance phases.

## Authority and binding checks

`WorkbookBindings` belongs to one extension-host instance. Bindings retain explicit URI/session/logical-sheet/range identities and never follow focus. Content commits stale prior authorization. A fresh read can authorize the current revision; structural selection staleness remains until rebinding. Explicit new ranges can use the current revision after a fresh read. Close/revert invalidate the old binding; recovery creates a new session rather than migrating a token.

`freshness.test.ts` verifies A→B→A content edits retain monotonically different revisions; an immutable earlier read root remains unchanged. It also verifies structural rebinding, independent same-name documents, separate binding registries and closed/reopened epochs. `assertFreshRevision` is used before preparation, with commit-time root/revision validation retained. Save, selection and indexing do not create content revisions.

## Disk publication and aliases

`captureFileBaseline` surrounds hashing with physical-ID/size/mtime/ctime checks. Session publication validates the stored baseline, retains independent previous/candidate copies, rechecks immediately before replacement and checks the written identity/hash afterwards. File copies use separate inodes so an external in-place write cannot corrupt a retained candidate too. A new target uses exclusive link publication; an existing target uses rename with bounded retries. A locked target returns a failure with preserved files. External processes do not participate in a shared CAS: the checks and watchers detect tested races, but no universal guarantee against arbitrary external writes after the final check is claimed.

`conflict-faults.test.ts` uses real filesystem files to verify same-byte external replacement, changes before publication, and interference after publication. The last case keeps the external file and both previous/candidate copies, reporting `published:true`. Real XLSX stage/save and failure tests remain in the US1 suite.

`WorkbookRegistry` reserves canonical paths and physical identities before opening, including pending opens. Hard-link aliases cannot obtain another writable session in this window; different same-name files are independent. Save updates physical ownership and revert replaces the session epoch. Closing releases only the corresponding owner. File watchers monitor requested and canonical paths. Own saves pause checks and install the new baseline before resuming; external changes preserve current state and provide an explicitly targeted Save As command.

## Receipts and native proof

`receipts.ts` centralizes payload fingerprints, terminal record validation, idempotent retry and live-session operation-status lookup. The same transaction/payload returns its committed receipt before checking its original expected revision; a different payload conflicts. Status lookup never opens a closed document. Existing transaction tests retain precommit root-publication faults, cancellation, postcommit projection failure and lost-response retry coverage.

Actual checks:

- `npm test -w @genoffice/workbook-session`: 26 passed, 13 files.
- Extension unit tests after publisher wiring: 10 passed, 5 files, including three real publication race scenarios.
- `npm run vscode:typecheck`: passed.
- Actual VS Code 1.109.5 Extension Host with Xvfb: **4 passed**. The new product test verifies binding staleness, receipt retry/query, own-save watcher suppression, revert/new epoch, closed-session status rejection, actual hard-link alias rejection, external file notification, blocked normal save and explicit reload. Alias rejection emits its expected VS Code diagnostic.

Product unified history and durable save/recovery are the next phase. Product LM tools, Python and skills remain unfinished. No real chat, Windows/macOS/SSH, performance or release acceptance is implied.
