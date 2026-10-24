# G0 design proof — 2026-10-21

## Executed environment

Linux x64 (this server), Node 22.21.0, Univer 0.25.1, Chromium 151.0.7922.34, VS Code 1.109.5, Xvfb display :93. This is a feasibility gate, not product acceptance or a Windows/macOS/SSH result.

## Evidence

- `npm run test:e2e -w @genoffice/sheets-editor`: 1 passed. The test runs actual Univer in Chromium, creates separate main and candidate runtimes, edits the first candidate operation then throws at operation 2, and checks the main snapshot is unchanged. A real `beforeCommandExecuted` listener vetoes manual value mutation before application. Successful row insertion shifts the original cell; chart metadata from the existing gateway and the workbook snapshot survive a JSON round trip; installing the inverse snapshot restores the initial state. All four created runtimes are disposed, with zero remaining in the test counter.
- `DISPLAY=:93 LD_LIBRARY_PATH=/tmp/genoffice-xvfb/root/usr/lib64 npm run test:integration -w apps/vscode-office -- --label minimum --fail-zero`: 2 passed. A separate test-only extension contributes a custom editor and an LM tool. VS Code invokes the registered tool; preparation does not edit; invoke rejects a preparation made stale by another edit. Actual save, undo after save and redo commands call the provider and preserve document values. Undo/redo do not emit duplicate edit events. The rejection diagnostic in the host log is expected.
- `npm run test -w @genoffice/workbook-session`: 4 passed, including `g0-package-proof.test.ts`. The existing OOXML gateway serializes a row insertion and a new chart into one private package, writes and reads that file, verifies the shifted coordinate and native chart XML, then writes the inverse identity root from the immutable original bytes. The original part inventory and SHA-256 digests are restored. Replaying the serialized forward plan produces identical part digests. Unrelated parts pass the gateway preservation check.
- `npm run vscode:typecheck`: passed after the spike type corrections.

## Findings and production constraints

The proven whitelist is value interception, row insertion, column-chart metadata, OOXML row/chart serialization, and stable CustomEditor/LM lifecycle APIs. This does not prove the full production command whitelist, formula calculation, protected structures, chart UI, large-file performance, durable transactions or recovery.

Univer headless structure commands still require locale initialization. Its command listener may veto by throwing; production interception must recognize explicit user commands, suppress candidate/projection commands, flush IME and forward a typed intent to the host. A successful isolated mutation alone is not a canonical commit.

The legacy gateway adds `fullCalcOnLoad` even to an empty mutation plan. An identity overlay must therefore serialize original bytes directly to preserve every original part. Nonidentity save must use the original base and complete overlay, never the last saved file. The G0 inverse test takes this identity path; it does not claim that applying the opposite structural command alone reconstructs deleted content.

Existing renderer chart history uses closures and is unsuitable for durable undo. The spike carries charts as JSON feature state separately from Univer snapshots. Production changesets must retain all affected feature maps and structural/reference state before committing. Existing executor partial failure is allowed only within a disposable candidate.

No proposed API, global chat-send interception, model credentials, or independent Python runtime was used. The product extension currently only activates; the test editor/tool are confined to the test fixture and are not product functionality.

## Gate

G0 feasibility passed for proceeding to T012–T022. G1 and all user-story acceptance gates remain incomplete.
