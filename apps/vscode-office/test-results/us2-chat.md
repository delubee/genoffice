# US2 — native workbook tools

Verified locally on Linux x64 with VS Code 1.109.5, 2026-10-22. These are actual
registered product tools exercised through `vscode.lm.invokeTool`, not a second
chat implementation or a model-driven conversation.

## Implemented

Five stable native tools: `genoffice_workbook_context`, `_read`, `_aggregate`,
`_apply`, `_operationStatus`. Contributions are generated from the same runtime
Zod schemas; build checks reject stale manifest schemas. Exchange tools are not
registered until their implementations exist.

- Explicit URI or window-local binding; no focus/basename inference. Context with
  no target returns the open choices and requires_target. The bind-selection
  command flushes input, requests the actual selection/state/revision, validates
  them against authority, and copies the frozen binding reference.
- Current immutable reads include unsaved typed values and calculation status.
  Ordinary reads: 2,000 cells; formats: 200; result body: 1 MiB, otherwise an
  explicit error rather than truncation. Aggregates page at most 2,000 cells per
  source request, up to 1,000,000 cells and 50 frequencies. Blank and empty text
  differ; sums include numeric kind, not dates or numeric-looking strings.
  Unverified formula caches are refused for aggregates.
- Apply validates the operation whitelist, complete batch structure/content
  separation and existing DSL limits. Shared host execution independently
  enforces expansion, range, protection, formula and undo budgets. One successful
  batch is one native undo step; it does not save the XLSX.
- Native preparation describes file, worksheets, ranges, operations and side
  effects without materializing or modifying content. Apply requests native
  confirmation; invocation rechecks revision and cancellation. No Python runner,
  Chat Participant, credentials or approval-bypassing fallback was introduced.
- Receipts are queried without opening closed files. Same transaction/payload
  retry returns its original receipt before rejecting its old expectedRevision.
- Structure invalidates a frozen selection even after rereading. An explicitly
  declared new range can use usesFrozenSelection=false only after a fresh read;
  using the original selection requires rebinding.
- Read/context expose projection errors. Hidden or detached views use canonical
  host state; pending input returns editor_busy, never old disk content.

## Evidence

```sh
npm run vscode:typecheck
npm run test -w @genoffice/workbook-session
npm run test -w apps/vscode-office
npm run test:integration -w apps/vscode-office -- --label minimum --fail-zero
npm run test:integration -w apps/vscode-office -- --label minimum --grep 'Product native workbook tools' --fail-zero
npm run test:e2e -w @genoffice/sheets-editor
```

Session suite: 35 passed, including an independently checked 100,000-cell
aggregate with 50 bounded source pages. Extension unit suite: 16 passed. Native
suite: 6 passed, including the new product-tool scenario. Typecheck passed after
correcting the test helper's input type. The browser run passed 9/10; the remaining
fixture incorrectly mixed structure and content, was corrected to separate
batches, and all three candidate browser tests then passed. This failure was a
test contract violation, not an accepted mixed-batch behavior.

Actual native scenario: manually change unsaved A2 from 12.5 to 25; context/read
and aggregate see 25; focus B; apply through A's binding sets A2=40 while B's
state ID stays unchanged; retry leaves revision at 2; status returns committed;
stale revision is rejected; native undo restores 25. A subsequent row insertion
stales the frozen selection; a fresh read does not erase structural staleness.
The final targeted run also verifies a separately declared new range.

## Not verified

No signed-in model/agent chat was available for a real conversation. Model tool
choice, user-facing confirmation rejection, and actual Chat/terminal focus
switching remain manual acceptance items. The tests prove the product's stable
native tool boundary, not that every VS Code agent supports these tools or skills.
Windows/macOS/current VS Code/SSH acceptance remains outstanding.
