# US3 — exchange failure and cancellation evidence

Actual Vitest suites: `packages/workbook-session/tests/exchange-types.test.ts`, `import-contract.test.ts`, `import-plan.test.ts`, and extension `python-analysis.test.ts`, `python-types.test.ts`, `import-cancel.test.ts`. See the named test sources for exact assertions; all are run again in the final regression gate.

Validated boundaries include finite typed values, shape and record coverage, metadata/record/file limits, source registry identity, target authority, manifest/data digests, source revision, immutable staged copy, strict cursor and transaction retry, 30-minute expiration, and cancellation. Symlinked workspace/task paths are rejected before outside directories are created. Changing a public manifest does not create source authority. Whole-workbook read-only and protected targets reject writes.

Pre-commit cancellation adds no empty sheet and changes no content. Post-commit cancellation reports the actual receipt and partial progress; committed history remains undoable after the plan/private staging is removed. Closing a session drops plans and snapshot authority while retaining canonical history. User workspace scripts and result files remain untouched. The full-output proof records cancellation after batch three and all 100 final batch receipts separately.

Missing Python, missing dependency, script failure and process cancellation are exercised by a harness. Product tools do not install Python or dependencies or provide an alternate process execution channel. Native terminal approval and real model behavior remain unverified.
